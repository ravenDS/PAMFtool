' PamfHeaderWriter.vb - github.com/ravenDS/PAMFtool
'
' Build PAMF header (one or more 2048-byte sectors)
'
' Fields:
'   +0x08  u32  header sectors                           (2 when the EP map does not fit one sector)
'   +0x0C  u32  data sectors
'   +0x50  u32  size A                                   = 0x64 + 0x30*(n-1)
'   +0x56  u48  start_pts                                (90 kHz, 90000)
'   +0x5C  u48  end_pts                                  (first video PTS + video duration)
'   +0x62  u32  mux_rate_bound                           (units of 50 bytes/sec)
'   +0x66  u32  std_delay_bound                          (90 kHz)
'   +0x6D  u8   numStreams
'   +0x70  u32  size B                                   = 0x44 + 0x30*(n-1)
'   +0x74  u48  start_pts (duplicate)
'   +0x7A  u48  end_pts (duplicate)
'   +0x82  u32  size C                                   = 0x32 + 0x30*(n-1)
'   +0x87  u8   numStreams (duplicate of +0x6D)
'   +0x88  ...  stream entries (48 bytes each) and codec info
'   +...   ...  EP tables, one per video stream, 12 bytes per entry

Namespace PamfMux

    Public Class PamfHeaderWriter

        Public Const HeaderSize As Integer = 2048
        Public Const StreamEntrySize As Integer = 48

        ' 136-byte template covering 0x00..0x87
        Private Shared ReadOnly TemplateBytes As Byte() = New Byte() {
            &H50, &H41, &H4D, &H46, &H30, &H30, &H34, &H31,  ' 0x00: "PAMF0041"
            &H0, &H0, &H0, &H1, &H0, &H0, &H0, &H0,          ' 0x08: ver=1, numPacks=0
            &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,          ' 0x10
            &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,          ' 0x18
            &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,          ' 0x20
            &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,          ' 0x28
            &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,          ' 0x30
            &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,          ' 0x38
            &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,          ' 0x40
            &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,          ' 0x48
            &H0, &H0, &H0, &H64, &H0, &H0, &H0, &H0,         ' 0x50: size A (n=1) + zeros
            &H0, &H1, &H5F, &H90, &H0, &H0, &H0, &H0,        ' 0x58: 90000, dur high zeros
            &H0, &H0, &H0, &H1, &H0, &H0, &H0, &H1,          ' 0x60: dur tail + const 0001, mux_rate placeholder + const 0001
            &H5F, &H90, &H0, &H0, &H0, &H1, &H0, &H1,        ' 0x68: const, n=1 at 0x6D, const 0001
            &H0, &H0, &H0, &H44, &H0, &H0, &H0, &H1,         ' 0x70: size B (n=1), const
            &H5F, &H90, &H0, &H0, &H0, &H0, &H0, &H0,        ' 0x78: const, duration2 placeholder
            &H0, &H1, &H0, &H0, &H0, &H32, &H0, &H1          ' 0x80: const, size C (n=1), n=1
        }

        Private _streams As New List(Of StreamEntry)()

        Private Class StreamEntry
            Public StreamTypeByte As Byte
            Public Channel As Byte
            Public PesStreamId As Byte
            Public SubStreamId As Byte
            Public PStdBufferRaw As UShort     ' packed P_STD buffer field for entry+6..7
            Public CodecInfo As Byte()
            Public EpEntries As List(Of PamfMux.EpEntry)
        End Class

        ' Pack P_STD buffer: 2 unused bits, 1 scale bit, 13 size bits
        ' scale=0 means size is in 128-byte units, scale=1 means 1024-byte units
        Private Shared Function PackPStdBuffer(scale As Integer, size As Integer) As UShort
            Return CUShort(((scale And 1) << 13) Or (size And &H1FFF))
        End Function

        ' typical P_STD values observed in real Sony PAMF encodes:
        '   AVC 720p / 1080i : scale=1, size ~1505 (~1.5 MB)
        '   M2V 720p         : scale=1, size 264..1234 depending on encoder; Sony's own is ~546
        '   AT3+ / AC-3      : scale=1, size 20 (20 KB is plenty for ~256 kbps compressed audio)
        '   LPCM             : scale=1, size 128 (~128 KB) LPCM at 48 kHz stereo 16-bit is ~1.5 Mbps sustained
        ' we use conservative defaults per codec
        Private Const PStdBufferAvc1080 As UShort = CUShort((1 << 13) Or 1505)   ' 0x25E1
        Private Const PStdBufferM2v As UShort = CUShort((1 << 13) Or 546)        ' 0x2222
        Private Const PStdBufferAudio As UShort = CUShort((1 << 13) Or 20)       ' 0x2014
        Private Const PStdBufferLpcm As UShort = CUShort((1 << 13) Or 128)       ' 0x2080

        ' P-STD buffer size (in 1024-byte units) for a given AVC level
        ' too small and complex frames stall the decoder mid-stream, too large is harmless surplus
        '
        ' values below come from observing reference PAMFs:
        '   level 3.1 CAVLC (720p 30fps)  -> 1505 KB
        '   level 4.1 CABAC (1080p 30)    -> 2482 KB
        '   level 4.1 CABAC (1080p 24)    -> 3703 KB
        ' Sony varies the L4.1 value based on peak-frame complexity per file, and we can't infer that from the SPS alone
        ' we pick the LARGER Sony-observed L4.1 value (3703) as the default so both L4.1 files decode without underruns
        Public Shared Function AvcPstdBufferRawForLevel(levelIdc As Byte) As UShort
            Dim kb As Integer
            Select Case CInt(levelIdc)
                Case Is <= 31 : kb = 1505    '    level 3.1
                Case 32 : kb = 1800          '    level 3.2
                Case 40 : kb = 2000          '    level 4.0
                Case 41 : kb = 3703          '    level 4.1
                Case 42 : kb = 4500          '    level 4.2
                Case Is <= 50 : kb = 6000    '    level 5.0
                Case Else : kb = 8000        '    level 5.1 and above
            End Select
            Return CUShort((1 << 13) Or (kb And &H1FFF))
        End Function

        Public Shared Function AvcPstdKbForLevel(levelIdc As Byte) As Integer
            Return CInt(AvcPstdBufferRawForLevel(levelIdc)) And &H1FFF
        End Function

        Public Sub AddAvcStream(channel As Byte, pesStreamId As Byte,
                                profileIdc As Byte, levelIdc As Byte,
                                frameMbsOnlyFlag As Byte, videoSignalInfoFlag As Byte,
                                frameRateCode As Byte, aspectRatioIdc As Byte,
                                widthMbs As Integer, heightMbs As Integer,
                                Optional sarWidth As Integer = 0,
                                Optional sarHeight As Integer = 0,
                                Optional cropLeft As Integer = 0,
                                Optional cropRight As Integer = 0,
                                Optional cropTop As Integer = 0,
                                Optional cropBottom As Integer = 0,
                                Optional videoFormat As Byte = 5,
                                Optional videoFullRangeFlag As Byte = 0,
                                Optional colourPrimaries As Byte = 1,
                                Optional transferCharacteristics As Byte = 1,
                                Optional matrixCoefficients As Byte = 1,
                                Optional cabacFlag As Byte = 0,
                                Optional deblockingFilterFlag As Byte = 0,
                                Optional minNumSlicePerPictureIdc As Byte = 3,
                                Optional nfwIdc As Byte = 0,
                                Optional maxMeanBitrate As Byte = 0)
            ' AVC info layout
            Dim ci(31) As Byte
            ci(0) = profileIdc
            ci(1) = levelIdc
            ci(2) = CByte((CInt(frameMbsOnlyFlag And 1) << 7) Or
                          (CInt(videoSignalInfoFlag And 1) << 6) Or
                          ((CInt(frameRateCode) + 1) And &HF))
            ci(3) = aspectRatioIdc
            If aspectRatioIdc = &HFF Then
                WriteU16BE(ci, 4, CUShort(sarWidth And &HFFFF))
                WriteU16BE(ci, 6, CUShort(sarHeight And &HFFFF))
            End If
            ci(9) = CByte(widthMbs And &HFF)
            ci(11) = CByte(heightMbs And &HFF)
            WriteU16BE(ci, 12, CUShort(cropLeft And &HFFFF))
            WriteU16BE(ci, 14, CUShort(cropRight And &HFFFF))
            WriteU16BE(ci, 16, CUShort(cropTop And &HFFFF))
            WriteU16BE(ci, 18, CUShort(cropBottom And &HFFFF))
            ci(20) = CByte(((videoFormat And 7) << 5) Or ((videoFullRangeFlag And 1) << 4))
            ci(21) = colourPrimaries
            ci(22) = transferCharacteristics
            ci(23) = matrixCoefficients
            ci(24) = CByte(((cabacFlag And 1) << 7) Or
                           ((deblockingFilterFlag And 1) << 6) Or
                           ((minNumSlicePerPictureIdc And 3) << 4) Or
                           (nfwIdc And 3))
            ci(25) = maxMeanBitrate
            _streams.Add(New StreamEntry() With {
                .StreamTypeByte = &H1B, .Channel = channel,
                .PesStreamId = pesStreamId, .SubStreamId = 0,
                .PStdBufferRaw = AvcPstdBufferRawForLevel(levelIdc),
                .CodecInfo = ci
            })
        End Sub

        ' override the P-STD buffer size on the LAST-added AVC stream
        ' used when the caller wants to match a reference file's exact buffer
        ' no-op if the last stream is not AVC
        Public Sub OverrideLastAvcPstd(kb As Integer)
            For i As Integer = _streams.Count - 1 To 0 Step -1
                If _streams(i).StreamTypeByte = &H1B Then
                    _streams(i).PStdBufferRaw = CUShort((1 << 13) Or (kb And &H1FFF))
                    Return
                End If
            Next
        End Sub

        ' override the max_mean_bitrate byte (codec_info[25]) on the last-added AVC stream
        ' Sony encodes this per-file (11 for 1080p L4.1 CABAC, 5 for 720p L3.1 CAVLC) and some games may inspect it
        Public Sub OverrideLastAvcMaxMeanBitrate(v As Byte)
            For i As Integer = _streams.Count - 1 To 0 Step -1
                If _streams(i).StreamTypeByte = &H1B Then
                    _streams(i).CodecInfo(25) = v
                    Return
                End If
            Next
        End Sub

        Public Sub AddM2vStream(channel As Byte, pesStreamId As Byte,
                                profileAndLevel As Byte, progressiveSeq As Byte,
                                videoSignalInfoFlag As Byte, frameRateCode As Byte,
                                aspectRatioIdc As Byte,
                                widthMbs As Integer, heightMbs As Integer,
                                widthPx As Integer, heightPx As Integer,
                                Optional colourPrimaries As Byte = 1,
                                Optional transferCharacteristics As Byte = 1,
                                Optional matrixCoefficients As Byte = 1)
            '  M2V info layout:
            '   ci(0)  profileAndLevel (raw MPEG-2 byte: 0x44 = MP@HL)
            '   ci(1)  unused
            '   ci(2)  packed: bit7=progressiveSequence, bit6=videoSignalInfoFlag, bits3:0=frameRateInfo (no +1 offset like AVC)
            '   ci(3)  aspectRatioIdc
            '   ci(4..5) sarWidth
            '   ci(6..7) sarHeight
            '   ci(8)  reserved1
            '   ci(9)  horizontalSize / 16
            '   ci(10) reserved2
            '   ci(11) verticalSize / 16
            '   ci(12..13) horizontalSizeValue (full pixel width)
            '   ci(14..15) verticalSizeValue   (full pixel height)
            '   ci(16..19) reserved (=0)
            '   ci(20) packed: bits7:5=videoFormat (often hardcoded 5 here), bit4=videoFullRangeFlag
            '   ci(21..23) colour_primaries / transfer_characteristics / matrix_coefficients
            Dim ci(31) As Byte
            ci(0) = profileAndLevel
            ci(2) = CByte((CInt(progressiveSeq And 1) << 7) Or
                          (CInt(videoSignalInfoFlag And 1) << 6) Or
                          (CInt(frameRateCode) And &HF))
            ci(3) = aspectRatioIdc
            ci(9) = CByte(widthMbs And &HFF)
            ci(11) = CByte(heightMbs And &HFF)
            WriteU16BE(ci, 12, CUShort(widthPx And &HFFFF))
            WriteU16BE(ci, 14, CUShort(heightPx And &HFFFF))
            ci(20) = &HA0
            ci(21) = colourPrimaries
            ci(22) = transferCharacteristics
            ci(23) = matrixCoefficients
            _streams.Add(New StreamEntry() With {
                .StreamTypeByte = &H2, .Channel = channel,
                .PesStreamId = pesStreamId, .SubStreamId = 0,
                .PStdBufferRaw = PStdBufferM2v,
                .CodecInfo = ci
            })
        End Sub

        Public Sub AddAtrac3plusStream(channel As Byte, subStreamId As Byte,
                                       numChannels As Byte, samplingFreqCode As Byte)
            ' audio codec_info layout:
            '   ci(0..1) unknown u16 BE (=0)
            '   ci(2)    channels
            '   ci(3)    freq code (1 = 48 kHz)
            '   ci(4)    bps (LPCM only)
            Dim ci(31) As Byte
            ci(2) = numChannels
            ci(3) = samplingFreqCode
            _streams.Add(New StreamEntry() With {
                .StreamTypeByte = &HDC, .Channel = channel,
                .PesStreamId = &HBD, .SubStreamId = subStreamId,
                .PStdBufferRaw = PStdBufferAudio,
                .CodecInfo = ci
            })
        End Sub

        Public Sub AddAc3Stream(channel As Byte, subStreamId As Byte,
                                numChannels As Byte, samplingFreqCode As Byte)
            Dim ci(31) As Byte
            ci(2) = numChannels
            ci(3) = samplingFreqCode
            _streams.Add(New StreamEntry() With {
                .StreamTypeByte = &H81, .Channel = channel,
                .PesStreamId = &HBD, .SubStreamId = subStreamId,
                .PStdBufferRaw = PStdBufferAudio,
                .CodecInfo = ci
            })
        End Sub

        Public Sub AddLpcmStream(channel As Byte, subStreamId As Byte,
                                 sampleRate As Integer,
                                 numChannels As Byte, bitsPerSample As Integer)
            ' for LPCM the audio struct also carries a bit-depth code at ci(4).
            ' Sony PAMF use a coded byte here, not the raw bit count:
            '   0x40 = 16-bit  (observed in PS3 game PAMFs)
            '   0x50 = 24-bit  (inferred from the bit-field shape)
            Dim bpsCode As Byte
            Select Case bitsPerSample
                Case 16 : bpsCode = &H40
                Case 24 : bpsCode = &H50
                Case Else : bpsCode = CByte(bitsPerSample And &HFF)  ' fallback
            End Select
            Dim ci(31) As Byte
            ci(2) = numChannels
            ci(3) = &H1   ' 48 kHz
            ci(4) = bpsCode
            _streams.Add(New StreamEntry() With {
                .StreamTypeByte = &H80, .Channel = channel,
                .PesStreamId = &HBD, .SubStreamId = subStreamId,
                .PStdBufferRaw = PStdBufferLpcm,
                .CodecInfo = ci
            })
        End Sub

        ' P-STD buffer size (1024-byte units) of stream i, as finally used by the muxer
        Public Sub SetPstdKb(streamIndex As Integer, kb As Integer)
            _streams(streamIndex).PStdBufferRaw = PackPStdBuffer(1, Math.Max(1, Math.Min(kb, &H1FFF)))
        End Sub

        ' EP map of stream i (video streams only), Nothing/empty = no EP table for that stream
        Public Sub SetEpEntries(streamIndex As Integer, entries As List(Of PamfMux.EpEntry))
            _streams(streamIndex).EpEntries = entries
        End Sub

        Public ReadOnly Property StreamCount As Integer
            Get
                Return _streams.Count
            End Get
        End Property

        ' bytes used by the header before rounding to whole sectors
        Private Function UsedBytes() As Integer
            Dim used As Integer = &H88 + _streams.Count * StreamEntrySize
            For Each s In _streams
                If s.EpEntries IsNot Nothing Then used += 12 * s.EpEntries.Count
            Next
            Return used
        End Function

        Public ReadOnly Property HeaderSectors As Integer
            Get
                Return (UsedBytes() + HeaderSize - 1) \ HeaderSize
            End Get
        End Property

        ' muxRateUnits is the value to write at offset 0x62
        ' in MPEG-2 PS mux_rate convention (50 bytes/sec per unit) for a 24 Mbps stream thats 24_000_000 / 8 / 50 = 60000 = 0xEA60
        Public Function Build(dataSectors As Long, startPts90 As Long, endPts90 As Long,
                              muxRateUnits As Integer, stdDelayBoundTicks As Long) As Byte()
            Dim sectors As Integer = HeaderSectors
            Dim header(sectors * HeaderSize - 1) As Byte
            Array.Copy(TemplateBytes, header, TemplateBytes.Length)
            Dim n As Integer = _streams.Count
            WriteU32BE(header, &H8, CUInt(sectors))
            WriteU32BE(header, &HC, CUInt(dataSectors))
            WriteU32BE(header, &H50, CUInt(&H64 + &H30 * (n - 1)))
            WriteU32BE(header, &H70, CUInt(&H44 + &H30 * (n - 1)))
            WriteU48BE(header, &H56, startPts90)
            WriteU48BE(header, &H5C, endPts90)
            WriteU48BE(header, &H74, startPts90)
            WriteU48BE(header, &H7A, endPts90)
            ' mux_rate_bound is u32 at 0x62
            WriteU32BE(header, &H62, CUInt(muxRateUnits And &HFFFFFFFFL))
            ' std_delay_bound is u32 at 0x66, in 1/90000 second units (1.00 s = 90000 unless -std-delay)
            WriteU32BE(header, &H66, CUInt(stdDelayBoundTicks And &HFFFFFFFFL))
            header(&H6D) = CByte(n)
            WriteU16BE(header, &H84, CUShort(&H32 + &H30 * (n - 1)))
            header(&H87) = CByte(n)
            BuildStreamEntriesAndEp(header)
            Return header
        End Function

        Private Sub BuildStreamEntriesAndEp(header As Byte())
            ' PamfStreamHeader layout:
            '   +0      stream_coding_type
            '   +1..+3  reserved (zeros)
            '   +4      stream_id (e.g. 0xE0 video, 0xBD private)
            '   +5      private_stream_id
            '   +6..+7  p_std_buffer (u16 BE, packed: 2 unused + 1 scale + 13 size)
            '   +8..+11 ep_offset  (u32 BE, byte offset of the stream's EP table from header start, 0 if none)
            '   +12..+15 ep_num    (u32 BE, count of EP entries)
            '   +16..+47 codec-specific info (32 bytes)
            '
            ' EP entry (12 bytes):
            '   value0 (u16 BE) : bits 15:14 = indexN - 1 (reference pictures to decode from this entry, 1..4)
            '                     bit  13    = 0
            '                     bits 12:0  = packs from the RAP pack to the pack holding the last byte of the
            '                                  indexN-th reference picture (nThRefPictureOffset = (v + 1) * 2048)
            '   pts (u48 BE)
            '   rpnOffset (u32 BE) : pack index of the RAP pack (2048-byte units from start of data area)
            Dim n As Integer = _streams.Count
            Dim epAt As Integer = &H88 + n * StreamEntrySize
            For i As Integer = 0 To n - 1
                Dim entryOff As Integer = &H88 + i * StreamEntrySize
                Dim s As StreamEntry = _streams(i)
                header(entryOff + 0) = s.StreamTypeByte
                header(entryOff + 4) = s.PesStreamId
                header(entryOff + 5) = s.SubStreamId
                WriteU16BE(header, entryOff + 6, s.PStdBufferRaw)
                Array.Copy(s.CodecInfo, 0, header, entryOff + 16, 32)
                If s.EpEntries Is Nothing OrElse s.EpEntries.Count = 0 Then Continue For
                WriteU32BE(header, entryOff + 8, CUInt(epAt))
                WriteU32BE(header, entryOff + 12, CUInt(s.EpEntries.Count))
                For Each e In s.EpEntries
                    Dim idxN As Integer = Math.Max(1, Math.Min(4, e.IndexN))
                    Dim raw As Integer = Math.Max(0, Math.Min(&H1FFF, e.RefPictureOffset))
                    WriteU16BE(header, epAt + 0, CUShort(((idxN - 1) << 14) Or raw))
                    WriteU48BE(header, epAt + 2, e.Pts)
                    WriteU32BE(header, epAt + 8, CUInt(e.PackIndex And &HFFFFFFFFL))
                    epAt += 12
                Next
            Next
        End Sub

        Private Shared Sub WriteU48BE(buf As Byte(), off As Integer, v As Long)
            WriteU16BE(buf, off, CUShort((v >> 32) And &HFFFFL))
            WriteU32BE(buf, off + 2, CUInt(v And &HFFFFFFFFL))
        End Sub

        Private Shared Sub WriteU32BE(buf As Byte(), off As Integer, v As UInteger)
            buf(off + 0) = CByte((v >> 24) And &HFF)
            buf(off + 1) = CByte((v >> 16) And &HFF)
            buf(off + 2) = CByte((v >> 8) And &HFF)
            buf(off + 3) = CByte(v And &HFF)
        End Sub

        Private Shared Sub WriteU16BE(buf As Byte(), off As Integer, v As UShort)
            buf(off + 0) = CByte((CInt(v) >> 8) And &HFF)
            buf(off + 1) = CByte(CInt(v) And &HFF)
        End Sub

    End Class

End Namespace