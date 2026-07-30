'  PamfExtractor.vb - github.com/ravenDS/PAMFtool
'  
'  Read a PlayStation Advance Movie Format (.PAMF / .PAM) file, parse the SCE header, and demultiplex
'
'  Outputs depend on whats inside the container:
'      Video : .h264 (AVC, Annex-B) or .m2v (MPEG-2 Video ES)
'      Audio : .at3p (ATRAC3plus raw frames) / .ac3 / .lpcm (raw PCM, BE)

'  PAMF layout:
'      0x000 .. 0x007: "PAMF0041"
'      0x008 .. 0x00B: header_version (BE u32, =1)
'      0x00C .. 0x00F: header_data_size (BE u32, bytes after fixed prefix)
'      0x054 .. 0x067: muxing/timing block (mux rate, 90 kHz clock unit ...)
'      0x086 .. 0x087: numStreams (BE u16)
'      0x088 ..      : stream entry table, 48 bytes per entry
'                        +0x00 streamType (1)   - PAMF type code
'                        +0x01 channel    (1)   - for audio sub-id
'                        +0x02 esFilterId (16)  - filterMajor=PES stream_id
'                        +0x12 codec info (variable, codec-specific)

'  padded to 2048-byte sector. MPEG-2 PS data begins at 0x800.

Imports System.IO
Imports System.Text

Public Enum PamfStreamType As Byte
    Unknown = &H0
    AVC = &H1B   ' H.264
    MPEG2Video = &H2
    ATRAC3plus = &HDC
    AC3 = &H81   ' 
    LPCM = &H80
    UserData = &HDD
End Enum

Public Class PamfStreamInfo
    Public Property Index As Integer
    Public Property StreamType As PamfStreamType
    Public Property Channel As Byte
    Public Property PesStreamId As Byte        ' MPEG-PS stream_id (0xE0, 0xBD, ...)
    Public Property SubStreamId As Byte        ' for private stream 1 (audio sub-id)
    Public Property PStdBufferRaw As UShort    ' raw bytes 6..7 of stream entry
    Public Property EpOffset As UInteger       ' u32 BE at entry+8
    Public Property EpNum As UInteger          ' u32 BE at entry+12
    Public Property CodecInfo As Byte()        ' raw 32-byte codec-specific blob
    ' video-specific (AVC + M2V)
    Public Property Width As Integer
    Public Property Height As Integer
    Public Property ProfileIdc As Byte
    Public Property LevelIdc As Byte
    Public Property FrameRateInfo As Byte
    Public Property AspectRatioIdc As Byte
    Public Property SarWidth As Integer            ' only meaningful when AR=0xFF
    Public Property SarHeight As Integer
    Public Property VideoSignalInfoFlag As Byte
    Public Property FrameMbsOnlyFlag As Byte       ' AVC only
    Public Property ProgressiveSequence As Byte    ' M2V only
    Public Property FrameCropLeftOffset As Integer
    Public Property FrameCropRightOffset As Integer
    Public Property FrameCropTopOffset As Integer
    Public Property FrameCropBottomOffset As Integer
    Public Property VideoFormat As Byte            ' ci(20) bits 7..5
    Public Property VideoFullRangeFlag As Byte     ' ci(20) bit 4
    Public Property ColourPrimaries As Byte        ' ci(21)
    Public Property TransferCharacteristics As Byte ' ci(22)
    Public Property MatrixCoefficients As Byte     ' ci(23)
    ' AVC-only ci(24) / ci(25)
    Public Property CabacFlag As Byte              ' ci(24) bit 7
    Public Property DeblockingFilterFlag As Byte   ' ci(24) bit 6
    Public Property MinNumSlicePerPictureIdc As Byte ' ci(24) bits 5..4
    Public Property NfwIdc As Byte                 ' ci(24) bits 1..0
    Public Property MaxMeanBitrate As Byte         ' ci(25)
    ' M2V-only
    Public Property HorizontalSizeValue As Integer ' ci(12..13)
    Public Property VerticalSizeValue As Integer   ' ci(14..15)
    Public Property NumChannels As Byte
    Public Property SampleRate As Integer
    Public Property BitsPerSample As Byte
    Public Property OutputExtension As String

    ' decoded P_STD buffer (PAMF entry format: 2 unused + 1 scale + 13 size)
    ' scale=0 -> 128-byte units, scale=1 -> 1024-byte units
    Public ReadOnly Property PStdBufferScale As Byte
        Get
            Return CByte((CInt(PStdBufferRaw) >> 13) And 1)
        End Get
    End Property
    Public ReadOnly Property PStdBufferSize As Integer
        Get
            Return CInt(PStdBufferRaw) And &H1FFF
        End Get
    End Property
    Public ReadOnly Property PStdBufferBytes As Long
        Get
            Return CLng(PStdBufferSize) * If(PStdBufferScale = 1, 1024L, 128L)
        End Get
    End Property

    Public ReadOnly Property TypeName As String
        Get
            Select Case StreamType
                Case PamfStreamType.AVC : Return "AVC (H.264)"
                Case PamfStreamType.MPEG2Video : Return "MPEG-2 Video"
                Case PamfStreamType.ATRAC3plus : Return "ATRAC3plus"
                Case PamfStreamType.AC3 : Return "Dolby Digital (AC3)"
                Case PamfStreamType.LPCM : Return "LPCM"
                Case PamfStreamType.UserData : Return "User Data"
                Case Else : Return $"Unknown (0x{CByte(StreamType):X2})"
            End Select
        End Get
    End Property

    ' decoded frame rate, this is just code-to-fps lookup
    '
    ' PAMF use two different code tables for the same set of physical rates
    '
    '   AVC table              M2V table
    '   0 = 24000/1001         1 = 24000/1001
    '   1 = 24                 2 = 24
    '   2 = 25                 3 = 25
    '   3 = 30000/1001         4 = 30000/1001
    '   4 = 30                 5 = 30
    '   5 = 50                 6 = 50
    '   6 = 60000/1001         7 = 60000/1001
    '
    ' M2V values are off-by-one from AVC values
    Public ReadOnly Property FrameRateFromHeader As Double
        Get
            Dim code As Integer = CInt(FrameRateInfo)
            Select Case StreamType
                Case PamfStreamType.AVC
                    Select Case code
                        Case 0 : Return 24000.0 / 1001.0   ' 23.976
                        Case 1 : Return 24.0
                        Case 2 : Return 25.0
                        Case 3 : Return 30000.0 / 1001.0   ' 29.97
                        Case 4 : Return 30.0
                        Case 5 : Return 50.0
                        Case 6 : Return 60000.0 / 1001.0   ' 59.94
                    End Select
                Case PamfStreamType.MPEG2Video
                    Select Case code
                        Case 1 : Return 24000.0 / 1001.0   ' 23.976
                        Case 2 : Return 24.0
                        Case 3 : Return 25.0
                        Case 4 : Return 30000.0 / 1001.0   ' 29.97
                        Case 5 : Return 30.0
                        Case 6 : Return 50.0
                        Case 7 : Return 60000.0 / 1001.0   ' 59.94
                    End Select
            End Select
            Return 0.0
        End Get
    End Property

    ' SPS info populated by PamfFile.ScanVideoHeaders when the stream is AVC
    ' nothing when not AVC, or when the scan couldn't locate / parse an SPS
    Public Property Sps As H264SpsInfo

    ' MPEG-2 sequence_header + sequence_extension info populated by ScanVideoHeaders when the stream is MPEG-2 Video
    Public Property M2vSeq As M2vSequenceInfo

    ' Frame rate readed from video bitstream itself (authoritative)
    ' Return 0 if no SPS / sequence_header was parsed, or it lacked timing
    Public ReadOnly Property FrameRateFromBitstream As Double
        Get
            If Sps IsNot Nothing AndAlso Sps.FrameRate > 0.0 Then
                Return Sps.FrameRate
            End If
            If M2vSeq IsNot Nothing AndAlso M2vSeq.FrameRate > 0.0 Then
                Return M2vSeq.FrameRate
            End If
            Return 0.0
        End Get
    End Property

    ' alias for some old callers
    Public ReadOnly Property FrameRateFromSps As Double
        Get
            Return FrameRateFromBitstream
        End Get
    End Property

    ' frame rate: bitstream value if available, header value otherwise
    Public ReadOnly Property FrameRate As Double
        Get
            Dim fromBs As Double = FrameRateFromBitstream
            If fromBs > 0.0 Then Return fromBs
            Return FrameRateFromHeader
        End Get
    End Property

    ' ffmpeg-style form for common NTSC rates
    ' empty string if rate is unknown
    Public ReadOnly Property FrameRateFraction As String
        Get
            Dim fps As Double = FrameRate
            If fps = 0.0 Then Return ""
            ' Snap NTSC rates back to their exact fraction for display.
            If Math.Abs(fps - 24000.0 / 1001.0) < 0.001 Then Return "24000/1001"
            If Math.Abs(fps - 30000.0 / 1001.0) < 0.001 Then Return "30000/1001"
            If Math.Abs(fps - 60000.0 / 1001.0) < 0.001 Then Return "60000/1001"
            Return CInt(fps).ToString()
        End Get
    End Property

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.Append($"#{Index} {TypeName} PES=0x{PesStreamId:X2}")
        If StreamType = PamfStreamType.ATRAC3plus _
        OrElse StreamType = PamfStreamType.AC3 _
        OrElse StreamType = PamfStreamType.LPCM Then
            sb.Append($"/sub=0x{SubStreamId:X2}")
        End If

        ' prefer bitstream-derived dimensions when we have them
        Dim wPx As Integer = Width
        Dim hPx As Integer = Height
        If Sps IsNot Nothing AndAlso Sps.WidthPixels > 0 Then
            wPx = Sps.WidthPixels
            hPx = Sps.HeightPixels
        ElseIf M2vSeq IsNot Nothing AndAlso M2vSeq.WidthPixels > 0 Then
            wPx = M2vSeq.WidthPixels
            hPx = M2vSeq.HeightPixels
        End If
        If wPx > 0 Then sb.Append($"  {wPx}x{hPx}")

        Dim fromHdr As Double = FrameRateFromHeader
        Dim fromBs As Double = FrameRateFromBitstream
        If fromBs > 0.0 Then
            Dim sourceTag As String = If(Sps IsNot Nothing, "SPS", "seq_header")
            sb.Append($" @ {fromBs:0.###} fps")
            ' flag disagreements so the user sees when PAMF header is incorrect
            If fromHdr > 0.0 AndAlso Math.Abs(fromBs - fromHdr) > 0.01 Then
                sb.Append($" ({sourceTag}; PAMF header claims {fromHdr:0.###})")
            ElseIf fromHdr = 0.0 Then
                sb.Append($" ({sourceTag})")
            End If
            ' progressive / interlaced flag from bitstream
            If Sps IsNot Nothing Then
                sb.Append(If(Sps.FrameMbsOnlyFlag, " progressive", " interlaced"))
            ElseIf M2vSeq IsNot Nothing AndAlso M2vSeq.HasExtension Then
                sb.Append(If(M2vSeq.ProgressiveSequence, " progressive", " interlaced"))
            End If
        ElseIf fromHdr > 0.0 Then
            sb.Append($" @ {fromHdr:0.###} fps (from PAMF header)")
        End If

        If NumChannels > 0 Then sb.Append($"  {NumChannels}ch @ {SampleRate} Hz")
        Return sb.ToString()
    End Function
End Class

Public Class PamfFile

    Public Property FilePath As String
    Public Property FileSize As Long
    Public Property HeaderSize As Long          ' size of PAMF header (typically 0x800)
    Public Property StreamOffset As Long          ' where MPEG-2 PS begins
    Public Property StreamSize As Long          ' MPEG-2 PS payload size
    Public Property Streams As New List(Of PamfStreamInfo)()

    ' header global fields
    Public Property Version As String = ""               ' e.g. "0041"
    Public Property HeaderSectors As Integer             ' u32 @ 0x08
    Public Property DataSectors As Integer               ' u32 @ 0x0C (= num packs)
    Public Property PsmfMarksOffset As UInteger
    Public Property PsmfMarksSize As UInteger
    Public Property UnkOffset As UInteger
    Public Property UnkSize As UInteger
    Public Property SeqInfoSize As UInteger              ' u32 @ 0x50
    Public Property StartPts90 As Long                   ' 48-bit, combined
    Public Property EndPts90 As Long                     ' 48-bit, combined
    Public Property MuxRateBound As UInteger             ' units of 50 bytes/sec
    Public Property StdDelayBound As UInteger            ' units of 1/90000 sec
    Public Property TotalStreamNum As UInteger
    Public Property GroupingPeriodNum As Byte
    Public Property GroupingPeriodSize As UInteger       ' u32 @ 0x70
    Public Property GpStartPts90 As Long
    Public Property GpEndPts90 As Long
    Public Property GroupNum As Byte
    Public Property GroupSize As UInteger                ' u32 @ 0x82

    Public ReadOnly Property HasEpTable As Boolean
        Get
            For Each s In Streams
                If s.EpNum > 0UI Then Return True
            Next
            Return False
        End Get
    End Property

    Public ReadOnly Property TotalEpEntries As UInteger
        Get
            Dim sum As UInteger = 0UI
            For Each s In Streams
                sum += s.EpNum
            Next
            Return sum
        End Get
    End Property

    Private Const SectorSize As Integer = 2048

    Private Const HeaderScanCap As Integer = 8 * 1024 * 1024

    Public Sub Open(path As String)
        FilePath = path
        Dim fi As New FileInfo(path)
        FileSize = fi.Length

        Dim rawSize As Integer = CInt(Math.Min(CLng(HeaderScanCap), FileSize))
        Dim raw(rawSize - 1) As Byte
        Using fs As FileStream = File.OpenRead(path)
            Dim total As Integer = 0
            While total < rawSize
                Dim n As Integer = fs.Read(raw, total, rawSize - total)
                If n <= 0 Then Exit While
                total += n
            End While
        End Using

        ' magic + version
        If raw.Length < SectorSize OrElse
           raw(0) <> &H50 OrElse raw(1) <> &H41 OrElse
           raw(2) <> &H4D OrElse raw(3) <> &H46 Then
            Throw New InvalidDataException("Not a PAMF file (bad magic).")
        End If
        Version = Encoding.ASCII.GetString(raw, 4, 4)
        If Version <> "0041" Then
            ' other versions exist (0030, 0042, ...)
            Console.Error.WriteLine($"[warn] Unexpected PAMF version '{Version}', continuing.")
        End If

        ' parse rest of header
        ParseGlobalHeaderFields(raw)

        ' elementary mux begins at first 2048-byte sector
        ' we find it by looking for first MPEG-PS pack header (00 00 01 BA)
        StreamOffset = FindMpegPsStart(raw)
        StreamSize = FileSize - StreamOffset
        HeaderSize = StreamOffset

        ' stream table
        ParseStreamTable(raw)

        ' resolve output extensions
        For Each s In Streams
            s.OutputExtension = GuessExtension(s)
        Next

        ' for video streams (AVC and M2V), peek into the program stream area to grab the first SPS / sequence_header
        Try
            ScanVideoHeaders()
        Catch ex As Exception
            Console.Error.WriteLine($"[warn] Video header scan failed: {ex.Message}")
        End Try
    End Sub

    Private Sub ParseGlobalHeaderFields(raw As Byte())
        ' PamfHeader (0x00..0x4F)
        HeaderSectors = CInt(ReadU32BE(raw, &H8))
        DataSectors = CInt(ReadU32BE(raw, &HC))
        PsmfMarksOffset = ReadU32BESafe(raw, &H10)
        PsmfMarksSize = ReadU32BESafe(raw, &H14)
        UnkOffset = ReadU32BESafe(raw, &H18)
        UnkSize = ReadU32BESafe(raw, &H1C)

        ' PamfSequenceInfo @ 0x50
        SeqInfoSize = ReadU32BESafe(raw, &H50)
        Dim sptsHi As Long = CLng(ReadU16BE(raw, &H56))
        Dim sptsLo As Long = CLng(ReadU32BE(raw, &H58)) And &HFFFFFFFFL
        StartPts90 = (sptsHi << 32) Or sptsLo
        Dim eptsHi As Long = CLng(ReadU16BE(raw, &H5C))
        Dim eptsLo As Long = CLng(ReadU32BE(raw, &H5E)) And &HFFFFFFFFL
        EndPts90 = (eptsHi << 32) Or eptsLo
        MuxRateBound = ReadU32BESafe(raw, &H62)
        StdDelayBound = ReadU32BESafe(raw, &H66)
        TotalStreamNum = ReadU32BESafe(raw, &H6A)
        GroupingPeriodNum = raw(&H6F)

        ' PamfGroupingPeriod @ 0x70
        GroupingPeriodSize = ReadU32BESafe(raw, &H70)
        Dim gpSptsHi As Long = CLng(ReadU16BE(raw, &H74))
        Dim gpSptsLo As Long = CLng(ReadU32BE(raw, &H76)) And &HFFFFFFFFL
        GpStartPts90 = (gpSptsHi << 32) Or gpSptsLo
        Dim gpEptsHi As Long = CLng(ReadU16BE(raw, &H7A))
        Dim gpEptsLo As Long = CLng(ReadU32BE(raw, &H7C)) And &HFFFFFFFFL
        GpEndPts90 = (gpEptsHi << 32) Or gpEptsLo
        GroupNum = raw(&H81)

        ' PamfGroup @ 0x82
        GroupSize = ReadU32BESafe(raw, &H82)
    End Sub

    ' use long to avoid overflow exception
    Private Shared Function ReadU32BESafe(b As Byte(), off As Integer) As UInteger
        Return CUInt(CLng(ReadU32BE(b, off)) And &HFFFFFFFFL)
    End Function

    Private Function FindMpegPsStart(raw As Byte()) As Long
        ' pack start code 0x000001BA, limit search to first 64 KiB
        Dim limit As Integer = Math.Min(raw.Length - 4, &H10000)
        For i As Integer = 0 To limit
            If raw(i) = 0 AndAlso raw(i + 1) = 0 _
            AndAlso raw(i + 2) = 1 AndAlso raw(i + 3) = &HBA Then
                Return i
            End If
        Next
        Throw New InvalidDataException("Could not locate MPEG-2 PS pack header.")
    End Function

    Private Sub ParseStreamTable(raw As Byte())
        ' numStreams at 0x86 (BE u16) in version 0041 headers
        Dim numStreams As Integer = ReadU16BE(raw, &H86)
        Dim entryBase As Integer = &H88
        Dim entrySize As Integer = 48           ' 0x30 per entry

        If numStreams <= 0 OrElse numStreams > 32 _
        OrElse entryBase + numStreams * entrySize > raw.Length Then
            Console.Error.WriteLine("[warn] Stream table appears malformed; scanning PES instead.")
            ScanStreamsFromPes(raw)
            Return
        End If

        For i As Integer = 0 To numStreams - 1
            Dim off As Integer = entryBase + i * entrySize
            Dim s As New PamfStreamInfo() With {
                .Index = i,
                .StreamType = CType(raw(off + 0), PamfStreamType),
                .Channel = raw(off + 1)
            }

            ' Entry layout:
            '   +0  streamType (PAMF code: 0x1B AVC, 0x02 M2V, 0xDC AT3p, 0x81 AC3, 0x80 LPCM, 0xDD UserData)
            '   +1  channel
            '   +2..+3  reserved (typically 00 00)
            '   +4  PES stream_id (0xE0 for video, 0xBD for private_stream_1)
            '   +5  sub_stream_id (audio sub-id under private_stream_1)
            '   +6..+7  p_std_buffer (u16 BE: 2 unused | 1 scale | 13 size)
            '   +8..+11 ep_offset  (u32 BE, byte offset of EP table in header)
            '   +12..+15 ep_num    (u32 BE, entries in EP table)
            '   +16..+47 codec-specific info (CellPamfAvcInfo / CellPamfAtrac3plusInfo /...)

            s.PesStreamId = raw(off + 4)
            s.SubStreamId = raw(off + 5)
            s.PStdBufferRaw = CUShort(ReadU16BE(raw, off + 6))
            s.EpOffset = ReadU32BESafe(raw, off + 8)
            s.EpNum = ReadU32BESafe(raw, off + 12)

            Dim ci As Integer = off + 16
            s.CodecInfo = New Byte(31) {}
            Array.Copy(raw, ci, s.CodecInfo, 0, 32)

            Select Case s.StreamType
                Case PamfStreamType.AVC
                    ' AVC codec_info:
                    '   ci(0)  profileIdc
                    '   ci(1)  levelIdc
                    '   ci(2)  packed: bit7=frameMbsOnlyFlag, bit6=videoSignalInfoFlag,
                    '                  bits3:0 = frameRateCode + 1   (note +1 offset only for AVC!)
                    '   ci(3)  aspectRatioIdc
                    '   ci(4..5) sarWidth  (only valid when aspectRatioIdc=0xFF)
                    '   ci(6..7) sarHeight
                    '   ci(8)  reserved1 (=0)
                    '   ci(9)  horizontalSize / 16   (single byte, NOT a u16)
                    '   ci(10) reserved2 (=0)
                    '   ci(11) verticalSize / 16    (single byte)
                    '   ci(12..13) frameCropLeftOffset
                    '   ci(14..15) frameCropRightOffset
                    '   ci(16..17) frameCropTopOffset
                    '   ci(18..19) frameCropBottomOffset
                    '   ci(20) packed: bits 7..5 = videoFormat,
                    '                  bit 4 = fullRangeFlag
                    '   ci(21) colourPrimaries
                    '   ci(22) transferCharacteristics
                    '   ci(23) matrixCoefficients
                    '   ci(24) packed: bit7 = entropyCodingModeFlag (CABAC),
                    '                  bit6 = deblockingFilterFlag,
                    '                  bits 5..4 = minNumSlicePerPictureIdc,
                    '                  bits 1..0 = nfwIdc
                    '   ci(25) maxMeanBitrate
                    s.ProfileIdc = raw(ci + 0)
                    s.LevelIdc = raw(ci + 1)
                    Dim x2 As Byte = raw(ci + 2)
                    s.FrameMbsOnlyFlag = CByte((x2 >> 7) And 1)
                    s.VideoSignalInfoFlag = CByte((x2 >> 6) And 1)
                    ' AVC stored value is code + 1, subtract 1 to get CELL_PAMF_AVC_FRC_* code (0..6)
                    Dim avcFrcRaw As Integer = CInt(x2 And &HF)
                    s.FrameRateInfo = CByte(If(avcFrcRaw > 0, avcFrcRaw - 1, 0))
                    s.AspectRatioIdc = raw(ci + 3)
                    s.SarWidth = ReadU16BE(raw, ci + 4)
                    s.SarHeight = ReadU16BE(raw, ci + 6)
                    s.Width = CInt(raw(ci + 9)) * 16
                    s.Height = CInt(raw(ci + 11)) * 16
                    s.FrameCropLeftOffset = ReadU16BE(raw, ci + 12)
                    s.FrameCropRightOffset = ReadU16BE(raw, ci + 14)
                    s.FrameCropTopOffset = ReadU16BE(raw, ci + 16)
                    s.FrameCropBottomOffset = ReadU16BE(raw, ci + 18)
                    Dim x14 As Byte = raw(ci + 20)
                    s.VideoFormat = CByte((x14 >> 5) And 7)
                    s.VideoFullRangeFlag = CByte((x14 >> 4) And 1)
                    s.ColourPrimaries = raw(ci + 21)
                    s.TransferCharacteristics = raw(ci + 22)
                    s.MatrixCoefficients = raw(ci + 23)
                    Dim x18 As Byte = raw(ci + 24)
                    s.CabacFlag = CByte((x18 >> 7) And 1)
                    s.DeblockingFilterFlag = CByte((x18 >> 6) And 1)
                    s.MinNumSlicePerPictureIdc = CByte((x18 >> 4) And 3)
                    s.NfwIdc = CByte(x18 And 3)
                    s.MaxMeanBitrate = raw(ci + 25)

                Case PamfStreamType.MPEG2Video
                    ' M2V codec_info:
                    '   ci(0)  profileAndLevel (raw MPEG-2 byte: 0x44 = MP@HL, 0x48 = MP@ML)
                    '   ci(1)  unused
                    '   ci(2)  packed: bit7=progressiveSequence, bit6=videoSignalInfoFlag, bits3:0 = frameRateCode  (no +1 offset, unlike AVC)
                    '   ci(3)  aspectRatioIdc
                    '   ci(4..5) sarWidth  ci(6..7) sarHeight
                    '   ci(8)  reserved1
                    '   ci(9)  horizontalSize / 16
                    '   ci(10) reserved2
                    '   ci(11) verticalSize / 16
                    '   ci(12..13) horizontalSizeValue (full pixel width, u16 BE)
                    '   ci(14..15) verticalSizeValue   (full pixel height)
                    '   ci(20) x14: videoFormat + fullRangeFlag (same bit layout as AVC)
                    '   ci(21..23) colour primaries / transfer / matrix
                    s.ProfileIdc = raw(ci + 0)
                    Dim x2 As Byte = raw(ci + 2)
                    s.ProgressiveSequence = CByte((x2 >> 7) And 1)
                    s.VideoSignalInfoFlag = CByte((x2 >> 6) And 1)
                    s.FrameRateInfo = CByte(x2 And &HF)
                    s.AspectRatioIdc = raw(ci + 3)
                    s.SarWidth = ReadU16BE(raw, ci + 4)
                    s.SarHeight = ReadU16BE(raw, ci + 6)
                    ' prefer explicit pixel size if present else derive from MB count
                    s.HorizontalSizeValue = ReadU16BE(raw, ci + 12)
                    s.VerticalSizeValue = ReadU16BE(raw, ci + 14)
                    If s.HorizontalSizeValue > 0 AndAlso s.VerticalSizeValue > 0 Then
                        s.Width = s.HorizontalSizeValue
                        s.Height = s.VerticalSizeValue
                    Else
                        s.Width = CInt(raw(ci + 9)) * 16
                        s.Height = CInt(raw(ci + 11)) * 16
                    End If
                    Dim x14 As Byte = raw(ci + 20)
                    s.VideoFormat = CByte((x14 >> 5) And 7)
                    s.VideoFullRangeFlag = CByte((x14 >> 4) And 1)
                    s.ColourPrimaries = raw(ci + 21)
                    s.TransferCharacteristics = raw(ci + 22)
                    s.MatrixCoefficients = raw(ci + 23)

                Case PamfStreamType.ATRAC3plus, PamfStreamType.AC3
                    ' audio codec_info:
                    '   ci(0..1) unknown u16 (=0)
                    '   ci(2)    channels (1, 2, 6, 8)
                    '   ci(3)    samplingFrequency code (1 = 48 kHz, only value Sony uses)
                    s.NumChannels = raw(ci + 2)
                    Dim fsCode As Integer = raw(ci + 3)
                    s.SampleRate = If(fsCode = 1, 48000, 0)

                Case PamfStreamType.LPCM
                    ' LPCM bit-depth is a coded byte at ci(4)
                    '   0x01 / 0x40 => 16-bit
                    '   0x03 / 0x50 => 24-bit  (0x50 inferred, unverified)
                    ' anything else -> 0 = unknown (WriteWavHeader falls back to 16)
                    s.NumChannels = raw(ci + 2)
                    Dim fsCode As Integer = raw(ci + 3)
                    s.SampleRate = If(fsCode = 1, 48000, 0)
                    Dim bpsCode As Integer = raw(ci + 4)
                    Select Case bpsCode
                        Case &H1, &H40 : s.BitsPerSample = 16
                        Case &H3, &H50 : s.BitsPerSample = 24
                        Case Else : s.BitsPerSample = 0
                    End Select
            End Select

            Streams.Add(s)
        Next
    End Sub

    Private Sub ScanStreamsFromPes(raw As Byte())
        ' fallback when the header table can't be trusted, discover stream IDs in the PS payload (scan first 4 mb only)
        Dim seenVideo As New HashSet(Of Byte)()
        Dim seenAudio As New HashSet(Of Tuple(Of Byte, Byte))()
        Dim ps As Long = StreamOffset
        Dim len As Long = Math.Min(raw.LongLength - ps, 4L * 1024L * 1024L)
        Dim i As Long = ps
        While i < ps + len - 16
            If raw(CInt(i)) = 0 AndAlso raw(CInt(i + 1)) = 0 AndAlso raw(CInt(i + 2)) = 1 Then
                Dim sid As Byte = raw(CInt(i + 3))
                If sid >= &HE0 AndAlso sid <= &HEF Then
                    seenVideo.Add(sid)
                ElseIf sid = &HBD Then
                    Dim hdrLen As Integer = raw(CInt(i + 8))
                    Dim subId As Byte = raw(CInt(i + 9 + hdrLen))
                    seenAudio.Add(Tuple.Create(sid, subId))
                End If
            End If
            i += 1
        End While

        Dim idx As Integer = 0
        For Each sid In seenVideo
            Streams.Add(New PamfStreamInfo() With {
                .Index = idx, .PesStreamId = sid,
                .StreamType = PamfStreamType.AVC ' best guess
            })
            idx += 1
        Next
        For Each t In seenAudio
            Dim subType As PamfStreamType = AudioSubIdToType(t.Item2)
            Streams.Add(New PamfStreamInfo() With {
                .Index = idx, .PesStreamId = t.Item1, .SubStreamId = t.Item2,
                .StreamType = subType
            })
            idx += 1
        Next
    End Sub

    Private Function AudioSubIdToType(subId As Byte) As PamfStreamType
        ' PAMF private_stream_1 sub-id ranges (PSS-style):
        '   0x00 - 0x0F  ATRAC3plus
        '   0x30 - 0x3F  AC3
        '   0xA0 - 0xAF  LPCM
        Select Case subId
            Case Is <= &HF : Return PamfStreamType.ATRAC3plus
            Case &H30 To &H3F : Return PamfStreamType.AC3
            Case &HA0 To &HAF : Return PamfStreamType.LPCM
            Case Else : Return PamfStreamType.Unknown
        End Select
    End Function

    Private Function GuessExtension(s As PamfStreamInfo) As String
        Select Case s.StreamType
            Case PamfStreamType.AVC : Return ".h264"
            Case PamfStreamType.MPEG2Video : Return ".m2v"
            Case PamfStreamType.ATRAC3plus : Return ".at3"    ' RIFF WAVE-wrapped at3+ (playable)
            Case PamfStreamType.AC3 : Return ".ac3"
            Case PamfStreamType.LPCM : Return ".lpcm"   ' raw 16/24-bit BE PCM
            Case PamfStreamType.UserData : Return ".udat"
            Case Else : Return ".bin"
        End Select
    End Function

    Public Sub ExtractAll(outputDir As String, Optional wrapLpcmAsWav As Boolean = True)
        Directory.CreateDirectory(outputDir)
        Dim baseName As String = Path.GetFileNameWithoutExtension(FilePath)

        ' Open one writer per stream we plan to keep.
        Dim writers As New Dictionary(Of Integer, BinaryWriter)()
        Dim lpcmStreams As New Dictionary(Of Integer, PamfStreamInfo)()
        Dim lpcmCounts As New Dictionary(Of Integer, Long)()    ' bytes emitted
        Dim lpcmSwappers As New Dictionary(Of Integer, LpcmBeToLeSwapper)()
        Dim at3Streams As New Dictionary(Of Integer, PamfStreamInfo)()
        Dim at3Strippers As New Dictionary(Of Integer, Atrac3PlusAuStripper)()

        Try
            For Each s In Streams
                If s.StreamType = PamfStreamType.UserData Then Continue For
                Dim outName As String = $"{baseName}.s{s.Index:D2}_{s.TypeName.Split(" "c)(0).ToLower()}{s.OutputExtension}"
                Dim outPath As String = Path.Combine(outputDir, outName)
                Dim bw As New BinaryWriter(File.Create(outPath))
                writers(s.Index) = bw
                Console.WriteLine($"  -> {outName}   ({s})")

                If s.StreamType = PamfStreamType.LPCM AndAlso wrapLpcmAsWav Then
                    ' Reserve 44 bytes for WAV header, filled in after
                    bw.Write(New Byte(43) {})
                    lpcmStreams(s.Index) = s
                    lpcmCounts(s.Index) = 0
                    Dim bps As Integer = If(s.BitsPerSample > 0, CInt(s.BitsPerSample), 16)
                    ' mono LPCM padded with a dummy 2nd channel on the wire
                    Dim actualCh As Integer = Math.Max(1, CInt(s.NumChannels))
                    Dim wireCh As Integer = actualCh + (actualCh And 1)
                    lpcmSwappers(s.Index) = New LpcmBeToLeSwapper(bw, bps \ 8,
                                                                  channelsOnWire:=wireCh,
                                                                  channelsToWrite:=actualCh)
                ElseIf s.StreamType = PamfStreamType.ATRAC3plus Then
                    ' Reserve room for RIFF AT3+ header, filled when frame count is known
                    ' WriteAt3RiffHeader writes exactly Atrac3PlusAuStripper.At3RiffHeaderLen bytes
                    bw.Write(New Byte(Atrac3PlusAuStripper.At3RiffHeaderLen - 1) {})
                    at3Streams(s.Index) = s
                    at3Strippers(s.Index) = New Atrac3PlusAuStripper(bw)
                End If
            Next

            ' Walk the MPEG-2 PS once, dispatch payload by (stream_id, sub_id)
            Using fs As FileStream = File.OpenRead(FilePath)
                fs.Position = StreamOffset
                DemuxProgramStream(fs, writers, lpcmStreams, lpcmCounts,
                                   at3Strippers, lpcmSwappers)
            End Using

            For Each kv In writers
                kv.Value.Flush()
            Next

            ' Patch LPCM WAV headers
            For Each kv In lpcmStreams
                Dim w As BinaryWriter = writers(kv.Key)
                w.BaseStream.Position = 0
                WriteWavHeader(w, kv.Value, lpcmCounts(kv.Key))
            Next

            ' Patch ATRAC3plus RIFF headers (now that we know the frame count),
            ' and append the atsc chunk (per-frame ATS extra_config_data) at EOF.
            For Each kv In at3Streams
                Dim w As BinaryWriter = writers(kv.Key)
                Dim s As PamfStreamInfo = kv.Value
                Dim stripper As Atrac3PlusAuStripper = at3Strippers(kv.Key)
                ' 1) trailer first (extends file to its final length), so the
                '    RIFF size we write in the header actually matches on-disk.
                w.BaseStream.Seek(0, SeekOrigin.End)
                WriteAt3AtscTrailer(w, stripper)
                ' 2) patch the fixed-size prefix (RIFF / fmt / fact / data-hdr)
                w.BaseStream.Position = 0
                WriteAt3RiffHeader(w, s, stripper)
                Console.WriteLine($"     at3+ s{kv.Key:D2}: {stripper.FramesWritten} frames of {stripper.StrippedFrameSize} bytes " &
                                  $"({stripper.FramesWritten * stripper.StrippedFrameSize:N0} ES bytes)")
            Next

        Finally
            For Each w In writers.Values
                w.Dispose()
            Next
        End Try
    End Sub

    ' Peek into the stream, accumulate first 64 kb of video PES payload per video stream, then parse:
    ' SPS NAL for AVC, sequence_header for M2V
    Private Sub ScanVideoHeaders()
        ' Bail early if no video streams to scan
        Dim haveVideo As Boolean = False
        For Each s In Streams
            If s.StreamType = PamfStreamType.AVC _
            OrElse s.StreamType = PamfStreamType.MPEG2Video Then
                haveVideo = True : Exit For
            End If
        Next
        If Not haveVideo Then Return

        ' Build a sid -> accumulator map for the video PES IDs
        Dim accumulator As New Dictionary(Of Byte, List(Of Byte))()
        Dim sidToStream As New Dictionary(Of Byte, PamfStreamInfo)()
        For Each s In Streams
            If s.StreamType = PamfStreamType.AVC _
            OrElse s.StreamType = PamfStreamType.MPEG2Video Then
                accumulator(s.PesStreamId) = New List(Of Byte)()
                sidToStream(s.PesStreamId) = s
            End If
        Next

        Const ScanBudgetBytes As Integer = 1024 * 1024     ' 1 mb of PS is big
        Const PerStreamCap As Integer = 64 * 1024       ' 64 kb collected per stream

        Using fs As FileStream = File.OpenRead(FilePath)
            fs.Position = StreamOffset
            Dim buf(ScanBudgetBytes - 1) As Byte
            Dim read As Integer = fs.Read(buf, 0, buf.Length)

            Dim p As Integer = 0
            While p < read - 6
                If buf(p) <> 0 OrElse buf(p + 1) <> 0 OrElse buf(p + 2) <> 1 Then
                    p += 1
                    Continue While
                End If
                Dim sid As Byte = buf(p + 3)
                Select Case sid
                    Case &HBA   ' pack_header
                        If p + 14 > read Then Exit While
                        p += 14 + (buf(p + 13) And 7)

                    Case &HB9   ' MPEG_program_end_code
                        p += 4

                    Case &HBB, &HBC, &HBE, &HBF
                        If p + 6 > read Then Exit While
                        Dim plen As Integer = (CInt(buf(p + 4)) << 8) Or buf(p + 5)
                        If p + 6 + plen > read Then Exit While
                        p += 6 + plen

                    Case &HBD     ' audio PES, ignore for video header scan
                        If p + 6 > read Then Exit While
                        Dim plen As Integer = (CInt(buf(p + 4)) << 8) Or buf(p + 5)
                        If p + 6 + plen > read Then Exit While
                        p += 6 + plen

                    Case Else
                        If sid >= &HE0 AndAlso sid <= &HEF Then
                            If p + 9 > read Then Exit While
                            Dim pesLen As Integer = (CInt(buf(p + 4)) << 8) Or buf(p + 5)
                            Dim total As Integer = 6 + pesLen
                            If p + total > read Then Exit While
                            Dim hdrLen As Integer = buf(p + 8)
                            Dim payOff As Integer = p + 9 + hdrLen
                            Dim payLen As Integer = (p + total) - payOff
                            If accumulator.ContainsKey(sid) _
                            AndAlso accumulator(sid).Count < PerStreamCap _
                            AndAlso payLen > 0 Then
                                Dim slice(payLen - 1) As Byte
                                Array.Copy(buf, payOff, slice, 0, payLen)
                                accumulator(sid).AddRange(slice)
                            End If
                            p += total
                        Else
                            p += 4    ' unknown start code, resync
                        End If
                End Select

                ' Stop early once every video stream has filled its quota
                Dim allFull As Boolean = True
                For Each kv In accumulator
                    If kv.Value.Count < PerStreamCap Then allFull = False : Exit For
                Next
                If allFull Then Exit While
            End While
        End Using

        ' Parse the accumulated bytes per stream, dispatching by codec
        For Each kv In accumulator
            If kv.Value.Count = 0 Then Continue For
            Dim s As PamfStreamInfo = sidToStream(kv.Key)
            Dim payload As Byte() = kv.Value.ToArray()
            Select Case s.StreamType
                Case PamfStreamType.AVC
                    Dim sps As H264SpsInfo = H264SpsParser.ParseFirstSps(
                        payload, 0, payload.Length)
                    If sps IsNot Nothing Then s.Sps = sps
                Case PamfStreamType.MPEG2Video
                    Dim seq As M2vSequenceInfo = MpegSequenceHeaderParser.ParseFirstSequenceHeader(
                        payload, 0, payload.Length)
                    If seq IsNot Nothing Then s.M2vSeq = seq
            End Select
        Next
    End Sub

    ' MPEG-2 Program Stream:
    '   0x000001BA  pack_start_code
    '   0x000001BB  system_header
    '   0x000001BC  program_stream_map
    '   0x000001BD  private_stream_1  (audio)
    '   0x000001BE  padding
    '   0x000001BF  private_stream_2  (program info / DVD nav)
    '   0x000001E0..EF  video stream
    '   0x000001B9  MPEG_program_end_code
    Private Sub DemuxProgramStream(fs As FileStream,
                                   writers As Dictionary(Of Integer, BinaryWriter),
                                   lpcmStreams As Dictionary(Of Integer, PamfStreamInfo),
                                   lpcmCounts As Dictionary(Of Integer, Long),
                                   at3Strippers As Dictionary(Of Integer, Atrac3PlusAuStripper),
                                   lpcmSwappers As Dictionary(Of Integer, LpcmBeToLeSwapper))

        ' Index streams by PES (and sub-id) for O(1) dispatch
        Dim videoByPes As New Dictionary(Of Byte, PamfStreamInfo)()
        Dim audioByKey As New Dictionary(Of Integer, PamfStreamInfo)()
        For Each s In Streams
            If s.PesStreamId >= &HE0 AndAlso s.PesStreamId <= &HEF Then
                videoByPes(s.PesStreamId) = s
            ElseIf s.PesStreamId = &HBD Then
                audioByKey((CInt(s.PesStreamId) << 8) Or s.SubStreamId) = s
            End If
        Next

        Const BufCap As Integer = 1024 * 1024
        Dim buf(BufCap - 1) As Byte
        Dim pending As Integer = 0          ' bytes carried from previous chunk
        Dim eof As Boolean = False

        Do
            ' Refill buffer
            Dim want As Integer = BufCap - pending
            Dim got As Integer = fs.Read(buf, pending, want)
            If got = 0 Then eof = True
            Dim filled As Integer = pending + got

            Dim p As Integer = 0
            While p + 6 <= filled
                ' Need start code 00 00 01 xx
                If buf(p) <> 0 OrElse buf(p + 1) <> 0 OrElse buf(p + 2) <> 1 Then
                    p += 1
                    Continue While
                End If
                Dim startCode As Byte = buf(p + 3)

                Select Case startCode
                    Case &HBA   ' pack header (14 bytes + stuffing)
                        If p + 14 > filled Then Exit While
                        Dim stuffing As Integer = buf(p + 13) And &H7
                        Dim packLen As Integer = 14 + stuffing
                        If p + packLen > filled Then Exit While
                        p += packLen

                    Case &HB9   ' MPEG_program_end_code
                        p += 4

                    Case &HBB, &HBC, &HBE, &HBF, &HF0 To &HFF
                        ' System header / PSM / padding / private_stream_2
                        ' Layout: 4-byte start code + 2-byte length + payload
                        If p + 6 > filled Then Exit While
                        Dim pLen As Integer = (CInt(buf(p + 4)) << 8) Or buf(p + 5)
                        If p + 6 + pLen > filled Then Exit While
                        p += 6 + pLen

                    Case &HBD   ' private_stream_1 (audio)
                        Dim consumed As Integer
                        If Not HandlePes(buf, p, filled, audioByKey, writers,
                                         lpcmStreams, lpcmCounts, isAudio:=True,
                                         consumed:=consumed,
                                         at3Strippers:=at3Strippers,
                                         lpcmSwappers:=lpcmSwappers) Then
                            Exit While
                        End If
                        p += consumed

                    Case &HE0 To &HEF   ' video
                        Dim consumed As Integer
                        If Not HandlePes(buf, p, filled, audioByKey, writers,
                                         lpcmStreams, lpcmCounts, isAudio:=False,
                                         consumed:=consumed,
                                         videoMap:=videoByPes,
                                         at3Strippers:=at3Strippers,
                                         lpcmSwappers:=lpcmSwappers) Then
                            Exit While
                        End If
                        p += consumed

                    Case Else
                        ' Unknown: skip 4 bytes and resync
                        p += 4
                End Select
            End While

            ' Carry trailing bytes (which might start a partial pack/PES) to next refill so we never split start code
            pending = filled - p
            If pending > 0 Then
                Buffer.BlockCopy(buf, p, buf, 0, pending)
            End If

            If eof AndAlso pending < 6 Then Exit Do
        Loop
    End Sub

    Private Function HandlePes(buf As Byte(),
                               p As Integer,
                               filled As Integer,
                               audioByKey As Dictionary(Of Integer, PamfStreamInfo),
                               writers As Dictionary(Of Integer, BinaryWriter),
                               lpcmStreams As Dictionary(Of Integer, PamfStreamInfo),
                               lpcmCounts As Dictionary(Of Integer, Long),
                               isAudio As Boolean,
                               ByRef consumed As Integer,
                               Optional videoMap As Dictionary(Of Byte, PamfStreamInfo) = Nothing,
                               Optional at3Strippers As Dictionary(Of Integer, Atrac3PlusAuStripper) = Nothing,
                               Optional lpcmSwappers As Dictionary(Of Integer, LpcmBeToLeSwapper) = Nothing) As Boolean

        If p + 9 > filled Then
            consumed = 0
            Return False    ' need more data
        End If

        Dim sid As Byte = buf(p + 3)
        Dim pesLen As Integer = (CInt(buf(p + 4)) << 8) Or buf(p + 5)
        Dim total As Integer = 6 + pesLen
        If p + total > filled Then
            consumed = 0
            Return False
        End If

        ' MPEG-2 PES extension prefix: marker '10' in top 2 bits of byte 6
        Dim ext1 As Byte = buf(p + 6)
        Dim hdrLen As Integer = buf(p + 8)
        Dim payOff As Integer = p + 9 + hdrLen
        Dim payEnd As Integer = p + total
        Dim payLen As Integer = payEnd - payOff

        If payLen <= 0 Then
            consumed = total
            Return True
        End If

        If isAudio Then
            ' Private stream 1: first byte of payload is sub_stream_id
            ' Layout (PAMF):
            '   byte 0 : sub_id
            '   byte 1 : num frame headers
            '   byte 2..3 : first AU offset (BE u16)
            '   byte 4..  : ES payload (next 4 bytes are an LPCM/audio aux header
            '               for LPCM only)
            Dim subId As Byte = buf(payOff)
            Dim key As Integer = (CInt(sid) << 8) Or subId
            Dim s As PamfStreamInfo = Nothing
            If audioByKey.TryGetValue(key, s) AndAlso writers.ContainsKey(s.Index) Then
                ' 4-byte audio sub-header, then ES data. This is the layout for
                ' Sony's real PS3 PAMFs put PCM samples directly after the sub-header, without any per-frame aux buf
                Dim subHdr As Integer = 4
                Dim esOff As Integer = payOff + subHdr
                Dim esLen As Integer = payEnd - esOff
                If esLen > 0 Then
                    ' ATRAC3plus AU carry an 8-byte PAMF-specific prefix (0FD0 4855 00000000) that must be stripped for ffmpeg.
                    ' AUs span PES boundaries, so we write through a per-stream rolling stripper
                    If at3Strippers IsNot Nothing AndAlso at3Strippers.ContainsKey(s.Index) Then
                        at3Strippers(s.Index).Append(buf, esOff, esLen)
                    ElseIf lpcmSwappers IsNot Nothing AndAlso lpcmSwappers.ContainsKey(s.Index) Then
                        ' PAMF LPCM samples are big-endian, may straddle PES boundaries.
                        ' for mono streams, the swapper drops the dummy silence channel
                        Dim sw As LpcmBeToLeSwapper = lpcmSwappers(s.Index)
                        Dim before As Long = sw.BytesWritten
                        sw.Append(buf, esOff, esLen)
                        lpcmCounts(s.Index) += sw.BytesWritten - before
                    Else
                        writers(s.Index).Write(buf, esOff, esLen)
                        If lpcmCounts.ContainsKey(s.Index) Then
                            lpcmCounts(s.Index) += esLen
                        End If
                    End If
                End If
            End If
        Else
            ' Video: payload is already Annex-B / MPEG-2 ES bytes
            Dim s As PamfStreamInfo = Nothing
            If videoMap IsNot Nothing AndAlso videoMap.TryGetValue(sid, s) _
            AndAlso writers.ContainsKey(s.Index) Then
                writers(s.Index).Write(buf, payOff, payLen)
            End If
        End If

        consumed = total
        Return True
    End Function

    Private Sub WriteWavHeader(bw As BinaryWriter, s As PamfStreamInfo, dataBytes As Long)
        ' PAMF LPCM is big-endian samples, the data has already been byte-swapped by LpcmBeToLeSwapper before this
        ' WAV is directly playable
        Dim sampleRate As Integer = If(s.SampleRate > 0, s.SampleRate, 48000)
        Dim channels As Integer = If(s.NumChannels > 0, CInt(s.NumChannels), 2)
        Dim bps As Integer = If(s.BitsPerSample > 0, CInt(s.BitsPerSample), 16)
        Dim byteRate As Integer = sampleRate * channels * (bps \ 8)
        Dim blockAlign As Integer = channels * (bps \ 8)

        bw.Write(Encoding.ASCII.GetBytes("RIFF"))
        bw.Write(CInt(36 + dataBytes))
        bw.Write(Encoding.ASCII.GetBytes("WAVE"))
        bw.Write(Encoding.ASCII.GetBytes("fmt "))
        bw.Write(16)                              ' fmt chunk size
        bw.Write(CShort(1))                       ' PCM
        bw.Write(CShort(channels))
        bw.Write(sampleRate)
        bw.Write(byteRate)
        bw.Write(CShort(blockAlign))
        bw.Write(CShort(bps))
        bw.Write(Encoding.ASCII.GetBytes("data"))
        bw.Write(CInt(dataBytes))
    End Sub

    Private Sub WriteAt3RiffHeader(bw As BinaryWriter,
                                   s As PamfStreamInfo,
                                   stripper As Atrac3PlusAuStripper)
        ' write RIFF/WAVE file using WAVE_FORMAT_EXTENSIBLE with Sony ATRAC3plus SubFormat GUID
        '
        ' data chunk that follows carries `frameCount` raw_data_frame blocks with ATS header stripped
        ' nBlockAlign = stripped size
        '
        ' custom "atsc" chunk carries ATS extra_config_data (bytes 4-7 of ATS header)
        ' stored per-frame (4 bytes per frame) so muxer can round-trip streams where extra_config_data varies frame-to-frame
        '
        ' SubFormat GUID for ATRAC3plus (little-endian):
        '   BFAA23E9-58CB-7144-A119-FFFA01E4CE62

        Dim channels As Integer = If(s.NumChannels > 0, CInt(s.NumChannels), 2)
        Dim sampleRate As Integer = If(s.SampleRate > 0, s.SampleRate, 48000)
        Dim frameSize As Integer = stripper.StrippedFrameSize
        If frameSize <= 0 Then
            ' fallback, this shouldn't happen for a stream with any AT3+ frames
            frameSize = 688
        End If
        Dim frameCount As Long = stripper.FramesWritten
        Dim avgBps As Integer = CInt(CLng(frameSize) * sampleRate \ Atrac3PlusAuStripper.SamplesPerFrame)
        Dim dataBytes As Long = frameCount * frameSize
        Dim samplesTot As Long = frameCount * Atrac3PlusAuStripper.SamplesPerFrame

        Dim guid() As Byte = New Byte() {
            &HBF, &HAA, &H23, &HE9, &H58, &HCB, &H71, &H44,
            &HA1, &H19, &HFF, &HFA, &H1, &HE4, &HCE, &H62}

        ' fmt chunk: 18 (base WAVEFORMATEX) + 22 (extension) = 40 bytes
        Dim fmtSize As Integer = 40

        ' trailing atsc chunk payload = frameCount * 4
        ' chunk header (8 bytes) is included when it's present
        ' frameCount == 0 means no atsc chunk at all, so no 8-byte header either
        Dim atscPayload As Long = frameCount * 4L
        Dim atscTotal As Long = If(atscPayload > 0, atscPayload + 8L, 0L)

        ' RIFF header (12) + fmt (8+40=48) + fact (8+8=16) + data hdr (8) + data (dataBytes) + atsc trailer = full file
        Dim riffSize As Long = CLng(4 + 48 + 16 + 8) + dataBytes + atscTotal

        ' KSAUDIO channel masks
        Dim mask As UInteger
        Select Case channels
            Case 1 : mask = &H4UI        ' FC
            Case 2 : mask = &H3UI        ' FL FR
            Case 6 : mask = &H3FUI       ' FL FR FC LFE BL BR (5.1)
            Case 8 : mask = &H63FUI      ' FL FR FC LFE BL BR SL SR (7.1)
            Case Else : mask = 0UI
        End Select

        bw.Write(Encoding.ASCII.GetBytes("RIFF"))
        bw.Write(CUInt(Math.Min(riffSize, &HFFFFFFFFL)))
        bw.Write(Encoding.ASCII.GetBytes("WAVE"))

        bw.Write(Encoding.ASCII.GetBytes("fmt "))
        bw.Write(CInt(fmtSize))
        bw.Write(CUShort(&HFFFE))             ' wFormatTag = WAVE_FORMAT_EXTENSIBLE
        bw.Write(CUShort(channels))
        bw.Write(CInt(sampleRate))
        bw.Write(CInt(avgBps))
        bw.Write(CUShort(frameSize))          ' nBlockAlign = raw_data_frame size (ATS header stripped)
        bw.Write(CUShort(0))                  ' wBitsPerSample (compressed)
        bw.Write(CUShort(22))                 ' cbSize (extension length)
        bw.Write(CUShort(Atrac3PlusAuStripper.SamplesPerFrame))   ' wValidBitsPerSample = samples/block
        bw.Write(mask)
        bw.Write(guid)

        bw.Write(Encoding.ASCII.GetBytes("fact"))
        bw.Write(CInt(8))
        bw.Write(CInt(samplesTot And &HFFFFFFFFL))
        bw.Write(CInt(0))                     ' delay samples / reserved

        bw.Write(Encoding.ASCII.GetBytes("data"))
        bw.Write(CUInt(Math.Min(dataBytes, &HFFFFFFFFL)))
    End Sub

    ' append the "atsc" chunk after the data chunk
    ' payload = FramesWritten * 4 bytes
    ' every 4-byte block is one frame's ATS extra_config_data in stream order
    Private Sub WriteAt3AtscTrailer(bw As BinaryWriter, stripper As Atrac3PlusAuStripper)
        Dim n As Long = stripper.FramesWritten
        If n <= 0 Then Return
        bw.Write(Encoding.ASCII.GetBytes("atsc"))
        bw.Write(CInt(n * 4L))
        For Each ec In stripper.PerFrameExtraConfig
            bw.Write(ec, 0, 4)
        Next
    End Sub

    ' -- Helpers --------------------------------------------------------------

    Private Shared Function ReadU16BE(b As Byte(), off As Integer) As Integer
        Return (CInt(b(off)) << 8) Or b(off + 1)
    End Function

    Private Shared Function ReadU32BE(b As Byte(), off As Integer) As Integer
        Return (CInt(b(off)) << 24) Or (CInt(b(off + 1)) << 16) _
            Or (CInt(b(off + 2)) << 8) Or b(off + 3)
    End Function

End Class

' Each ATRAC3plus access unit emitted by PAMF carry 8-byte PAMF-specific prefix before real at3+ bitstream:
'
' +0..+1   0x0FD0      PAMF audio AU sync
' +2..+3   0x4855      constant marker (likely unit-type / version field)
' +4..+7   0x00000000  reserved / zero
' +8..+695 688-byte standard ATRAC3plus frame
'
' a PAMF ATRAC3plus access unit is exactly one ATRAC-X raw_data_frame preceded by an 8-byte ATS header 
' the ATS header starts with 0x0FD0 and encodes frame length in next 16 bits
'
' frame_size = ((data & 0x3ff) + 1) * 8 + 8   [ATS_HEADER_SIZE = 8]
'
' size is uniform across a stream but VARIES BY CHANNEL COUNT AND BITRATE:
'   stereo 128 kbps  ->  696 bytes  (data & 0x3ff = 85)
'   5.1   ~315 kbps  -> 1720 bytes  (data & 0x3ff = 213)
'
' frames span PES boundaries, so
' - feed every audio-payload byte through per-stream rolling buffer
' - size-lock on the first ATS header
' - emit each frame with 8-byte ATS header stripped off (needed for ffmpeg/VLC compatibility)
'
Friend Class Atrac3PlusAuStripper

    Public Const AtsHeaderSize As Integer = 8      ' 0x0FD0 sync (2) + size (2) + reserved (4)
    Public Const SyncWord As Integer = &HFD0
    Public Const SamplesPerFrame As Integer = 2048

    ' Header byte count produced by WriteAt3RiffHeader. ExtractAll reserves these bytes at the head of every at3+ output before demux:
    '   12 bytes RIFF / size / WAVE
    '   48 bytes fmt  chunk header + 40-byte WAVEFORMATEXTENSIBLE payload
    '   16 bytes fact chunk header + 8-byte payload
    '    8 bytes data chunk header
    ' the atsc chunk is written AFTER the data chunk so it can be variable-length (FramesWritten * 4 bytes of payload)
    Public Const At3RiffHeaderLen As Integer = 84

    Private ReadOnly _bw As BinaryWriter
    Private _frameSize As Integer = 0     ' full frame size incl. ATS header, resolved from the first ATS we see
    Private _sizeProbe As Byte()          ' rolling 4-byte probe used until _frameSize is known
    Private _sizeProbeFilled As Integer = 0
    Private _frame As Byte()              ' full-frame buffer (allocated once frameSize is known)
    Private _filled As Integer = 0
    Private ReadOnly _perFrameExtra As New List(Of Byte())()
    Public Property FramesWritten As Long

    ' StrippedFrameSize is 0 until first ATS header has been parsed
    ' callers should read it after demux to fill the RIFF fmt chunk's nBlockAlign
    Public ReadOnly Property StrippedFrameSize As Integer
        Get
            Return If(_frameSize > 0, _frameSize - AtsHeaderSize, 0)
        End Get
    End Property

    ' One 4-byte entry per emitted frame - the ATS header's bytes 4-7 for that
    ' frame. Written out by the extractor as an "atsc" chunk of size
    ' FramesWritten * 4 so the muxer can restore each frame's ATS header
    ' byte-for-byte.
    Public ReadOnly Property PerFrameExtraConfig As List(Of Byte())
        Get
            Return _perFrameExtra
        End Get
    End Property

    Public Sub New(bw As BinaryWriter)
        _bw = bw
        _sizeProbe = New Byte(3) {}
    End Sub

    Public Sub Append(source As Byte(), offset As Integer, count As Integer)
        Dim p As Integer = offset
        Dim remain As Integer = count

        ' haven't seen the ATS header yet
        ' buffer bytes into the 4-byte probe until we can parse frame size
        If _frameSize = 0 Then
            While remain > 0 AndAlso _sizeProbeFilled < 4
                _sizeProbe(_sizeProbeFilled) = source(p)
                _sizeProbeFilled += 1
                p += 1
                remain -= 1
            End While

            If _sizeProbeFilled = 4 Then
                Dim sync As Integer = (CInt(_sizeProbe(0)) << 8) Or _sizeProbe(1)
                If sync <> SyncWord Then
                    Throw New InvalidDataException(
                        $"ATRAC3plus stream doesn't start with sync word 0x0FD0; got 0x{sync:X4}. " &
                        "Extraction of this stream is unsupported.")
                End If
                Dim data As Integer = (CInt(_sizeProbe(2)) << 8) Or _sizeProbe(3)
                Dim n As Integer = data And &H3FF
                If n >= &H200 Then
                    Throw New InvalidDataException(
                        $"ATRAC3plus ATS header advertises out-of-range frame size (n=0x{n:X}).")
                End If
                _frameSize = (n + 1) * 8 + AtsHeaderSize
                _frame = New Byte(_frameSize - 1) {}
                ' seed the frame buffer with the 4 probe bytes we've already consumed
                Buffer.BlockCopy(_sizeProbe, 0, _frame, 0, 4)
                _filled = 4
            End If
        End If

        ' accumulate whole frames, emit each minus 8-byte ATS header
        ' (frame_size - 8) bytes of raw_data_frame content
        While remain > 0 AndAlso _frameSize > 0
            Dim need As Integer = _frameSize - _filled
            Dim take As Integer = If(need <= remain, need, remain)
            Buffer.BlockCopy(source, p, _frame, _filled, take)
            _filled += take
            p += take
            remain -= take
            If _filled = _frameSize Then
                ' snapshot THIS frame's bytes 4-7 (per-frame - not just the first)
                Dim ec(3) As Byte
                Buffer.BlockCopy(_frame, 4, ec, 0, 4)
                _perFrameExtra.Add(ec)
                _bw.Write(_frame, AtsHeaderSize, _frameSize - AtsHeaderSize)
                FramesWritten += 1
                _filled = 0
            End If
        End While
    End Sub

End Class

'
' PAMF LPCM samples are big-endian
' standard WAV expects little-endian, samples can be 16-bit or 24-bit and a single sample may straddle a PES boundary
'
' mono streams are stored as two channels (mono + silent dummy channel)
Friend Class LpcmBeToLeSwapper

    Private ReadOnly _bw As BinaryWriter
    Private ReadOnly _bytesPerSample As Integer   ' bytes for one channel sample (2 = 16-bit, 3 = 24-bit)
    Private ReadOnly _channelsOnWire As Integer   ' channels present in the PAMF (padded to even for mono)
    Private ReadOnly _channelsToWrite As Integer  ' channels the WAV output should carry (unpadded)
    Private ReadOnly _sample As Byte()
    Private _filled As Integer = 0
    Private _channelIdx As Integer = 0            ' which channel of the current sample-time is currently being filled
    Private _bytesWritten As Long = 0

    ' total bytes emitted to WAV so far, differs from input length when channelsOnWire > channelsToWrite (mono padding)
    Public ReadOnly Property BytesWritten As Long
        Get
            Return _bytesWritten
        End Get
    End Property

    ' overload for existing call sites, assumes no padding
    Public Sub New(bw As BinaryWriter, bytesPerSample As Integer)
        Me.New(bw, bytesPerSample, channelsOnWire:=1, channelsToWrite:=1)
    End Sub

    Public Sub New(bw As BinaryWriter, bytesPerSample As Integer,
                   channelsOnWire As Integer, channelsToWrite As Integer)
        If bytesPerSample <> 2 AndAlso bytesPerSample <> 3 Then
            Throw New ArgumentException(
                $"Unsupported LPCM sample size: {bytesPerSample} bytes/sample. " &
                "PAMF LPCM is 16-bit or 24-bit.")
        End If
        If channelsOnWire < channelsToWrite Then
            Throw New ArgumentException("channelsOnWire must be >= channelsToWrite")
        End If
        _bw = bw
        _bytesPerSample = bytesPerSample
        _channelsOnWire = channelsOnWire
        _channelsToWrite = channelsToWrite
        _sample = New Byte(bytesPerSample - 1) {}
    End Sub

    Public Sub Append(source As Byte(), offset As Integer, count As Integer)
        Dim p As Integer = offset
        Dim remain As Integer = count
        While remain > 0
            Dim need As Integer = _bytesPerSample - _filled
            Dim take As Integer = If(need <= remain, need, remain)
            Buffer.BlockCopy(source, p, _sample, _filled, take)
            _filled += take
            p += take
            remain -= take
            If _filled = _bytesPerSample Then
                ' one channel sample assembled.
                ' emit if its a "real" channel, skip it if its a dummy channel
                If _channelIdx < _channelsToWrite Then
                    For i As Integer = _bytesPerSample - 1 To 0 Step -1
                        _bw.Write(_sample(i))
                    Next
                    _bytesWritten += _bytesPerSample
                End If
                _filled = 0
                _channelIdx += 1
                If _channelIdx = _channelsOnWire Then _channelIdx = 0
            End If
        End While
    End Sub

End Class