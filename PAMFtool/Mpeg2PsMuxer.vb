' Mpeg2PsMuxer.vb - github.com/ravenDS/PAMFtool
' Multi-stream MPEG-2 Program Stream muxer for PAMF

Imports System.IO
Imports PAMFtool.PamfMux.Mpeg2PsPrimitives

Namespace PamfMux

    Public Class AccessUnit
        Public Property StreamIndex As Integer
        Public Property Data As Byte()
        Public Property Pts As Long
        Public Property Dts As Long
        Public Property IsRandomAccessPoint As Boolean
        ' true for AUs that are used as REFERENCE for other frames (I, P slices in AVC / any anchor picture in M2V)
        ' false for non-reference B slices
        ' detected from AVC nal_ref_idc of the first VCL NALU: 0 = non-ref, non-zero = ref
        ' used only to set bit 7 of the flag byte in private_stream_2 marker
        Public Property IsReferenceFrame As Boolean = True
        ' video-only metadata, populated by RegisterAndQueueM2v, used to fill private_stream_2 emitted at each AU start pack
        Public Property VideoPictureIndex As Integer     ' 0-based, first picture in stream = 0
        Public Property VideoVbvDelay As UShort          ' from MPEG-2 picture_header
    End Class

    Public Class PamfMuxStream
        Public Property Index As Integer
        Public Property Codec As PamfStreamType
        Public Property PesStreamId As Byte
        Public Property SubStreamId As Byte
        Public Property NumChannels As Byte    ' for LPCM PES sub-header
        Public Property BitsPerSample As Byte  ' for LPCM PES sub-header
        ' P-STD buffer size (1024-byte) declared in this stream system_header + in video/audio PES ext for this stream
        ' set by AddXxxStream to codec usual value (AVC=1505, M2V=546, LPCM=128, AT3+/AC-3=20)
        Public Property PstdBufferSize As Integer
        Public ReadOnly Property AuQueue As New Queue(Of AccessUnit)()
        Public Property LastQueuedPts As Long
        Public ReadOnly Property EpEntries As New List(Of EpEntry)()
        Public Property NextAuEmitIndex As Integer

        Public ReadOnly Property IsVideo As Boolean
            Get
                Return Codec = PamfStreamType.AVC OrElse Codec = PamfStreamType.MPEG2Video
            End Get
        End Property

        Public ReadOnly Property IsAudio As Boolean
            Get
                Return Codec = PamfStreamType.ATRAC3plus _
                    OrElse Codec = PamfStreamType.AC3 _
                    OrElse Codec = PamfStreamType.LPCM
            End Get
        End Property
    End Class

    Public Class EpEntry
        Public Property Pts As Long
        Public Property ByteOffset As Long
    End Class

    Public Class Mpeg2PsMuxer

        Public Const PtsClockHz As Long = 90000L
        Public Const AudioLeadTicks As Long = 9000L
        Public Const InitialScr As Long = 30L

        ' Sony PAMF encodes use mux_rate_bound = 48 Mbps
        ' (0x1D4C0 in the PAMF sequence-info header, unit is 50 bytes/s so 120000 * 50 * 8 = 48000000)
        ' this is the value written into the pack_header mux_rate field - the physical delivery capacity of the transport
        Public Property MuxRateBps As Integer = 48_000_000

        ' rate at which SCR advances between packs, independent of MuxRateBps 
        Public Property EffectiveDeliveryBps As Integer = 48_000_000

        ' private_stream_2 emission cadence for AVC RAPs
        ' when > 0, emit the 66-byte marker every N AUs carrying an N-frame-size lookahead.
        ' when 0, emit only the legacy 4-byte AVC RAP marker at every IDR (older behavior, kept for compatibility)
        Public Property Ps2FramesPerBlock As Integer = 12
        ' Legacy: audio P-STD used for non-LPCM audio when a stream doesn't have its own.
        ' Video P-STD now lives per-stream on PamfMuxStream.PstdBufferSize.
        Public Property AudioPstdBufferSize As Integer = 20
        Public ReadOnly Property Streams As New List(Of PamfMuxStream)()
        Public Property PayloadStartOffset As Long = 0

        Private ReadOnly _splitState As New Dictionary(Of Integer, Integer)()

        ' offsets (in output stream) of every private_stream_2 tag we emit
        ' populated in WritePackedStream, used at the end to patch each tag bytes 2-3 (sectors_to_next - 1) when the actual packing is known
        Private ReadOnly _ps2Offsets As New List(Of Long)()

        ' bytes consumed of the current AU that hasn't been fully emitted (next goes into the next audio PES of the same stream)
        Private ReadOnly _audioPartial As New Dictionary(Of Integer, Integer)()

        Public Function AddStream(codec As PamfStreamType) As PamfMuxStream
            ' pamfTypeChannelToStream rules:
            '
            ' AVC / M2V : stream_id = 0xE0 | ch         sub = 0
            ' ATRAC3+   : stream_id = 0xBD              sub = ch
            ' AC-3      : stream_id = 0xBD              sub = 0x30 | ch
            ' LPCM      : stream_id = 0xBD              sub = 0x40 | ch
            ' User data : stream_id = 0xBD              sub = 0x20 | ch
            '
            ' ch counts within each codec (not across all audio streams) so the first AT3+ track is ch=0 even if AC-3 is before
            Dim s As New PamfMuxStream() With {
                .Index = Streams.Count,
                .Codec = codec
            }
            Select Case codec
                Case PamfStreamType.AVC
                    Dim vc As Integer = 0
                    For Each x In Streams
                        If x.IsVideo Then vc += 1
                    Next
                    s.PesStreamId = CByte(&HE0 Or vc)
                    s.SubStreamId = 0
                    s.PstdBufferSize = 1505   ' AVC 720p / 1080i
                Case PamfStreamType.MPEG2Video
                    Dim vc As Integer = 0
                    For Each x In Streams
                        If x.IsVideo Then vc += 1
                    Next
                    s.PesStreamId = CByte(&HE0 Or vc)
                    s.SubStreamId = 0
                    s.PstdBufferSize = 546    ' M2V 720p (Sony reference)
                Case PamfStreamType.ATRAC3plus
                    Dim ch As Integer = CountStreamsOfCodec(PamfStreamType.ATRAC3plus)
                    s.PesStreamId = &HBD
                    s.SubStreamId = CByte(ch)
                    s.PstdBufferSize = 20
                Case PamfStreamType.AC3
                    Dim ch As Integer = CountStreamsOfCodec(PamfStreamType.AC3)
                    s.PesStreamId = &HBD
                    s.SubStreamId = CByte(&H30 Or ch)
                    s.PstdBufferSize = 20
                Case PamfStreamType.LPCM
                    Dim ch As Integer = CountStreamsOfCodec(PamfStreamType.LPCM)
                    s.PesStreamId = &HBD
                    s.SubStreamId = CByte(&H40 Or ch)
                    s.PstdBufferSize = 128
                Case Else
                    Throw New ArgumentException("Unsupported codec for muxing: " & codec.ToString())
            End Select
            Streams.Add(s)
            Return s
        End Function

        Private Function CountStreamsOfCodec(c As PamfStreamType) As Integer
            Dim n As Integer = 0
            For Each x In Streams
                If x.Codec = c Then n += 1
            Next
            Return n
        End Function

        Public Sub QueueAu(stream As PamfMuxStream, au As AccessUnit)
            au.StreamIndex = stream.Index
            stream.AuQueue.Enqueue(au)
            If au.Pts > stream.LastQueuedPts Then stream.LastQueuedPts = au.Pts
        End Sub

        Public Sub WritePackedStream(output As Stream)
            PayloadStartOffset = output.Position
            Dim scr As Long = InitialScr
            Dim muxRateUnits As Integer = MuxRateBps \ 8 \ 50
            _ps2Offsets.Clear()

            If Not HasMpeg2VideoStream() Then
                Dim initialVideoPstd As Integer = FirstVideoPstd()
                EmitInitialPack(output, scr, muxRateUnits, initialVideoPstd)
            End If

            While HasMoreData()
                Dim s As PamfMuxStream = PickNextStream()
                If s Is Nothing Then Exit While

                ' SCR pacing: when EffectiveDeliveryBps < MuxRateBps we want bytes to arrive on the playback timeline
                ' we anchor SCR directly to the AU being emitted:
                '
                '  at every AU-start pack: SCR = max(prev+small, AU.DTS - lead)
                '  between AU-starts:      SCR = prev + small  (fast advance)
                '
                ' buffer_lead = 90000 ticks (1 s) matches std_delay_bound
                Dim pacing As Boolean = (EffectiveDeliveryBps > 0 AndAlso EffectiveDeliveryBps < MuxRateBps)
                If pacing AndAlso Not _splitState.ContainsKey(s.Index) AndAlso s.AuQueue.Count > 0 Then
                    Dim head As AccessUnit = s.AuQueue.Peek()
                    Dim target As Long = head.Dts - PtsClockHz
                    If target > scr Then scr = target
                End If

                Dim packStart As Long = output.Position
                WritePackHeader(output, scr, 0, muxRateUnits)

                Dim spaceLeft As Integer = SectorSize - PackHeaderLen

                If s.IsVideo AndAlso s.AuQueue.Count > 0 AndAlso
                   Not _splitState.ContainsKey(s.Index) Then
                    Dim head As AccessUnit = s.AuQueue.Peek()
                    If s.Codec = PamfStreamType.MPEG2Video Then
                        WriteSystemHeader(output, muxRateUnits, s.PstdBufferSize)
                        Dim ps2Off As Long = output.Position
                        WriteSonyPictureMarker(output, s.PesStreamId,
                                               head.VideoPictureIndex = 0, head.VideoVbvDelay)
                        _ps2Offsets.Add(ps2Off)
                        spaceLeft = SectorSize - CInt(output.Position - packStart)
                    ElseIf Ps2FramesPerBlock > 0 AndAlso s.NextAuEmitIndex Mod Ps2FramesPerBlock = 0 Then
                        WriteSystemHeader(output, muxRateUnits, s.PstdBufferSize)
                        Dim n As Integer = Ps2FramesPerBlock
                        Dim sizes(n - 1) As Integer
                        Dim refs(n - 1) As Boolean
                        Dim k As Integer = 0
                        For Each au In s.AuQueue
                            If k >= n Then Exit For
                            sizes(k) = au.Data.Length
                            refs(k) = au.IsReferenceFrame
                            k += 1
                        Next
                        WriteAvcRapMarkerSony(output, s.PesStreamId, n, sizes, refs)
                        spaceLeft = SectorSize - CInt(output.Position - packStart)
                    ElseIf Ps2FramesPerBlock = 0 AndAlso head.IsRandomAccessPoint Then
                        ' legacy 4-byte stub, emitted at every IDR AU
                        WriteSystemHeader(output, muxRateUnits, s.PstdBufferSize)
                        WriteRapMarker(output, s.PesStreamId)
                        spaceLeft = SectorSize - CInt(output.Position - packStart)
                    End If
                End If

                EmitOnePesIntoSector(output, s, spaceLeft, packStart)

                ' advance at MuxRateBps (fast, 30 ticks/pack)
                ' if pacing is on, the AU-start jump above does the actual playback-alignment work
                Dim packBytes As Long = output.Position - packStart
                scr += (packBytes * PtsClockHz * 8L) \ MuxRateBps
                If scr >= (1L << 33) Then scr -= (1L << 33)
            End While

            ' MPEG_program_end_code, pad to sector boundary
            WriteProgramEnd(output)
            Dim modPos As Long = output.Position Mod SectorSize
            If modPos > 0 Then
                Dim padLen As Integer = CInt(SectorSize - modPos)
                If padLen >= 7 Then
                    WritePaddingStream(output, padLen)
                Else
                    For i As Integer = 0 To padLen - 1
                        output.WriteByte(&HFF)
                    Next
                End If
            End If

            ' patch each private_stream_2
            If _ps2Offsets.Count > 0 Then
                Dim endPos As Long = output.Position
                For i As Integer = 0 To _ps2Offsets.Count - 1
                    Dim curOff As Long = _ps2Offsets(i)
                    Dim nextOff As Long = If(i + 1 < _ps2Offsets.Count,
                                             _ps2Offsets(i + 1), endPos)
                    Dim gapSectors As Long = (nextOff - curOff) \ SectorSize
                    Dim val As Integer = CInt(Math.Max(0L, Math.Min(gapSectors - 1L, &HFFFFL)))
                    output.Position = curOff + 8   ' skip 4 SC + 2 length + 2 (payload[0..1])
                    output.WriteByte(CByte((val >> 8) And &HFF))
                    output.WriteByte(CByte(val And &HFF))
                Next
                output.Position = endPos
            End If
        End Sub

        ' Restored from the previously-shipping muxer for AVC compatibility: at each
        ' video RAP we emit a system_header followed by a small 4-byte private_stream_2
        ' payload that identifies the video stream. The tool's own demuxer doesn't
        ' parse this today; it's here because Sony PAMF hardware expects it (removing
        ' it entirely produced files the PS3 refused to play back).

        ' private_stream_2 marker for AVC RAPs:
        '
        ' N frames per block
        ' payload = 18 + 4*N bytes (66 for N=12, 134 for N=29, 58 for N=10)
        ' emitted every N AUs

        ' carries frame-size lookahead so the hardware CABAC decoder can pre-allocate resources
        '
        ' payload layout:
        '   [0]      0x01 (fixed record marker)
        '   [1]      video_stream_id (0xE0 for the first video stream)
        '   [2..3]   BE u16 - cumulative_bytes_thru_frame_0 / 2048
        '   [4..5]   BE u16 - cumulative_bytes_thru_frame_1 / 2048
        '   [6..7]   BE u16 - cumulative_bytes_thru_frame_2 / 2048
        '   [8..9]   BE u16 - cumulative_bytes_thru_frame_3 / 2048
        '   [10..13] 00 00 00 00 (reserved)
        '   [14..15] BE u16 - value 4*N + 2 (semantic unknown, formula holds across all three observed N values 12,29,10)
        '   [16..17] BE u16 - frame count N in this block
        '   [18]     00
        '   [19..]   N groups of 4 bytes, one per frame in the next N-frame window (in decode order).
        '            each group encodes a 24-bit unsigned frame size (bytes of the AU as it appears in the ES)
        '            plus a reference-frame flag:
        '                [ (is_ref ? 0x80 : 0x00) | ((size >> 16) & 3),
        '                  (size >> 8) & 0xFF,
        '                  size & 0xFF,
        '                  0x00 ]
        '            capped to 0x3FFFFF (~4 MB)
        Private Sub WriteAvcRapMarkerSony(out As Stream, videoStreamId As Byte,
                                           framesPerBlock As Integer,
                                           frameSizes As Integer(),
                                           isReference As Boolean())
            Dim payloadLen As Integer = 18 + 4 * framesPerBlock
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_PrivateStream2)
            out.WriteByte(CByte((payloadLen >> 8) And &HFF))
            out.WriteByte(CByte(payloadLen And &HFF))

            ' header (18 bytes: indices 0..17)
            out.WriteByte(&H1)                          ' [0]
            out.WriteByte(videoStreamId)                ' [1]

            ' cumulative sector offsets thru frames 0..3.
            ' if fewer than 4 frames are available (end of stream), remaining offsets stay 0
            Dim cum As Long = 0
            For k As Integer = 0 To 3
                If k < frameSizes.Length Then
                    cum += frameSizes(k)
                    Dim sectors As Integer = CInt(cum \ 2048L)
                    If sectors > &HFFFF Then sectors = &HFFFF
                    out.WriteByte(CByte((sectors >> 8) And &HFF))
                    out.WriteByte(CByte(sectors And &HFF))
                Else
                    out.WriteByte(0) : out.WriteByte(0)
                End If
            Next

            out.WriteByte(0) : out.WriteByte(0)          ' [10..11]
            out.WriteByte(0) : out.WriteByte(0)          ' [12..13]
            Dim field14 As Integer = 4 * framesPerBlock + 2
            out.WriteByte(CByte((field14 >> 8) And &HFF))
            out.WriteByte(CByte(field14 And &HFF))       ' [14..15] = 4N+2
            out.WriteByte(CByte((framesPerBlock >> 8) And &HFF))
            out.WriteByte(CByte(framesPerBlock And &HFF)) ' [16..17] = N

            ' N frame-size groups (4*N bytes, starting at byte 18)
            ' each group: [leading_zero=0, flag, hi_size, lo_size]
            For k As Integer = 0 To framesPerBlock - 1
                Dim size As Integer = 0
                Dim isRef As Boolean = True
                If k < frameSizes.Length Then size = frameSizes(k)
                If k < isReference.Length Then isRef = isReference(k)
                If size > &H3FFFFF Then size = &H3FFFFF
                Dim flag As Integer = If(isRef, &H80, &H0) Or ((size >> 16) And &H3)
                out.WriteByte(0)
                out.WriteByte(CByte(flag And &HFF))
                out.WriteByte(CByte((size >> 8) And &HFF))
                out.WriteByte(CByte(size And &HFF))
            Next
        End Sub

        Private Sub WriteRapMarker(out As Stream, videoStreamId As Byte)
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_PrivateStream2)
            out.WriteByte(0) : out.WriteByte(4)             ' PES length = 4
            out.WriteByte(0)                                ' payload byte 0
            out.WriteByte(videoStreamId)                    ' payload byte 1 = stream id
            out.WriteByte(&HFF) : out.WriteByte(&HFF)       ' payload bytes 2-3
        End Sub

        Private Sub EmitInitialPack(out As Stream, scr As Long, muxRateUnits As Integer,
                                    videoPstdBufferSize As Integer)
            Dim packStart As Long = out.Position
            WritePackHeader(out, scr, 0, muxRateUnits)
            WriteSystemHeader(out, muxRateUnits, videoPstdBufferSize)
            Dim used As Integer = CInt(out.Position - packStart)
            Dim remaining As Integer = SectorSize - used
            If remaining >= 7 Then
                WritePaddingStream(out, remaining)
            ElseIf remaining > 0 Then
                For i As Integer = 0 To remaining - 1
                    out.WriteByte(&H0)
                Next
            End If
        End Sub

        ' Look up the first video stream's per-stream P-STD buffer size, for the
        ' system_header that goes into sector 0's initial pack. Falls back to 1505
        ' (the old AVC-oriented default) if we can't find a video stream.
        Private Function FirstVideoPstd() As Integer
            For Each s In Streams
                If s.IsVideo Then Return s.PstdBufferSize
            Next
            Return 1505
        End Function

        Private Function HasMpeg2VideoStream() As Boolean
            For Each s In Streams
                If s.Codec = PamfStreamType.MPEG2Video Then Return True
            Next
            Return False
        End Function

        ' Sony private_stream_2 (0xBF) tag
        ' layout (22 bytes):
        '   [0]      0x01                              record marker
        '   [1]      video_stream_id (0xE0 typically)  which video stream this tag is for
        '   [2..3]   sectors_to_next_ps2 - 1           patched by WritePackedStream after mux
        '   [4..9]   FF FF FF FF FF FF                 reserved
        '   [10..13] 00 00 00 00                       reserved (possibly prev-tag offset)
        '   [14..15] 00 06                             const (unknown Sony field)
        '   [16..17] 00 01                             const (unknown Sony field)
        '   [18..19] 0x2002 for the first picture in the stream, 0x2001 otherwise
        '            (sequence-start marker vs regular picture)
        '   [20..21] vbv_delay from the MPEG-2 picture_header (best-effort; Sony's
        '            observed values are ~2.75x larger than the raw picture-header
        '            field, suggesting a mux-buffer-model derived value we don't
        '            reproduce exactly, but the relative variation between pictures
        '            matches so hardware treats it the same way)
        Private Sub WriteSonyPictureMarker(out As Stream,
                                           videoStreamId As Byte,
                                           isFirstPicture As Boolean,
                                           vbvDelay As UShort)
            ' 4-byte start code + 2-byte length + 22-byte payload = 28 bytes total
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_PrivateStream2)
            out.WriteByte(0) : out.WriteByte(&H16)      ' length = 22
            out.WriteByte(&H1)                          ' [0]
            out.WriteByte(videoStreamId)                 ' [1]
            out.WriteByte(0) : out.WriteByte(0)          ' [2..3] patched later
            For i As Integer = 0 To 5
                out.WriteByte(&HFF)                      ' [4..9]
            Next
            For i As Integer = 0 To 3
                out.WriteByte(0)                         ' [10..13]
            Next
            out.WriteByte(0) : out.WriteByte(&H6)        ' [14..15] 0x0006
            out.WriteByte(0) : out.WriteByte(&H1)        ' [16..17] 0x0001
            out.WriteByte(&H20)                          ' [18]
            out.WriteByte(If(isFirstPicture, CByte(&H2), CByte(&H1)))  ' [19]
            out.WriteByte(CByte((CInt(vbvDelay) >> 8) And &HFF))       ' [20]
            out.WriteByte(CByte(CInt(vbvDelay) And &HFF))              ' [21]
        End Sub

        Private Function HasMoreData() As Boolean
            For Each s In Streams
                If s.AuQueue.Count > 0 Then Return True
            Next
            Return False
        End Function

        Private Function PickNextStream() As PamfMuxStream
            Dim best As PamfMuxStream = Nothing
            Dim bestKey As Long = Long.MaxValue
            For Each s In Streams
                If s.AuQueue.Count = 0 Then Continue For
                Dim head As AccessUnit = s.AuQueue.Peek()
                Dim key As Long = head.Pts
                ' only compressed audio (AT3+/AC-3) needs the pre-buffer lead
                If s.IsAudio AndAlso s.Codec <> PamfStreamType.LPCM Then
                    key -= AudioLeadTicks
                End If
                If key < bestKey Then
                    bestKey = key
                    best = s
                End If
            Next
            Return best
        End Function

        ' emit one PES for streams into the available sector space after the pack header
        ' if the AU is too big we emit the head portion now and leave the tail in _splitState for the next pack
        Private Sub EmitOnePesIntoSector(out As Stream,
                                         s As PamfMuxStream,
                                         availBytes As Integer,
                                         packStart As Long)
            Dim au As AccessUnit = s.AuQueue.Peek()

            Dim consumed As Integer = 0
            If _splitState.ContainsKey(s.Index) Then
                consumed = _splitState(s.Index)
            End If
            Dim remaining As Integer = au.Data.Length - consumed

            If s.IsVideo Then
                EmitVideoPesChunk(out, s, au, consumed, remaining, availBytes, packStart)
            Else
                EmitAudioPesChunk(out, s, au, consumed, remaining, availBytes, packStart)
            End If
        End Sub

        Private Sub EmitVideoPesChunk(out As Stream,
                                      s As PamfMuxStream,
                                      au As AccessUnit,
                                      consumed As Integer,
                                      remaining As Integer,
                                      availBytes As Integer,
                                      packStart As Long)
            ' Video packing:
            '   - one data PES per pack
            '   - the PES fills the rest of the sector
            '   - if the AU finishes with room left and the next AU fits,
            '     concatenate it into the same PES
            '   - any remaining bytes stuffed with 0xFF inside PES payload
            Dim isFirstChunk As Boolean = (consumed = 0)
            If isFirstChunk AndAlso au.IsRandomAccessPoint Then
                s.EpEntries.Add(New EpEntry() With {
                    .Pts = au.Pts,
                    .ByteOffset = packStart - PayloadStartOffset
                })
            End If

            ' AU-start PES carries PTS+DTS+P-STD (13-byte extension, total 22B header)
            ' continuation PES carry no timestamps and no P-STD (0-byte extension, total 9B header)
            ' see WriteVideoPesHeaderContinuation
            Dim hdrLen As Integer = If(isFirstChunk, VideoPesHeaderLen, VideoPesHeaderContinuationLen)
            Dim payloadFit As Integer = availBytes - hdrLen
            If payloadFit <= 0 Then
                ' no room at all - fill bytes
                For i As Integer = 0 To availBytes - 1
                    out.WriteByte(&HFF)
                Next
                Return
            End If

            Dim chunkLen As Integer = Math.Min(payloadFit, remaining)
            Dim auFullyConsumed As Boolean = (consumed + chunkLen >= au.Data.Length)

            ' multi-AU packing: only when current AU finishes here and next AU is a whole-fit
            ' do not split the second AU
            Dim extras As New List(Of AccessUnit)()
            Dim extrasBytes As Integer = 0
            If auFullyConsumed Then
                Dim spare As Integer = payloadFit - chunkLen
                While spare > 0
                    Dim n As AccessUnit = PeekAfterHead(s.AuQueue, extras.Count)
                    If n Is Nothing OrElse n.Data.Length > spare Then Exit While
                    extras.Add(n)
                    extrasBytes += n.Data.Length
                    spare -= n.Data.Length
                End While
            End If

            ' size video PES to just the real AU content it carries
            ' any leftover sector space goes into a separate padding_stream packet (0xBE)
            Dim actualPayload As Integer = chunkLen + extrasBytes
            Dim leftover As Integer = payloadFit - actualPayload
            Dim pesPayloadLen As Integer
            Dim inPesStuff As Integer
            Dim emitPaddingStream As Boolean
            Dim paddingStreamLen As Integer
            If leftover >= 7 Then
                pesPayloadLen = actualPayload
                inPesStuff = 0
                emitPaddingStream = True
                paddingStreamLen = leftover
            Else
                pesPayloadLen = payloadFit
                inPesStuff = leftover
                emitPaddingStream = False
                paddingStreamLen = 0
            End If

            If isFirstChunk Then
                WriteVideoPesHeader(out, s.PesStreamId, pesPayloadLen, au.Pts, au.Dts,
                                    s.PstdBufferSize)
            Else
                WriteVideoPesHeaderContinuation(out, s.PesStreamId, pesPayloadLen)
            End If
            out.Write(au.Data, consumed, chunkLen)
            For Each e In extras
                out.Write(e.Data, 0, e.Data.Length)
            Next
            For i As Integer = 0 To inPesStuff - 1
                out.WriteByte(&HFF)
            Next
            If emitPaddingStream Then
                WritePaddingStream(out, paddingStreamLen)
            End If

            If auFullyConsumed Then
                s.AuQueue.Dequeue()
                s.NextAuEmitIndex += 1
                _splitState.Remove(s.Index)
                For Each e In extras
                    s.AuQueue.Dequeue()
                    s.NextAuEmitIndex += 1
                Next
            Else
                _splitState(s.Index) = consumed + chunkLen
            End If
        End Sub

        ' peek at queue index `idxAfterHead` (0 = head, 1 = next, ...).
        Private Function PeekAfterHead(q As Queue(Of AccessUnit), idxAfterHead As Integer) As AccessUnit
            Dim targetIdx As Integer = idxAfterHead + 1
            If q.Count <= targetIdx Then Return Nothing
            Dim arr As AccessUnit() = q.ToArray()
            Return arr(targetIdx)
        End Function

        Private Sub EmitAudioPesChunk(out As Stream,
                                      s As PamfMuxStream,
                                      au As AccessUnit,
                                      consumed As Integer,
                                      remaining As Integer,
                                      availBytes As Integer,
                                      packStart As Long)
            ' audio packing
            _splitState.Remove(s.Index)

            Dim isLpcm As Boolean = (s.Codec = PamfStreamType.LPCM)

            ' PAMF LPCM PES layout:
            '     [0]    sub_stream_id  (0x40..0x4F)
            '
            '     [1]    stream config byte, constant per stream. 0x31 for 48 kHz stereo 16-bit LPCM. Exact layout unknown
            '            observed values:
            '              0x31 : 48 kHz stereo 16-bit
            '            other configs need verification
            '
            '     [2..3] first_access_unit_pointer, high nibble is a constant marker (0x4)
            '            the low 12 bits are the byte offset of the first new-AU boundary in the sample area of this PES
            '            0xFFFF when no new AU begins in this PES (continuation-only)
            '
            '     [4..]  LPCM samples directly

            Dim availForData As Integer = availBytes - AudioPesHeaderLen
            If availForData <= 0 Then
                EmitOnlyPaddingForRest(out, availBytes)
                Return
            End If

            ' step 1: continuation from previous PES of this stream
            ' LPCM is allowed here too: at high sample rate / channel count a 20 ms
            ' LPCM AU can exceed one sector's audio payload budget and must span
            ' multiple PESes, exactly like a large AT3+/AC-3 frame.
            Dim contBytes As Integer = 0
            Dim contData As Byte() = Nothing
            Dim contStart As Integer = 0
            If _audioPartial.ContainsKey(s.Index) Then
                Dim partialConsumed As Integer = _audioPartial(s.Index)
                Dim partialAu As AccessUnit = s.AuQueue.Peek()
                contData = partialAu.Data
                contStart = partialConsumed
                contBytes = Math.Min(contData.Length - partialConsumed, availForData)
            End If

            ' step 2: pack whole AUs that fit in the remaining space
            Dim packed As New List(Of AccessUnit)()
            Dim packedBytes As Integer = 0
            Dim spaceForWholeAus As Integer = availForData - contBytes

            Dim contFinished As Boolean = (contBytes > 0 AndAlso
                                           contStart + contBytes >= contData.Length)

            ' if we just finished a continuation, dequeue that AU before packing whole ones
            If contFinished Then
                s.AuQueue.Dequeue()
                _audioPartial.Remove(s.Index)
            End If

            ' pack whole AUs only when we're not still mid-continuation
            If contBytes = 0 OrElse contFinished Then
                While s.AuQueue.Count > 0
                    Dim head As AccessUnit = s.AuQueue.Peek()
                    If packedBytes + head.Data.Length > spaceForWholeAus Then Exit While
                    s.AuQueue.Dequeue()
                    packed.Add(head)
                    packedBytes += head.Data.Length
                End While
            End If

            ' step 3: optionally split the next AU to fill the sector
            Dim splitAu As AccessUnit = Nothing
            Dim splitBytes As Integer = 0
            If s.AuQueue.Count > 0 Then
                Dim remainingSpace As Integer = spaceForWholeAus - packedBytes
                If remainingSpace > 0 Then
                    splitAu = s.AuQueue.Peek()
                    splitBytes = Math.Min(splitAu.Data.Length, remainingSpace)
                    If splitBytes >= splitAu.Data.Length Then
                        ' whole fit (edge case), treat as a whole AU
                        s.AuQueue.Dequeue()
                        packed.Add(splitAu)
                        packedBytes += splitAu.Data.Length
                        splitAu = Nothing
                        splitBytes = 0
                    End If
                End If
            End If

            If contBytes = 0 AndAlso packed.Count = 0 AndAlso splitAu Is Nothing Then
                ' nothing to emit (even a single AU doesn't fit)
                EmitOnlyPaddingForRest(out, availBytes)
                Return
            End If

            Dim ptsAu As AccessUnit
            If packed.Count > 0 Then
                ptsAu = packed(0)
            ElseIf splitAu IsNot Nothing Then
                ptsAu = splitAu
            Else
                ' continuation-only PES, PTS is from the AU we're continuing
                ptsAu = au
            End If

            Dim totalAuBytes As Integer = contBytes + packedBytes + splitBytes
            Dim subHeaderAndPayload As Integer = AudioSubHeaderLen + totalAuBytes
            WriteAudioPesHeader(out, subHeaderAndPayload, ptsAu.Pts, s.PstdBufferSize)

            ' first_access_unit_pointer:
            '   LPCM  : 0x4000 marker in the high nibble, low 12 bits are the offset of the first new-AU boundary within the sample area
            '           0xFFFF when this PES is purely continuation of a previous AU.
            '   AT3+/AC-3 : straight byte offset (matches the existing tool convention).
            Dim hasNewAu As Boolean = (packed.Count > 0) OrElse (splitAu IsNot Nothing)
            Dim firstAuPtr As UShort
            Dim numFrameHeaders As Byte
            If isLpcm Then
                firstAuPtr = If(hasNewAu, CUShort(&H4000 Or (contBytes And &HFFF)), CUShort(&HFFFF))
                ' Sony's per-stream config byte at sub-header[1].fall back to 0x31 as a reasonable default
                ' needs extending with different LPCM configs
                numFrameHeaders = &H31
            Else
                firstAuPtr = CUShort(contBytes)
                numFrameHeaders = 0
            End If
            WriteAudioSubHeader(out, s.SubStreamId, numFrameHeaders, firstAuPtr)

            ' NO 13-byte extra header for LPCM, samples immediately follow 4-byte audio sub-header, Sony PAMFs do not have anything here

            ' continuation tail, then whole AUs, then split AU head
            If contBytes > 0 Then
                out.Write(contData, contStart, contBytes)
            End If
            For Each a In packed
                out.Write(a.Data, 0, a.Data.Length)
            Next
            If splitAu IsNot Nothing AndAlso splitBytes > 0 Then
                out.Write(splitAu.Data, 0, splitBytes)
                _audioPartial(s.Index) = splitBytes
            ElseIf contBytes > 0 AndAlso Not contFinished Then
                ' continuation-only PES: the head AU still has bytes we haven't emitted;
                ' record how far in we are so the next PES resumes from the right offset
                _audioPartial(s.Index) = contStart + contBytes
            End If

            ' pad leftover sector bytes (rare)
            Dim used As Integer = AudioPesHeaderLen + totalAuBytes
            Dim spareInSector As Integer = availBytes - used
            If spareInSector > 0 Then
                If spareInSector >= 7 Then
                    WritePaddingStream(out, spareInSector)
                Else
                    For i As Integer = 0 To spareInSector - 1
                        out.WriteByte(&HFF)
                    Next
                End If
            End If
        End Sub

        Private Sub EmitOnlyPaddingForRest(out As Stream, availBytes As Integer)
            If availBytes >= 7 Then
                WritePaddingStream(out, availBytes)
            Else
                For i As Integer = 0 To availBytes - 1
                    out.WriteByte(&HFF)
                Next
            End If
        End Sub

    End Class

End Namespace