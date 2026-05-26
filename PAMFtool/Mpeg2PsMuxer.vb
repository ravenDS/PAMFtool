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
    End Class

    Public Class PamfMuxStream
        Public Property Index As Integer
        Public Property Codec As PamfStreamType
        Public Property PesStreamId As Byte
        Public Property SubStreamId As Byte
        Public Property NumChannels As Byte    ' for LPCM PES sub-header
        Public Property BitsPerSample As Byte  ' for LPCM PES sub-header
        Public ReadOnly Property AuQueue As New Queue(Of AccessUnit)()
        Public Property LastQueuedPts As Long
        Public ReadOnly Property EpEntries As New List(Of EpEntry)()

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
        Public Const InitialScr As Long = 30030L

        Public Property MuxRateBps As Integer = 24_000_000
        ' P-STD buffer sizes (1024-byte units). These appear in system_header per-stream entries & in PAMF codec_info p_std_buffer field, so they must match!
        ' Values for PAMF: 1505 KB for AVC video, 20 KB for audio
        Public Property VideoPstdBufferSize As Integer = 1505
        Public Property AudioPstdBufferSize As Integer = 20
        Public ReadOnly Property Streams As New List(Of PamfMuxStream)()
        Public Property PayloadStartOffset As Long = 0

        Private ReadOnly _splitState As New Dictionary(Of Integer, Integer)()

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
                Case PamfStreamType.AVC, PamfStreamType.MPEG2Video
                    Dim vc As Integer = 0
                    For Each x In Streams
                        If x.IsVideo Then vc += 1
                    Next
                    s.PesStreamId = CByte(&HE0 Or vc)
                    s.SubStreamId = 0
                Case PamfStreamType.ATRAC3plus
                    Dim ch As Integer = CountStreamsOfCodec(PamfStreamType.ATRAC3plus)
                    s.PesStreamId = &HBD
                    s.SubStreamId = CByte(ch)
                Case PamfStreamType.AC3
                    Dim ch As Integer = CountStreamsOfCodec(PamfStreamType.AC3)
                    s.PesStreamId = &HBD
                    s.SubStreamId = CByte(&H30 Or ch)
                Case PamfStreamType.LPCM
                    Dim ch As Integer = CountStreamsOfCodec(PamfStreamType.LPCM)
                    s.PesStreamId = &HBD
                    s.SubStreamId = CByte(&H40 Or ch)
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

            EmitInitialPack(output, scr, muxRateUnits)

            While HasMoreData()
                Dim s As PamfMuxStream = PickNextStream()
                If s Is Nothing Then Exit While

                Dim packStart As Long = output.Position
                WritePackHeader(output, scr, 0, muxRateUnits)

                Dim spaceLeft As Integer = SectorSize - PackHeaderLen

                ' RAP marker: if this pack starts a video AU that is random access point, emit system_header + private_stream_2(stream_id)
                ' so demuxer flags rap=true on that AU
                ' private_stream_2 PES is only parsed when it appears after a system_header
                If s.IsVideo AndAlso s.AuQueue.Count > 0 AndAlso
                   Not _splitState.ContainsKey(s.Index) Then
                    Dim head As AccessUnit = s.AuQueue.Peek()
                    If head.IsRandomAccessPoint Then
                        WriteSystemHeader(output, muxRateUnits, VideoPstdBufferSize)
                        WriteRapMarker(output, s.PesStreamId)
                        spaceLeft = SectorSize - CInt(output.Position - packStart)
                    End If
                End If

                EmitOnePesIntoSector(output, s, spaceLeft, packStart)

                ' advance SCR by elapsed bytes at MuxRateBps
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
        End Sub

        ' private_stream_2 (0xBF) PES whose payload first 2 bytes are a video stream identifier
        ' the byte AT payload offset 1 carries the actual stream_id like 0xE0
        Private Sub WriteRapMarker(out As Stream, videoStreamId As Byte)
            ' PES header: 00 00 01 BF + length(2) + payload(4)
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_PrivateStream2)
            out.WriteByte(0) : out.WriteByte(4)
            out.WriteByte(0)               ' high byte of stream_id field
            out.WriteByte(videoStreamId)   ' channel = stream_id & 0xF
            out.WriteByte(&HFF) : out.WriteByte(&HFF)
        End Sub

        Private Sub EmitInitialPack(out As Stream, scr As Long, muxRateUnits As Integer)
            Dim packStart As Long = out.Position
            WritePackHeader(out, scr, 0, muxRateUnits)
            WriteSystemHeader(out, muxRateUnits, VideoPstdBufferSize)
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
                If s.IsAudio Then key -= AudioLeadTicks
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

            Dim payloadFit As Integer = availBytes - VideoPesHeaderLen
            If payloadFit <= 0 Then
                ' No room at all - fill bytes
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

            ' PES occupies entire sector remainder, size pes_packet_length accordingly and pad with 0xFF after data
            WriteVideoPesHeader(out, s.PesStreamId, payloadFit, au.Pts, au.Dts,
                                VideoPstdBufferSize)
            out.Write(au.Data, consumed, chunkLen)
            For Each e In extras
                out.Write(e.Data, 0, e.Data.Length)
            Next
            Dim stuff As Integer = payloadFit - chunkLen - extrasBytes
            For i As Integer = 0 To stuff - 1
                out.WriteByte(&HFF)
            Next

            If auFullyConsumed Then
                s.AuQueue.Dequeue()
                _splitState.Remove(s.Index)
                For Each e In extras
                    s.AuQueue.Dequeue()
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
            ' Audio packing
            _splitState.Remove(s.Index)

            Dim isLpcm As Boolean = (s.Codec = PamfStreamType.LPCM)
            Dim lpcmExtraBytes As Integer = If(isLpcm, 13, 0)

            Dim availForData As Integer = availBytes - AudioPesHeaderLen - lpcmExtraBytes
            If availForData <= 0 Then
                EmitOnlyPaddingForRest(out, availBytes)
                Return
            End If

            ' step 1: continuation from previous PES of this stream
            Dim contBytes As Integer = 0
            Dim contData As Byte() = Nothing
            Dim contStart As Integer = 0
            If Not isLpcm AndAlso _audioPartial.ContainsKey(s.Index) Then
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

            ' if we just finished a continuation, dequeue that AU before packing whole ones
            If contBytes > 0 AndAlso contStart + contBytes >= contData.Length Then
                s.AuQueue.Dequeue()
                _audioPartial.Remove(s.Index)
            End If

            ' for LPCM first AU is the one caller pre-peeked
            ' for non-LPCM with no continuation, same
            If contBytes = 0 OrElse Not _audioPartial.ContainsKey(s.Index) Then
                While s.AuQueue.Count > 0
                    Dim head As AccessUnit = s.AuQueue.Peek()
                    If packedBytes + head.Data.Length > spaceForWholeAus Then Exit While
                    s.AuQueue.Dequeue()
                    packed.Add(head)
                    packedBytes += head.Data.Length
                End While
            End If

            ' step 3: optionally split the next AU to fill the sector (non-LPCM only)
            Dim splitAu As AccessUnit = Nothing
            Dim splitBytes As Integer = 0
            If Not isLpcm AndAlso s.AuQueue.Count > 0 Then
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

            ' sub-header write 4 bytes, LPCM append 13 more for au_specific_info_buf
            Dim totalAuBytes As Integer = contBytes + packedBytes + splitBytes
            Dim subHeaderAndPayload As Integer = AudioSubHeaderLen + lpcmExtraBytes + totalAuBytes
            WriteAudioPesHeader(out, subHeaderAndPayload, ptsAu.Pts, AudioPstdBufferSize)

            Dim firstAuPtr As UShort = CUShort(If(isLpcm, 13, contBytes))
            ' numFrameHeaders: Sony always writes 0 here in real AA_CR.
            ' For AT3+/AC3 the demuxer read u32 at sub-header [0..3] and mask with 0xFFFF, so byte 1 is discarded
            ' for LPCM the mask is 0x7FF also discarding byte 1
            WriteAudioSubHeader(out, s.SubStreamId, 0, firstAuPtr)

            If isLpcm Then
                Dim extra(12) As Byte
                extra(0) = CByte(s.NumChannels)
                extra(1) = 1                          ' 48 kHz code
                extra(2) = CByte(s.BitsPerSample)
                out.Write(extra, 0, 13)
            End If

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
            End If

            ' pad leftover sector bytes (rare)
            Dim used As Integer = AudioPesHeaderLen + lpcmExtraBytes + totalAuBytes
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