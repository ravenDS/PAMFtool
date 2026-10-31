' Mpeg2PsMuxer.vb - github.com/ravenDS/PAMFtool
' Multi-stream MPEG-2 Program Stream muxer for PAMF
'
'
'  pack grid    every pack is 2048 bytes; pack k has SCR = SCR0 + k' * slot with slot = 2048 * 8 / mux_rate
'               (9216 27 MHz ticks at 48 Mbps)
'               when nothing may be sent at the next slot, SCR jumps to the first slot at which some stream becomes eligible (it never leaves the grid)
'
'  SCR0         origin + slot * max(1, ceil((E0 - origin) / slot)), E0 = eligibility time of the first pack
'               origin is 0 or 30000 ticks depending on the encoder generation (see ScrOriginTicks)
'
'  lead         a stream may fill a pack only when T - SCR <= std_delay_bound, T = DTS (video) or PTS (audio)
'               of the AU holding the pack's last payload byte (for a short video pack that ends on an AU
'               boundary - block end, or before an AU that opens the next PES - the AU that follows)
'
'  room         after the pack, the stream's buffered ES bytes (delivered minus AUs with T <= SCR) must
'               fit the P-STD buffer declared for the stream
'
'  choice       among eligible streams, the one whose next byte belongs to the AU with the earliest T;
'               video wins ties
'
'  video        one PES per pack, ES bytes back to back across AU boundaries; a ps2 block starts a new pack
'               (system header + private_stream_2 + PES with PTS/DTS/P-STD); the pack closing a block is
'               padded (padding packet, or PES header stuffing when at most 7 bytes are left);
'               inside a block a PTS (+DTS when it differs) goes on the first AU starting in a PES once its
'               DTS is >= 0.5 s after the previous stamp's PTS; when the AU after an AU j that starts further
'               inside a pack would be >= 0.6 s after it, the pack ends right before j (padded) and j opens
'               the next PES
'
'  audio        packs are filled, PTS only when an AU starts in the payload;
'               first_au pointer = offset of
'               that AU (0xFFFF / LPCM 0x47FF when none); a PES without an AU start ends right before
'               the next AU start (stuffing); P-STD extension on the first PES of each audio stream after
'               every video private_stream_2

Imports System.IO
Imports PAMFtool.PamfMux.Mpeg2PsPrimitives

Namespace PamfMux

    Public Class AccessUnit
        Public Property StreamIndex As Integer
        Public Property Data As Byte()
        Public Property Pts As Long                     ' 90 kHz
        Public Property Dts As Long                     ' 90 kHz (audio: same as Pts)
        ' time the buffer model uses, 27 MHz: video = exact DTS, audio = PTS * 300
        Public Property DecodeTime27 As Long
        ' video: random access point
        ' (IDR / recovery-point I picture for AVC, I picture with sequence header for MPEG-2)
        ' gets an EP entry when it starts a block
        Public Property IsRandomAccessPoint As Boolean
        ' video: used as reference (AVC nal_ref_idc <> 0, MPEG-2 I/P picture)
        Public Property IsReferenceFrame As Boolean = True
        ' MPEG-2 video: picture_coding_type (1 I, 2 P, 3 B), 0 for AVC
        Public Property PictureCodingType As Byte
        ' video: a private_stream_2 block starts at this AU
        Public Property StartsBlock As Boolean
    End Class

    Public Class PamfMuxStream
        Public Property Index As Integer
        Public Property Codec As PamfStreamType
        Public Property PesStreamId As Byte
        Public Property SubStreamId As Byte
        Public Property NumChannels As Byte    ' for LPCM PES sub-header
        Public Property BitsPerSample As Byte  ' for LPCM PES sub-header
        ' P-STD buffer size (1024-byte units) declared in PAMF header and in P-STD PES extension
        ' used as buffer bound of scheduler
        Public Property PstdBufferSize As Integer
        Public ReadOnly Property Aus As New List(Of AccessUnit)()
        Public ReadOnly Property EpEntries As New List(Of EpEntry)()

        ' planner state / results
        Friend AuStart As Long()          ' n + 1 entries, AuStart(n) = total ES bytes
        ' video: ES offset of the AU first start code prefix (00 00 01)
        ' decide whether an AU starts in a PES by this position
        Friend Anchor As Long()
        Friend T27 As Long()
        Friend AuEndPack As Integer()
        Friend Blocks As New List(Of VideoBlock)()
        Friend LateAus As Integer
        Friend MaxLate27 As Long
        Friend PeakOccupancy As Long
        Friend MaxLead27 As Long
        Friend Stamps As Integer

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

        Public ReadOnly Property TotalBytes As Long
            Get
                Return If(AuStart Is Nothing, 0L, AuStart(AuStart.Length - 1))
            End Get
        End Property
    End Class

    ' one EP map entry (one per ps2 block that starts at a random access point)
    Public Class EpEntry
        Public Property Pts As Long
        Public Property PackIndex As Long               ' rpnOffset: sectors from the start of the data area
        Public Property IndexN As Integer               ' 1..4 = reference pictures in the block (max 4)
        ' packs from the RAP pack to the pack holding the last byte of the IndexN-th reference picture
        ' (13 bits, reader computes nThRefPictureOffset = (value + 1) * 2048)
        Public Property RefPictureOffset As Integer
    End Class

    Friend Class VideoBlock
        Public FirstAu As Integer
        Public AuCount As Integer
        Public StartOffset As Long
        Public EndOffset As Long
        Public IsRap As Boolean
        Public StartPack As Integer = -1
        Public ReadOnly Refs As New List(Of Integer)()     ' first (up to) 4 reference AUs of the block
        Public ReadOnly Offsets As Integer() = {&HFFFF, &HFFFF, &HFFFF, &HFFFF}
    End Class

    Friend Structure PackPlan
        Public Scr27 As Long
        Public StreamIdx As Integer
        Public Pos As Long              ' ES offset of the first payload byte
        Public Len As Integer           ' ES payload bytes
        Public Rap As Boolean           ' video block-start pack (system header + private_stream_2)
        Public BlockIdx As Integer
        Public HasPts As Boolean
        Public HasDts As Boolean
        Public HasPstd As Boolean
        Public TsAu As Integer          ' AU whose PTS/DTS is written
        Public Stuffing As Integer      ' PES header stuffing bytes (0..7)
        Public Padding As Integer       ' padding packet length (0 or >= 8)
        Public Pointer As Integer       ' audio: first_au pointer field as written
    End Structure

    Public Class Mpeg2PsMuxer

        Public Const PtsClockHz As Long = 90000L
        Public Const MaxPs2Pictures As Integer = 492          ' ps2 of 18 + 4N bytes must fit the RAP pack
        ' mid-block video timestamps, measured against PTS of the previous video stamp:
        '   first AU starting in a PES is stamped when its DTS is >= StampIntervalTicks (0.5 s) after it
        '   (43544 < x <= 45045)
        '
        '   a pack is ended before an AU j starting further inside it when DTS of AU after j is >= CutIntervalTicks (0.6 s) after it
        '   (52552 < x <= 54054)
        '   so that j opens next PES and can carry the stamp
        Public Const StampIntervalTicks As Long = 45000L
        Public Const CutIntervalTicks As Long = 54000L
        Public Const MaxPstdKb As Integer = 8191              ' 13-bit P-STD_buffer_size
        ' fill up to 7 spare bytes with PES header stuffing and use a padding packet from 8 bytes on
        Public Const MinPaddingPacket As Integer = 8

        ' PAMF only accept these rates:
        '   48000 kbps   (default, used by most PAMFs)
        '   24000 kbps
        '   12000 kbps   (typical for SD content)
        ' pack_header mux_rate field encodes at 50-byte/s units
        ' 48 Mbps = 120000 units (0x1D4C0) stored also in PAMF sequence-info header
        Public Property MuxRateBps As Integer = 48_000_000

        Public Shared ReadOnly Property AllowedMuxRatesBps As Integer() =
            {12_000_000, 24_000_000, 48_000_000}

        Public Shared Function IsAllowedMuxRate(bps As Integer) As Boolean
            For Each v In AllowedMuxRatesBps
                If bps = v Then Return True
            Next
            Return False
        End Function

        ' std_delay_bound (90 kHz): the lead limit of the scheduler, also written into PAMF header
        Public Property StdDelayTicks As Long = 90000L

        ' SCR0 origin (90 kHz). Files show two encoder generations:
        '   0      : AVC High-profile game files (SCR0 = 30.72 at 48 Mbps for a normal first pack)
        '   30000  : everything else, AVC Main and MPEG-2 (SCR0 = 30030.72 at 48 Mbps)
        Public Property ScrOriginTicks As Long = 30000L

        ' True: append a sector with MPEG_program_end_code (+ padding) after the last pack
        Public Property WriteEndSector As Boolean = False

        Public ReadOnly Property Streams As New List(Of PamfMuxStream)()

        Private ReadOnly _plan As New List(Of PackPlan)()

        Public ReadOnly Property PackCount As Integer
            Get
                Return _plan.Count
            End Get
        End Property

        Public ReadOnly Property SlotTicks27 As Long
            Get
                Return CLng(SectorSize) * 8L * 27_000_000L \ CLng(MuxRateBps)
            End Get
        End Property

        Public Property ScrJumps As Integer

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
                    s.PstdBufferSize = If(codec = PamfStreamType.AVC, 1505, 546)   ' replaced by the stream's HRD / VBV value
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
            stream.Aus.Add(au)
        End Sub

        ' planning

        Private Class Candidate
            Public Plan As PackPlan
            Public EndPos As Long
            Public LeadAu As Integer
            Public Elig27 As Long
            Public Valid As Boolean
        End Class

        Private _pos As Long()
        Private _blockIdx As Integer()
        Private _lastStampPts As Long()
        Private _needExt As Boolean()
        Private _cand As Candidate()
        Private _endAuPtr As Integer()

        ' build per-stream arrays and video blocks, throws on inputs that planner cannot handle
        Private Sub Prepare()
            For Each s In Streams
                Dim n As Integer = s.Aus.Count
                If n = 0 Then Throw New InvalidDataException($"stream {s.Index} ({s.Codec}) has no access units")
                ReDim s.AuStart(n)
                ReDim s.T27(n - 1)
                ReDim s.AuEndPack(n - 1)
                Dim off As Long = 0
                For k As Integer = 0 To n - 1
                    Dim au As AccessUnit = s.Aus(k)
                    If au.Data Is Nothing OrElse au.Data.Length = 0 Then
                        Throw New InvalidDataException($"stream {s.Index}: empty access unit {k}")
                    End If
                    s.AuStart(k) = off
                    off += au.Data.Length
                    s.T27(k) = au.DecodeTime27
                    If k > 0 AndAlso s.T27(k) < s.T27(k - 1) Then
                        Throw New InvalidDataException($"stream {s.Index}: decode times are not monotonic at AU {k}")
                    End If
                    s.AuEndPack(k) = -1
                Next
                s.AuStart(n) = off
                ReDim s.Anchor(n - 1)
                For k As Integer = 0 To n - 1
                    Dim d As Byte() = s.Aus(k).Data
                    Dim a As Integer = 0
                    If s.IsVideo Then
                        While a + 2 < d.Length AndAlso a < 8 AndAlso Not (d(a) = 0 AndAlso d(a + 1) = 0 AndAlso d(a + 2) = 1)
                            a += 1
                        End While
                        If a + 2 >= d.Length OrElse a >= 8 Then a = 0
                    End If
                    s.Anchor(k) = s.AuStart(k) + a
                Next
                s.EpEntries.Clear()
                s.Blocks.Clear()
                If s.IsVideo Then
                    s.Aus(0).StartsBlock = True
                    Dim blk As VideoBlock = Nothing
                    For k As Integer = 0 To n - 1
                        If s.Aus(k).StartsBlock OrElse blk Is Nothing OrElse blk.AuCount >= MaxPs2Pictures Then
                            If blk IsNot Nothing Then blk.EndOffset = s.AuStart(k)
                            blk = New VideoBlock() With {
                                .FirstAu = k, .StartOffset = s.AuStart(k),
                                .IsRap = s.Aus(k).IsRandomAccessPoint
                            }
                            s.Blocks.Add(blk)
                        End If
                        blk.AuCount += 1
                        If s.Aus(k).IsReferenceFrame AndAlso blk.Refs.Count < 4 Then blk.Refs.Add(k)
                    Next
                    blk.EndOffset = s.AuStart(n)
                End If
            Next
        End Sub

        ' smallest k with AuStart(k) >= off (n when none)
        Private Shared Function FirstAuAtOrAfter(s As PamfMuxStream, off As Long) As Integer
            Dim lo As Integer = 0
            Dim hi As Integer = s.AuStart.Length - 1       ' = n
            While lo < hi
                Dim m As Integer = (lo + hi) \ 2
                If s.AuStart(m) >= off Then hi = m Else lo = m + 1
            End While
            Return lo
        End Function

        ' smallest k whose start code prefix is at or after off (n when none)
        Private Shared Function FirstAnchorAtOrAfter(s As PamfMuxStream, off As Long) As Integer
            Dim lo As Integer = 0
            Dim hi As Integer = s.Anchor.Length
            While lo < hi
                Dim m As Integer = (lo + hi) \ 2
                If s.Anchor(m) >= off Then hi = m Else lo = m + 1
            End While
            Return lo
        End Function

        ' AU holding byte off (0 <= off < total)
        Private Shared Function AuContaining(s As PamfMuxStream, off As Long) As Integer
            Dim lo As Integer = 0
            Dim hi As Integer = s.AuStart.Length - 2       ' = n - 1
            While lo < hi
                Dim m As Integer = (lo + hi + 1) \ 2
                If s.AuStart(m) <= off Then lo = m Else hi = m - 1
            End While
            Return lo
        End Function

        Private Shared Function CeilDiv(a As Long, b As Long) As Long
            If a >= 0 Then Return (a + b - 1L) \ b
            Return -((-a) \ b)
        End Function

        Private Function BuildCandidate(si As Integer) As Candidate
            Dim s As PamfMuxStream = Streams(si)
            Dim pos As Long = _pos(si)
            Dim total As Long = s.TotalBytes
            Dim c As New Candidate() With {.Valid = True}
            Dim p As New PackPlan() With {.StreamIdx = si, .Pos = pos, .TsAu = -1, .BlockIdx = -1}
            Dim cap As Integer
            Dim payload As Long

            If s.IsVideo Then
                While s.Blocks(_blockIdx(si)).EndOffset <= pos
                    _blockIdx(si) += 1
                End While
                Dim bi As Integer = _blockIdx(si)
                Dim blk As VideoBlock = s.Blocks(bi)
                Dim remBlock As Long = blk.EndOffset - pos
                Dim limit As Long = remBlock        ' payload may not run past this many bytes
                Dim n As Integer = s.Aus.Count
                Dim firstInPes As Integer               ' the AU this PES PTS would belong to
                If pos = blk.StartOffset Then
                    Dim au As AccessUnit = s.Aus(blk.FirstAu)
                    p.Rap = True
                    p.BlockIdx = bi
                    p.HasPts = True
                    p.HasDts = (au.Dts <> au.Pts)
                    p.HasPstd = True
                    p.TsAu = blk.FirstAu
                    firstInPes = blk.FirstAu
                    cap = SectorSize - PackHeaderLen - SystemHeaderLen - (6 + 18 + 4 * blk.AuCount) -
                          PesHeaderLength(True, p.HasDts, True, 0)
                Else
                    cap = SectorSize - PackHeaderLen - PesHeaderLength(False, False, False, 0)
                    Dim k As Integer = FirstAnchorAtOrAfter(s, pos)
                    firstInPes = k
                    If k < n AndAlso s.Anchor(k) < pos + Math.Min(CLng(cap), remBlock) AndAlso
                       s.Aus(k).Dts - _lastStampPts(si) >= StampIntervalTicks Then
                        Dim au As AccessUnit = s.Aus(k)
                        Dim hasDts As Boolean = (au.Dts <> au.Pts)
                        Dim capS As Integer = SectorSize - PackHeaderLen - PesHeaderLength(True, hasDts, False, 0)
                        ' with the longer header, AU must still start inside the payload, else no stamp here
                        If s.Anchor(k) < pos + Math.Min(CLng(capS), remBlock) Then
                            p.HasPts = True
                            p.HasDts = hasDts
                            p.TsAu = k
                            cap = capS
                        End If
                    End If
                End If
                ' an AU far enough after the last stamp that would start further inside this pack opens the next PES instead
                Dim lastStamp As Long = If(p.HasPts, s.Aus(p.TsAu).Pts, _lastStampPts(si))
                Dim winEnd As Long = pos + Math.Min(CLng(cap), limit)
                Dim j As Integer = FirstAnchorAtOrAfter(s, pos)
                While j < n AndAlso s.Anchor(j) < winEnd
                    If j <> firstInPes AndAlso j < n - 1 AndAlso s.Aus(j + 1).Dts - lastStamp >= CutIntervalTicks Then
                        limit = Math.Max(1L, s.AuStart(j) - pos)
                        Exit While
                    End If
                    j += 1
                End While
                payload = Math.Min(CLng(cap), limit)
                Dim leftover As Integer = cap - CInt(payload)
                If leftover >= MinPaddingPacket Then p.Padding = leftover Else p.Stuffing = leftover
                c.EndPos = pos + payload
                c.LeadAu = AuContaining(s, c.EndPos - 1)
                ' a short pack (padding or stuffing) that ends on an AU boundary counts against the AU that follows it
                ' (block end, or before an AU that opens next PES) 
                If p.Padding > 0 OrElse p.Stuffing > 0 Then
                    Dim nx As Integer = FirstAuAtOrAfter(s, c.EndPos)
                    If nx < n AndAlso s.AuStart(nx) = c.EndPos Then c.LeadAu = nx
                End If
            Else
                Dim ext As Boolean = _needExt(si)
                Dim hb As Integer = PesHeaderLength(False, False, ext, 0) + AudioSubHeaderLen
                Dim capPts As Integer = SectorSize - PackHeaderLen - hb - 5
                Dim capNo As Integer = capPts + 5
                Dim remaining As Long = total - pos
                Dim k As Integer = FirstAuAtOrAfter(s, pos)
                p.HasPstd = ext
                If k < s.Aus.Count AndAlso s.AuStart(k) < pos + Math.Min(CLng(capPts), remaining) Then
                    p.HasPts = True
                    p.TsAu = k
                    cap = capPts
                    payload = Math.Min(CLng(capPts), remaining)
                    Dim ptr As Integer = CInt(s.AuStart(k) - pos)
                    p.Pointer = If(s.Codec = PamfStreamType.LPCM, &H4000 Or ptr, ptr)
                Else
                    cap = capNo
                    Dim limit As Long = remaining
                    If k < s.Aus.Count Then limit = Math.Min(limit, s.AuStart(k) - pos)
                    payload = Math.Min(CLng(capNo), limit)
                    p.Pointer = If(s.Codec = PamfStreamType.LPCM, &H47FF, &HFFFF)
                End If
                Dim leftover As Integer = cap - CInt(payload)
                If leftover >= MinPaddingPacket Then p.Padding = leftover Else p.Stuffing = leftover
                c.EndPos = pos + payload
                c.LeadAu = AuContaining(s, c.EndPos - 1)
            End If
            p.Len = CInt(payload)
            c.Plan = p

            ' eligibility time (27 MHz) = max(lead rule, room rule)
            Dim leadT As Long = s.T27(c.LeadAu) - StdDelayTicks * 300L
            Dim roomT As Long = Long.MinValue
            Dim need As Long = c.EndPos - CLng(s.PstdBufferSize) * 1024L
            If need > 0 Then
                Dim j As Integer = FirstAuAtOrAfter(s, need)      ' AUs 0..j-1 must have been removed
                If j < 1 Then j = 1
                roomT = s.T27(j - 1)
            End If
            c.Elig27 = Math.Max(leadT, roomT)
            Return c
        End Function

        Public Sub Plan()
            If Streams.Count = 0 Then Throw New InvalidOperationException("no streams to mux")
            Prepare()
            _plan.Clear()
            ScrJumps = 0
            Dim n As Integer = Streams.Count
            ReDim _pos(n - 1)
            ReDim _blockIdx(n - 1)
            ReDim _lastStampPts(n - 1)
            ReDim _needExt(n - 1)
            ReDim _cand(n - 1)
            ReDim _endAuPtr(n - 1)
            For i As Integer = 0 To n - 1
                _needExt(i) = Streams(i).IsAudio
                Streams(i).LateAus = 0
                Streams(i).MaxLate27 = 0
                Streams(i).PeakOccupancy = 0
                Streams(i).MaxLead27 = Long.MinValue
                Streams(i).Stamps = 0
            Next
            Dim slot As Long = SlotTicks27
            Dim origin27 As Long = ScrOriginTicks * 300L

            ' first pack: the first video stream RAP (always start with it ??)
            Dim first As Integer = 0
            For i As Integer = 0 To n - 1
                If Streams(i).IsVideo Then first = i : Exit For
            Next
            _cand(first) = BuildCandidate(first)
            Dim scr0 As Long = origin27 + slot * Math.Max(1L, CeilDiv(_cand(first).Elig27 - origin27, slot))
            Dim scr As Long = scr0
            Emit(first, scr)

            Dim active As Integer = 0
            For i As Integer = 0 To n - 1
                If _pos(i) < Streams(i).TotalBytes Then active += 1
            Next
            While active > 0
                scr += slot
                Dim best As Integer = -1
                Dim minE As Long = Long.MaxValue
                For i As Integer = 0 To n - 1
                    If _pos(i) >= Streams(i).TotalBytes Then Continue For
                    If _cand(i) Is Nothing OrElse Not _cand(i).Valid Then _cand(i) = BuildCandidate(i)
                    If _cand(i).Elig27 < minE Then minE = _cand(i).Elig27
                Next
                If minE > scr Then
                    scr = scr0 + CeilDiv(minE - scr0, slot) * slot
                    ScrJumps += 1
                End If
                Dim bestKey As Long = Long.MaxValue
                For i As Integer = 0 To n - 1
                    If _pos(i) >= Streams(i).TotalBytes Then Continue For
                    If _cand(i).Elig27 > scr Then Continue For
                    Dim s As PamfMuxStream = Streams(i)
                    Dim key As Long = s.T27(AuContaining(s, _pos(i)))
                    If best < 0 OrElse key < bestKey OrElse
                       (key = bestKey AndAlso s.IsVideo AndAlso Not Streams(best).IsVideo) Then
                        best = i
                        bestKey = key
                    End If
                Next
                Emit(best, scr)
                If _pos(best) >= Streams(best).TotalBytes Then active -= 1
            End While

            FinishPlan()
        End Sub

        Private Sub Emit(si As Integer, scr As Long)
            Dim s As PamfMuxStream = Streams(si)
            Dim c As Candidate = If(_cand(si) IsNot Nothing AndAlso _cand(si).Valid, _cand(si), BuildCandidate(si))
            Dim p As PackPlan = c.Plan
            p.Scr27 = scr
            Dim packIndex As Integer = _plan.Count
            _plan.Add(p)

            ' bookkeeping: lead, buffer occupancy, AU end packs
            Dim lead As Long = s.T27(c.LeadAu) - scr
            If lead > s.MaxLead27 Then s.MaxLead27 = lead
            Dim removed As Long = 0
            Dim jr As Integer = FirstRemoved(s, scr)
            If jr > 0 Then removed = s.AuStart(jr)
            Dim occ As Long = c.EndPos - removed
            If occ > s.PeakOccupancy Then s.PeakOccupancy = occ

            Dim k As Integer = _endAuPtr(si)
            While k < s.Aus.Count AndAlso s.AuStart(k + 1) <= c.EndPos
                s.AuEndPack(k) = packIndex
                Dim late As Long = scr + SlotTicks27 - s.T27(k)
                If late > 0 Then
                    s.LateAus += 1
                    If late > s.MaxLate27 Then s.MaxLate27 = late
                End If
                k += 1
            End While
            _endAuPtr(si) = k

            If s.IsVideo Then
                If p.Rap Then
                    s.Blocks(p.BlockIdx).StartPack = packIndex
                    _lastStampPts(si) = s.Aus(p.TsAu).Pts
                    ' every audio stream repeats its P-STD extension after a video private_stream_2
                    For i As Integer = 0 To Streams.Count - 1
                        If Streams(i).IsAudio Then
                            _needExt(i) = True
                            _cand(i) = Nothing
                        End If
                    Next
                ElseIf p.HasPts Then
                    _lastStampPts(si) = s.Aus(p.TsAu).Pts
                    s.Stamps += 1
                End If
            Else
                _needExt(si) = False
            End If
            _pos(si) = c.EndPos
            _cand(si) = Nothing
        End Sub

        ' number of AUs with T27 <= scr (they have left the buffer)
        Private Shared Function FirstRemoved(s As PamfMuxStream, scr As Long) As Integer
            Dim lo As Integer = 0
            Dim hi As Integer = s.T27.Length
            While lo < hi
                Dim m As Integer = (lo + hi) \ 2
                If s.T27(m) <= scr Then lo = m + 1 Else hi = m
            End While
            Return lo
        End Function

        ' ps2 reference-picture offsets and EP entries, now that every AU's end pack is known
        Private Sub FinishPlan()
            For Each s In Streams
                If Not s.IsVideo Then Continue For
                For Each blk In s.Blocks
                    For q As Integer = 0 To 3
                        If q < blk.Refs.Count Then
                            Dim d As Integer = s.AuEndPack(blk.Refs(q)) - blk.StartPack
                            blk.Offsets(q) = Math.Max(0, Math.Min(&HFFFE, d))
                        Else
                            blk.Offsets(q) = &HFFFF
                        End If
                    Next
                    If blk.IsRap AndAlso blk.Refs.Count > 0 Then
                        Dim idx As Integer = blk.Refs.Count
                        s.EpEntries.Add(New EpEntry() With {
                            .Pts = s.Aus(blk.FirstAu).Pts,
                            .PackIndex = blk.StartPack,
                            .IndexN = idx,
                            .RefPictureOffset = Math.Min(blk.Offsets(idx - 1), &H1FFF)
                        })
                    End If
                Next
            Next
        End Sub

        ' writing

        Public Sub WritePacks(output As Stream)
            Dim muxRateUnits As Integer = MuxRateBps \ 8 \ 50
            Dim buf As New MemoryStream(SectorSize)
            Dim auPtr(Streams.Count - 1) As Integer
            For Each p In _plan
                Dim s As PamfMuxStream = Streams(p.StreamIdx)
                buf.SetLength(0)
                WritePackHeader(buf, p.Scr27 \ 300L, CInt(p.Scr27 Mod 300L), muxRateUnits)
                If p.Rap Then
                    WriteSystemHeader(buf, muxRateUnits, s.PstdBufferSize)
                    WritePs2(buf, s, s.Blocks(p.BlockIdx))
                End If
                Dim pts As Long = 0, dts As Long = 0
                If p.HasPts Then
                    pts = s.Aus(p.TsAu).Pts
                    dts = s.Aus(p.TsAu).Dts
                End If
                If s.IsVideo Then
                    WritePesHeader(buf, s.PesStreamId, p.Len, p.HasPts, pts, p.HasDts, dts,
                                   p.HasPstd, s.PstdBufferSize, p.Stuffing)
                Else
                    WritePesHeader(buf, SC_PrivateStream1, AudioSubHeaderLen + p.Len, p.HasPts, pts, False, 0,
                                   p.HasPstd, s.PstdBufferSize, p.Stuffing)
                    ' sub-header: sub_stream_id, config byte (LPCM 0x31 = 48 kHz stereo), first_au pointer
                    WriteAudioSubHeader(buf, s.SubStreamId,
                                        If(s.Codec = PamfStreamType.LPCM, CByte(&H31), CByte(0)),
                                        CUShort(p.Pointer And &HFFFF))
                End If
                CopyPayload(buf, s, p.Pos, p.Len, auPtr(p.StreamIdx))
                If p.Padding > 0 Then WritePaddingStream(buf, p.Padding)
                If buf.Length <> SectorSize Then
                    Throw New InvalidOperationException($"muxer bug: pack of {buf.Length} bytes (stream {p.StreamIdx}, ES offset {p.Pos})")
                End If
                output.Write(buf.GetBuffer(), 0, SectorSize)
            Next

            ' optional MPEG_program_end_code in its own sector
            If WriteEndSector Then
                buf.SetLength(0)
                WriteProgramEnd(buf)
                WritePaddingStream(buf, SectorSize - 4)
                output.Write(buf.GetBuffer(), 0, SectorSize)
            End If
        End Sub

        Private Shared Sub CopyPayload(out As Stream, s As PamfMuxStream, pos As Long, len As Integer,
                                       ByRef auPtr As Integer)
            While auPtr + 1 < s.AuStart.Length - 1 AndAlso s.AuStart(auPtr + 1) <= pos
                auPtr += 1
            End While
            Dim k As Integer = auPtr
            Dim left As Integer = len
            Dim cur As Long = pos
            While left > 0
                Dim au As AccessUnit = s.Aus(k)
                Dim inAu As Integer = CInt(cur - s.AuStart(k))
                Dim take As Integer = Math.Min(left, au.Data.Length - inAu)
                out.Write(au.Data, inAu, take)
                left -= take
                cur += take
                If inAu + take >= au.Data.Length Then k += 1
            End While
        End Sub

        ' private_stream_2 (0xBF), one per block, AVC and MPEG-2 alike:
        '   [0]      0x01
        '   [1]      video stream_id
        '   [2..9]   4 x u16: packs from this pack to the pack holding the last byte of 1st..4th ref pic of the block
        '            0xFFFF when the block has fewer
        '   [10..13] 0
        '   [14..15] 4N + 2
        '   [16..17] N = pictures in the block
        '   [18..]   N x 4 bytes, decode order:
        '            AVC   : 00, ref << 7 | size bits 22..16, size bits 15..8, size bits 7..0
        '            MPEG-2: picture_coding_type << 5, 24 bit size
        Private Shared Sub WritePs2(out As Stream, s As PamfMuxStream, blk As VideoBlock)
            Dim n As Integer = blk.AuCount
            Dim payloadLen As Integer = 18 + 4 * n
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_PrivateStream2)
            out.WriteByte(CByte((payloadLen >> 8) And &HFF))
            out.WriteByte(CByte(payloadLen And &HFF))
            out.WriteByte(&H1)
            out.WriteByte(s.PesStreamId)
            For q As Integer = 0 To 3
                out.WriteByte(CByte((blk.Offsets(q) >> 8) And &HFF))
                out.WriteByte(CByte(blk.Offsets(q) And &HFF))
            Next
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(0)
            Dim f14 As Integer = 4 * n + 2
            out.WriteByte(CByte((f14 >> 8) And &HFF))
            out.WriteByte(CByte(f14 And &HFF))
            out.WriteByte(CByte((n >> 8) And &HFF))
            out.WriteByte(CByte(n And &HFF))
            Dim isM2v As Boolean = (s.Codec = PamfStreamType.MPEG2Video)
            For k As Integer = blk.FirstAu To blk.FirstAu + n - 1
                Dim au As AccessUnit = s.Aus(k)
                Dim size As Integer = au.Data.Length
                If isM2v Then
                    If size > &HFFFFFF Then size = &HFFFFFF
                    out.WriteByte(CByte((au.PictureCodingType And 7) << 5))
                    out.WriteByte(CByte((size >> 16) And &HFF))
                Else
                    If size > &H7FFFFF Then size = &H7FFFFF
                    out.WriteByte(0)
                    out.WriteByte(CByte(If(au.IsReferenceFrame, &H80, 0) Or ((size >> 16) And &H7F)))
                End If
                out.WriteByte(CByte((size >> 8) And &HFF))
                out.WriteByte(CByte(size And &HFF))
            Next
        End Sub

        ' diagnostics

        Public Function DescribePlan() As List(Of String)
            Dim lines As New List(Of String)()
            Dim slot As Long = SlotTicks27
            lines.Add($"packs: {_plan.Count:N0}, SCR0 {FormatScr(If(_plan.Count > 0, _plan(0).Scr27, 0))}, SCR jumps: {ScrJumps:N0}")
            For Each s In Streams
                Dim kind As String = s.Codec.ToString()
                Dim line As String = $"  s{s.Index:D2} {kind}: {s.Aus.Count:N0} AUs, P-STD {s.PstdBufferSize} KB, peak buffer {s.PeakOccupancy:N0} B, max lead {s.MaxLead27 / 27000.0:F1} ms"
                If s.IsVideo Then
                    line &= $", {s.Blocks.Count} ps2 blocks, {s.EpEntries.Count} EP entries, {s.Stamps} mid-block stamps"
                End If
                If s.LateAus > 0 Then line &= $", LATE AUs: {s.LateAus} (max {s.MaxLate27 / 27000.0:F1} ms)"
                lines.Add(line)
            Next
            Return lines
        End Function

        Private Shared Function FormatScr(scr27 As Long) As String
            Return $"{scr27 \ 300L}.{scr27 Mod 300L:D3}"
        End Function

    End Class

End Namespace
