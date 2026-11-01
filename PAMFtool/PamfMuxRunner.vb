' PamfMuxRunner.vb - github.com/ravenDS/PAMFtool

Imports System.IO
Imports PAMFtool.PamfMux

Friend Module PamfMuxRunner

    Private Class MuxInput
        Public Path As String
        Public Index As Integer
        Public Kind As String        ' "avc" / "mpeg2" / "atrac3plus" / "ac3" / "lpcm"
        Public Extension As String   ' lowercased
    End Class

    ' frames per second as an exact fraction
    Private Structure FrameRate
        Public Num As Long
        Public Den As Long
        Public ReadOnly Property Value As Double
            Get
                Return Num / CDbl(Den)
            End Get
        End Property
    End Structure

    ' Sony starts new private_stream_2 block at random access point once current block has lasted 0.4 s
    ' (every block start is a RAP, none is closer than 0.4 s to the previous one unless RAP carries an SPS)
    Private Const MinBlockTicks As Long = 36000L
    ' encoder keeps RAPs 0.4 .. 1.0 s apart, warn about blocks clearly longer than that
    Private Const LongBlockTicks As Long = 99000L

    Private Class VideoTiming
        Public Rate As FrameRate
        Public TotalUnits As Long          ' field units (frame = 2)
        Public EndPts90 As Long
    End Class

    Private _videoTimings As New List(Of VideoTiming)()
    Private _firstVideoIsHighProfile As Boolean? = Nothing

    Public Sub Run(positional As List(Of String),
                   noEp As Boolean,
                   forceDeblock As Boolean,
                   forceNoDeblock As Boolean,
                   noAtsc As Boolean,
                   paceMbps As Double,
                   overridePstdKb As Integer,
                   overrideMmb As Integer,
                   ps2MinPictures As Integer,
                   overrideMuxRateBps As Integer,
                   overrideStdDelayTicks As Integer,
                   overrideInitialScr As Long,
                   endSector As Boolean)
        If positional.Count < 2 Then
            Console.Error.WriteLine("Usage: PAMFtool -mux <inputDir> <output.pamf> [-noep] [-noatsc] [-deblock | -nodeblock] [-pstd <KB>] [-mmb <n>] [-ps2-block <N>] [-muxrate <kbps>] [-std-delay <ticks>] [-initial-scr <ticks>] [-endsector]")
            Environment.Exit(1)
        End If
        Dim inDir As String = positional(0)
        Dim outPath As String = positional(1)
        If Not Directory.Exists(inDir) Then
            Console.Error.WriteLine("Input directory not found: " & inDir)
            Environment.Exit(1)
        End If

        Console.WriteLine("Scanning " & inDir & " for stream files...")
        Dim files As List(Of MuxInput) = ScanInputDirectory(inDir)
        If files.Count = 0 Then
            Console.Error.WriteLine("No recognized stream files. Expected <basename>.sNN_<kind>.<ext>.")
            Environment.Exit(1)
        End If
        For Each f In files
            Console.WriteLine("  s" & f.Index.ToString("D2") & " " & f.Kind &
                              "  ->  " & Path.GetFileName(f.Path))
        Next

        If noEp Then Console.WriteLine("  -noep: EP table will be omitted")
        If forceDeblock Then Console.WriteLine("  -deblock: forcing deblock byte to 1")
        If forceNoDeblock Then Console.WriteLine("  -nodeblock: forcing deblock byte to 0")
        If noAtsc Then Console.WriteLine("  -noatsc: ATRAC3+ ATS extra_config_data forced to zero (ignores atsc chunk)")
        If overridePstdKb > 0 Then Console.WriteLine($"  -pstd {overridePstdKb}: video P-STD buffer size forced to {overridePstdKb} KB")
        If overrideMmb >= 0 Then Console.WriteLine($"  -mmb {overrideMmb}: overriding max_mean_bitrate byte to {overrideMmb}")
        If overrideMuxRateBps > 0 Then Console.WriteLine($"  -muxrate {overrideMuxRateBps \ 1000} kbps: pack_header mux_rate (Sony spec: 48000/24000/12000)")
        If overrideInitialScr >= 0 Then Console.WriteLine($"  -initial-scr {overrideInitialScr}: SCR origin override")
        If endSector Then Console.WriteLine("  -endsector: a program_end_code sector is appended after the last pack")
        If ps2MinPictures > 0 Then
            Console.WriteLine($"  -ps2-block {ps2MinPictures}: ps2/EP blocks start at random access points at least {ps2MinPictures} pictures apart")
        ElseIf ps2MinPictures = 0 Then
            Console.WriteLine("  -ps2-block 0: the legacy 4-byte ps2 marker is no longer written; using Sony's block rule (RAP, >= 0.4 s)")
        End If
        If Not Double.IsNaN(paceMbps) Then
            Console.WriteLine("  -pace: ignored, packs are now scheduled with proper buffer model (1s lead, P-STD room).")
        End If

        Console.WriteLine("Muxing to: " & outPath)
        BuildPamfFromFiles(files, outPath, noEp, forceDeblock, forceNoDeblock, noAtsc,
                           overridePstdKb, overrideMmb, ps2MinPictures, overrideMuxRateBps,
                           overrideStdDelayTicks, overrideInitialScr, endSector)
        Dim sz As Long = New FileInfo(outPath).Length
        Console.WriteLine("Wrote " & outPath & " (" & sz.ToString("N0") & " bytes)")
    End Sub

    Private Function ScanInputDirectory(dir As String) As List(Of MuxInput)
        Dim result As New List(Of MuxInput)()
        Dim wavSidecars As New HashSet(Of String)()

        For Each p_ In Directory.GetFiles(dir, "*.at3")
            wavSidecars.Add(Path.GetFileNameWithoutExtension(p_))
        Next

        For Each p_ In Directory.GetFiles(dir)
            Dim name As String = Path.GetFileNameWithoutExtension(p_)
            Dim dotIdx As Integer = name.LastIndexOf("."c)
            If dotIdx < 0 Then Continue For
            Dim tail As String = name.Substring(dotIdx + 1)
            If tail.Length < 4 OrElse Not tail.StartsWith("s") Then Continue For
            Dim usIdx As Integer = tail.IndexOf("_"c)
            If usIdx < 0 Then Continue For
            Dim idxStr As String = tail.Substring(1, usIdx - 1)
            Dim idx As Integer
            If Not Integer.TryParse(idxStr, idx) Then Continue For

            Dim kind As String = tail.Substring(usIdx + 1).ToLowerInvariant()
            Dim ext As String = Path.GetExtension(p_).ToLowerInvariant()

            ' skip LightCodec WAV sidecars (decoded versions of .at3 inputs).
            If ext = ".wav" AndAlso wavSidecars.Contains(name) Then Continue For

            Dim codecKind As String = Nothing
            Select Case kind
                Case "avc" : codecKind = "avc"
                Case "mpeg-2", "mpeg2", "m2v" : codecKind = "mpeg2"
                Case "atrac3plus" : codecKind = "atrac3plus"
                Case "dolby" : codecKind = "ac3"
                Case "lpcm" : codecKind = "lpcm"
                Case Else
                    ' maybe unrecognized .wav (for example user-supplied LPCM with no .at3)
                    If ext = ".wav" Then codecKind = "lpcm"
            End Select
            If codecKind Is Nothing Then Continue For

            result.Add(New MuxInput() With {
                .Path = p_, .Index = idx, .Kind = codecKind, .Extension = ext
            })
        Next

        ' order by index 
        result.Sort(Function(a, b) a.Index.CompareTo(b.Index))
        Return result
    End Function

    Private Sub BuildPamfFromFiles(files As List(Of MuxInput), outPath As String,
                                   noEp As Boolean, forceDeblock As Boolean,
                                   forceNoDeblock As Boolean, noAtsc As Boolean,
                                   overridePstdKb As Integer,
                                   overrideMmb As Integer,
                                   ps2MinPictures As Integer,
                                   overrideMuxRateBps As Integer,
                                   overrideStdDelayTicks As Integer,
                                   overrideInitialScr As Long,
                                   endSector As Boolean)
        Dim mux As New PamfMuxer()
        mux.SkipEpTable = noEp
        mux.PsMuxer.WriteEndSector = endSector
        _videoTimings.Clear()
        _firstVideoIsHighProfile = Nothing
        If overrideMuxRateBps > 0 Then mux.PsMuxer.MuxRateBps = overrideMuxRateBps
        ' std_delay_bound: 1 s (90000) unless -std-delay says otherwise
        ' official muxer takes it from its mux settings (StdDelayType 1.00 / 0.75 / 0.50 s), not from the stream
        If overrideStdDelayTicks > 0 Then mux.StdDelayBoundTicks = overrideStdDelayTicks

        ' for each input, parse codec metadata, register the stream, then slice into AUs and queue
        For Each f In files
            Select Case f.Kind
                Case "avc" : RegisterAndQueueAvc(mux, f, forceDeblock, forceNoDeblock, overridePstdKb, overrideMmb, ps2MinPictures)
                Case "mpeg2" : RegisterAndQueueM2v(mux, f, overridePstdKb, ps2MinPictures)
                Case "atrac3plus" : RegisterAndQueueAt3p(mux, f, noAtsc)
                Case "ac3" : RegisterAndQueueAc3(mux, f)
                Case "lpcm" : RegisterAndQueueLpcm(mux, f)
            End Select
        Next

        ' SCR origin: AVC High-profile files use 0, AVC Main and MPEG-2 files 30000 (SCR0 = origin + n slots)
        ' -initial-scr keeps accepting the old pack-0 values: 30 -> origin 0, 30030 -> origin 30000
        Dim origin As Long = If(_firstVideoIsHighProfile.GetValueOrDefault(False), 0L, 30000L)
        If overrideInitialScr >= 0 Then
            Select Case overrideInitialScr
                Case 30L : origin = 0L
                Case 30030L : origin = 30000L
                Case Else : origin = overrideInitialScr
            End Select
        End If
        mux.PsMuxer.ScrOriginTicks = origin
        If overridePstdKb > 0 Then mux.AllowPstdGrowth = False

        ' start_pts / end_pts
        Dim times As PresentationTimes = SonyPresentationTimes(mux, _firstVideoIsHighProfile.GetValueOrDefault(False))
        If times.Found Then
            mux.StartPts90 = times.StartPts90
            mux.EndPts90 = times.EndPts90
        Else
            ' no video: start stays 90000, end = last audio
            ' AU + one AU duration
            Dim endPts As Long = 0
            For Each s In mux.PsMuxer.Streams
                If s.Aus.Count = 0 Then Continue For
                Dim last As AccessUnit = s.Aus(s.Aus.Count - 1)
                Dim dur As Long = If(s.Aus.Count > 1, last.Pts - s.Aus(s.Aus.Count - 2).Pts, 0L)
                endPts = Math.Max(endPts, last.Pts + dur)
            Next
            mux.EndPts90 = endPts
        End If

        Using outFs As New FileStream(outPath, FileMode.Create, FileAccess.Write)
            mux.WritePamfFile(outFs)
        End Using
        For Each line In mux.PsMuxer.DescribePlan()
            Console.WriteLine("  " & line)
        Next
    End Sub

    ' header start_pts / end_pts

    Private Structure PresentationTimes
        Public Found As Boolean
        Public StartPts90 As Long
        Public EndPts90 As Long
    End Structure

    ' start_pts / end_pts streams taken in header order:
    '  - each stream gives a [min, max] of its AU PTS values
    '    (video: all AUs, audio: every AU but the first, so audio stream with fewer than 2 AUs gives nothing)
    '  - start = smallest min so far, end = largest max so far
    '  - after each video stream, end = max(end, start + video duration)
    '    start being the value at that point (duration is the frame count x frame time, rounded down to 90 kHz)
    '
    ' audio only counts for newer toolchain
    ' AVC Main-profile files ignore audio here

    Private Function SonyPresentationTimes(mux As PamfMuxer, audioCounts As Boolean) As PresentationTimes
        Dim r As New PresentationTimes()
        Dim videoIndex As Integer = 0
        For Each s In mux.PsMuxer.Streams
            Dim video As Boolean = s.IsVideo
            Dim timing As VideoTiming = Nothing
            If video Then
                If videoIndex < _videoTimings.Count Then timing = _videoTimings(videoIndex)
                videoIndex += 1
            End If
            If Not video AndAlso Not audioCounts Then Continue For
            Dim first As Integer = If(video, 0, 1)
            If s.Aus.Count <= first Then Continue For
            Dim mn As Long = Long.MaxValue, mx As Long = Long.MinValue
            For k As Integer = first To s.Aus.Count - 1
                Dim p As Long = s.Aus(k).Pts
                If p < mn Then mn = p
                If p > mx Then mx = p
            Next
            If Not r.Found Then
                r.Found = True
                r.StartPts90 = mn
                r.EndPts90 = mx
            Else
                If mn < r.StartPts90 Then r.StartPts90 = mn
                If mx > r.EndPts90 Then r.EndPts90 = mx
            End If
            If timing IsNot Nothing Then
                Dim candidate As Long = r.StartPts90 + UnitsToTicks(timing.TotalUnits, timing.Rate, 90000L)
                If candidate > r.EndPts90 Then r.EndPts90 = candidate
            End If
        Next
        Return r
    End Function

    ' timing helpers

    Private Function FloorDiv(a As Decimal, b As Decimal) As Long
        Return CLng(Decimal.Floor(a / b))
    End Function

    ' clock ticks of a position given in field units (frame = 2), exact rational rounded down
    Private Function UnitsToTicks(units As Long, rate As FrameRate, clock As Long) As Long
        Return FloorDiv(CDec(units) * CDec(clock) * CDec(rate.Den), 2D * CDec(rate.Num))
    End Function

    Private Function Gcd(a As Long, b As Long) As Long
        a = Math.Abs(a) : b = Math.Abs(b)
        While b <> 0
            Dim t As Long = a Mod b
            a = b : b = t
        End While
        Return If(a = 0, 1, a)
    End Function

    Private Function MakeRate(num As Long, den As Long) As FrameRate
        Dim g As Long = Gcd(num, den)
        Return New FrameRate() With {.Num = num \ g, .Den = den \ g}
    End Function

    ' DTS/PTS/scheduler time of one video AU
    '   PTS = 90000 + display position, DTS = 90000 + decode position - D (D >= one frame)
    '   both rounded down from the exact fraction (59.94 fps stamps step 1501 / 1502)
    Private Sub SetVideoTimes(au As AccessUnit, displayUnits As Long, decodeUnits As Long, delayUnits As Long,
                              rate As FrameRate)
        au.Pts = 90000L + UnitsToTicks(displayUnits, rate, 90000L)
        au.Dts = 90000L + UnitsToTicks(decodeUnits - delayUnits, rate, 90000L)
        au.DecodeTime27 = au.Dts * 300L
    End Sub

    ' AVC

    Private Sub RegisterAndQueueAvc(mux As PamfMuxer, f As MuxInput,
                                    forceDeblock As Boolean,
                                    forceNoDeblock As Boolean,
                                    overridePstdKb As Integer,
                                    overrideMmb As Integer,
                                    ps2MinPictures As Integer)
        Dim bytes As Byte() = File.ReadAllBytes(f.Path)
        Dim auStarts As List(Of Integer) = FindAvcAuStarts(bytes)
        Dim an As AvcStreamAnalysis = AvcAnalyzer.Analyze(bytes, auStarts)
        Dim sps As H264SpsInfo = an.FirstSps
        If sps Is Nothing Then
            Throw New InvalidDataException(
                "Could not parse SPS from " & f.Path & "; bitstream may not be Annex-B H.264.")
        End If
        Dim pps As H264PpsInfo = an.FirstPps

        Dim rate As FrameRate
        If sps.HasVuiTiming AndAlso sps.NumUnitsInTick > 0UI AndAlso sps.TimeScale > 0UI Then
            rate = MakeRate(CLng(sps.TimeScale), 2L * CLng(sps.NumUnitsInTick))
        Else
            rate = MakeRate(30000, 1001)
            Console.WriteLine("  [warn] AVC SPS has no VUI timing, assuming 29.97 fps")
        End If
        Dim fps As Double = rate.Value
        If AvcFrameRateCodeFromFps(fps) < 0 Then
            Console.WriteLine($"  [warn] AVC frame rate {fps:F3} fps is not a PAMF frame rate (23.976/24/25/29.97/30/50/59.94); header says 29.97")
        End If
        Dim frameRateCode As Byte = CByte(Math.Max(0, AvcFrameRateCodeFromFps(fps)))
        If AvcFrameRateCodeFromFps(fps) < 0 Then frameRateCode = 3

        Dim mbsOnly As Byte = CByte(If(sps.FrameMbsOnlyFlag, 1, 0))
        Dim arIdc As Byte = If(sps.AspectRatioIdc > 0, sps.AspectRatioIdc, CByte(1))
        Dim hasVsi As Byte = CByte(If(sps.HasVideoSignalInfo, 1, 0))
        Dim videoFormat As Byte = If(sps.HasVideoSignalInfo, sps.VideoFormat, CByte(5))
        Dim fullRange As Byte = If(sps.HasVideoSignalInfo, sps.VideoFullRangeFlag, CByte(0))
        ' no VUI colour_description: 2 (unspecified)
        Dim colourPrim As Byte = If(sps.ColourPrimaries > 0, sps.ColourPrimaries, CByte(2))
        Dim transferCh As Byte = If(sps.TransferCharacteristics > 0, sps.TransferCharacteristics, CByte(2))
        Dim matrixCoeff As Byte = If(sps.MatrixCoefficients > 0, sps.MatrixCoefficients, CByte(2))
        Dim cabac As Byte = If(pps IsNot Nothing AndAlso pps.CabacFlag, CByte(1), CByte(0))

        ' codec_info deblockingFilterFlag: 1 = the stream uses in-loop deblocking filter
        ' every slice is checked: 0 only when all slices disable it (disable_deblocking_filter_idc = 1)
        ' a slice that cannot be parsed counts as "deblocking used" so decoder is never set up without it
        Dim deblock As Byte = If(an.SlicesDeblockOn > 0 OrElse an.SliceParseFailures > 0 OrElse an.SliceCount = 0,
                                 CByte(1), CByte(0))
        Console.WriteLine($"  AVC deblocking: {an.SlicesDeblockOn:N0} slices on, {an.SlicesDeblockOff:N0} off" &
                          If(an.SliceParseFailures > 0, $", {an.SliceParseFailures:N0} unparsed (counted as on)", "") &
                          $" -> deblockingFilterFlag {deblock}")
        If forceDeblock Then deblock = 1
        If forceNoDeblock Then deblock = 0

        If _firstVideoIsHighProfile Is Nothing Then
            _firstVideoIsHighProfile = H264SpsParser.IsHighProfileFamily(sps.ProfileIdc)
        End If

        ' max_mean_bitrate byte, computed from the AU sizes the same way Sony's composer does (see SonyMaxMeanBitrate)
        ' -mmb still overrides it
        Dim auSizes(auStarts.Count - 1) As Long
        Dim auIsField(auStarts.Count - 1) As Boolean
        For i As Integer = 0 To auStarts.Count - 1
            Dim e As Integer = If(i + 1 < auStarts.Count, auStarts(i + 1), bytes.Length)
            auSizes(i) = e - auStarts(i)
            auIsField(i) = an.Pictures(i).FieldPic
        Next
        Dim mmb As Byte = SonyMaxMeanBitrate(auSizes, auIsField)
        Dim ps As PamfMuxStream = mux.AddAvcStream(
            profileIdc:=sps.ProfileIdc, levelIdc:=sps.LevelIdc,
            frameMbsOnly:=mbsOnly, frameRateCode:=frameRateCode,
            widthPx:=sps.WidthPixels, heightPx:=sps.HeightPixels,
            aspectRatioIdc:=arIdc,
            sarWidth:=sps.SarWidth, sarHeight:=sps.SarHeight,
            cropLeft:=sps.FrameCropLeftOffset,
            cropRight:=sps.FrameCropRightOffset,
            cropTop:=sps.FrameCropTopOffset,
            cropBottom:=sps.FrameCropBottomOffset,
            hasVideoSignalInfo:=hasVsi,
            videoFormat:=videoFormat,
            fullRangeFlag:=fullRange,
            colourPrimaries:=colourPrim,
            transferChars:=transferCh,
            matrixCoeffs:=matrixCoeff,
            cabacFlag:=cabac,
            deblockFlag:=deblock)

        ' P-STD buffer: ceil((CPB / 8 + 40960) / 1024) KB
        ' original muxer takes CPB from encoder stream info, not from SPS
        ' CPB = max(SPS HRD cpb_size, HRD bit rate x std_delay)
        ' cpb_size = bit rate x 1 s
        ' without HRD, level table is kept, and muxer raises it if AUs would otherwise arrive late
        Dim cpbBits As Long = If(sps.NalHrdCpbSizeBits > 0, sps.NalHrdCpbSizeBits, sps.VclHrdCpbSizeBits)
        If cpbBits > 0 Then
            Dim delayBits As Long = CLng(Decimal.Floor(CDec(sps.HrdPeakBitrateBps) * mux.StdDelayBoundTicks / 90000D))
            Dim bufBits As Long = Math.Max(cpbBits, delayBits)
            ps.PstdBufferSize = CInt(Math.Min(CLng(Mpeg2PsMuxer.MaxPstdKb), (bufBits \ 8L + 40960L + 1023L) \ 1024L))
            If delayBits > cpbBits Then
                Console.WriteLine($"  AVC HRD cpb_size {cpbBits:N0} bits < bit rate x std_delay {delayBits:N0} bits -> P-STD {ps.PstdBufferSize} KB")
            Else
                Console.WriteLine($"  AVC HRD cpb_size {cpbBits:N0} bits -> P-STD {ps.PstdBufferSize} KB")
            End If
        Else
            Console.WriteLine($"  AVC SPS has no HRD: P-STD {ps.PstdBufferSize} KB from the level {sps.LevelIdc / 10.0:F1} table")
        End If
        If overridePstdKb > 0 Then ps.PstdBufferSize = overridePstdKb

        ' MMB: overrideMmb >= 0 means user passed -mmb, use that value
        ' otherwise use the value computed from AU sizes above
        If overrideMmb >= 0 Then
            mux.HeaderWriter.OverrideLastAvcMaxMeanBitrate(CByte(overrideMmb And &HFF))
        Else
            mux.HeaderWriter.OverrideLastAvcMaxMeanBitrate(mmb)
            Console.WriteLine($"  AVC max_mean_bitrate: {mmb} (peak 4-frame mean of {mmb * 100}..{mmb * 100 + 99} kbit per frame)")
        End If

        ' timing and ps2 blocks
        Dim n As Integer = auStarts.Count
        Dim minBlock27 As Long = (MinBlockTicks - 1L) * 300L
        Dim blockStartT27 As Long = 0
        Dim blockStartK As Integer = 0
        Dim blocks As Integer = 0, longBlocks As Integer = 0, nonRapStart As Integer = 0
        For i As Integer = 0 To n - 1
            Dim s As Integer = auStarts(i)
            Dim e As Integer = If(i + 1 < n, auStarts(i + 1), bytes.Length)
            Dim au(e - s - 1) As Byte
            Array.Copy(bytes, s, au, 0, e - s)
            Dim pic As AvcPictureInfo = an.Pictures(i)
            Dim a As New AccessUnit() With {
                .Data = au,
                .IsRandomAccessPoint = pic.IsRap,
                .IsReferenceFrame = pic.IsReference
            }
            SetVideoTimes(a, pic.DisplayUnits, pic.DecodeUnits, an.DelayUnits, rate)
            Dim starts As Boolean
            If i = 0 Then
                starts = True
            ElseIf pic.IsRap Then
                If ps2MinPictures > 0 Then
                    starts = pic.HasSps OrElse (i - blockStartK) >= ps2MinPictures
                Else
                    starts = pic.HasSps OrElse (a.DecodeTime27 - blockStartT27) >= minBlock27
                End If
            Else
                starts = False
            End If
            If starts Then
                If blocks > 0 AndAlso a.DecodeTime27 - blockStartT27 > LongBlockTicks * 300L Then longBlocks += 1
                If Not pic.IsRap Then nonRapStart += 1
                blocks += 1
                blockStartK = i
                blockStartT27 = a.DecodeTime27
            End If
            a.StartsBlock = starts
            mux.QueueAu(ps, a)
        Next
        Dim t As New VideoTiming() With {.Rate = rate, .TotalUnits = an.TotalUnits,
                                         .EndPts90 = 90000L + UnitsToTicks(an.TotalUnits, rate, 90000L)}
        _videoTimings.Add(t)
        Console.WriteLine($"  AVC: {n:N0} AUs, {fps:F3} fps, DTS offset {an.DelayUnits / 2.0:0.#} frame(s), {blocks:N0} ps2 blocks")
        If nonRapStart > 0 Then Console.WriteLine("  [warn] the stream does not start with a random access point (IDR or recovery-point I picture)")
        If longBlocks > 0 Then Console.WriteLine($"  [warn] {longBlocks} ps2 blocks are longer than 1.1 s (Sony's RAP interval is 0.4 .. 1.0 s); seeking / EP granularity follows the encoder's RAP spacing")
    End Sub

    ' max_mean_bitrate byte (AVC codec_info byte 25),
    ' derives it from the multiplexer's per-AU size list:
    '  - two field AUs make one frame (their sizes are added), a frame AU is one frame
    '  - for every run of 4 consecutive frames: mean frame size in bits, in units of 100000 bits, rounded down
    '  - only frames - 4 runs are looked at (the run ending on the last frame is skipped)
    '  - the result is the largest run value, kept in one byte (values above 255 wrap)
    '  - fewer than 5 frames: 0
    ' 32 bit unsigned like the original
    Friend Function SonyMaxMeanBitrate(auSizes As IList(Of Long), auIsField As IList(Of Boolean)) As Byte
        Dim frames As New List(Of Long)(auSizes.Count)
        Dim pendingField As Long = 0
        For i As Integer = 0 To auSizes.Count - 1
            If auIsField(i) Then
                If pendingField = 0 Then
                    pendingField = auSizes(i)
                Else
                    frames.Add(auSizes(i) + pendingField)
                    pendingField = 0
                End If
            Else
                frames.Add(auSizes(i))
            End If
        Next
        Dim best As Integer = 0
        For i As Integer = 0 To frames.Count - 5
            Dim bits As Long = 0
            For k As Integer = 0 To 3
                bits = (frames(i + k) * 8L + bits) And &HFFFFFFFFL
            Next
            Dim v As Integer = CInt(((bits >> 2) \ 100000L) And &HFFL)
            If v > best Then best = v
        Next
        Return CByte(best)
    End Function

    Private Function FindAvcAuStarts(bytes As Byte()) As List(Of Integer)
        ' H.264 AU boundary detection
        '
        ' reliable boundary signals:
        '  - NAL type 9 (Access Unit Delimiter)
        '  - first VCL NAL of a picture has first_mb_in_slice == 0

        Dim nalStartCode As New List(Of Integer)()
        Dim nalType As New List(Of Integer)()
        Dim nalIsFirstSlice As New List(Of Boolean)()
        Dim i As Integer = 0
        While i < bytes.Length - 4
            Dim nalHdr As Integer = FindAnnexBNal(bytes, i)
            If nalHdr < 0 Then Exit While
            Dim startCodeAt As Integer = NalStartCodePos(bytes, nalHdr)
            Dim nt As Integer = bytes(nalHdr) And &H1F
            Dim isVcl As Boolean = (nt = 1 OrElse nt = 5)
            Dim isFirstSlice As Boolean = isVcl AndAlso
                (nalHdr + 1 < bytes.Length) AndAlso ((bytes(nalHdr + 1) And &H80) <> 0)
            nalStartCode.Add(startCodeAt)
            nalType.Add(nt)
            nalIsFirstSlice.Add(isFirstSlice)
            i = nalHdr + 1
        End While

        Dim result As New List(Of Integer)()
        Dim n As Integer = nalStartCode.Count
        If n = 0 Then
            result.Add(0)
            Return result
        End If

        ' first AU starts at the first NAL
        result.Add(nalStartCode(0))
        Dim haveVclInCurrentAu As Boolean = (nalType(0) = 1 OrElse nalType(0) = 5)
        Dim pendingNonVclStart As Integer = -1

        For idx As Integer = 1 To n - 1
            Dim nt As Integer = nalType(idx)
            Dim isVcl As Boolean = (nt = 1 OrElse nt = 5)
            Dim isAud As Boolean = (nt = 9)
            Dim startsNewPicture As Boolean =
                (isAud AndAlso haveVclInCurrentAu) OrElse
                (isVcl AndAlso nalIsFirstSlice(idx) AndAlso haveVclInCurrentAu)

            If startsNewPicture Then
                Dim auStart As Integer = If(pendingNonVclStart >= 0,
                                            pendingNonVclStart, nalStartCode(idx))
                result.Add(auStart)
                haveVclInCurrentAu = isVcl
                pendingNonVclStart = -1
            ElseIf isVcl Then
                haveVclInCurrentAu = True
                pendingNonVclStart = -1
            ElseIf haveVclInCurrentAu Then
                ' non-VCL after a VCL, provisional next-AU start
                If pendingNonVclStart < 0 Then pendingNonVclStart = nalStartCode(idx)
            End If
        Next

        Return result
    End Function

    Private Function FindAnnexBNal(bytes As Byte(), startAt As Integer) As Integer
        Dim i As Integer = startAt
        While i < bytes.Length - 4
            If bytes(i) = 0 AndAlso bytes(i + 1) = 0 Then
                If bytes(i + 2) = 1 Then Return i + 3
                If bytes(i + 2) = 0 AndAlso bytes(i + 3) = 1 Then Return i + 4
            End If
            i += 1
        End While
        Return -1
    End Function

    Private Function NalStartCodePos(bytes As Byte(), nalHdrPos As Integer) As Integer
        ' walk backwards from nalHdrPos to find the 00 00 01 or 00 00 00 01
        If nalHdrPos >= 4 AndAlso bytes(nalHdrPos - 4) = 0 AndAlso bytes(nalHdrPos - 3) = 0 _
        AndAlso bytes(nalHdrPos - 2) = 0 AndAlso bytes(nalHdrPos - 1) = 1 Then
            Return nalHdrPos - 4
        End If
        Return nalHdrPos - 3
    End Function

    ' PAMF AVC frame_rate code (header stores code + 1), -1 when fps is not a PAMF rate
    Private Function AvcFrameRateCodeFromFps(fps As Double) As Integer
        If Math.Abs(fps - 24000.0 / 1001.0) < 0.01 Then Return 0
        If Math.Abs(fps - 24.0) < 0.01 Then Return 1
        If Math.Abs(fps - 25.0) < 0.01 Then Return 2
        If Math.Abs(fps - 30000.0 / 1001.0) < 0.01 Then Return 3
        If Math.Abs(fps - 30.0) < 0.01 Then Return 4
        If Math.Abs(fps - 50.0) < 0.01 Then Return 5
        If Math.Abs(fps - 60000.0 / 1001.0) < 0.01 Then Return 6
        Return -1
    End Function

    ' MPEG-2 frame_rate_code, -1 when fps is not a PAMF rate
    Private Function M2vFrameRateCodeFromFps(fps As Double) As Integer
        If Math.Abs(fps - 24000.0 / 1001.0) < 0.01 Then Return 1
        If Math.Abs(fps - 24.0) < 0.01 Then Return 2
        If Math.Abs(fps - 25.0) < 0.01 Then Return 3
        If Math.Abs(fps - 30000.0 / 1001.0) < 0.01 Then Return 4
        If Math.Abs(fps - 30.0) < 0.01 Then Return 5
        If Math.Abs(fps - 50.0) < 0.01 Then Return 6
        If Math.Abs(fps - 60000.0 / 1001.0) < 0.01 Then Return 7
        Return -1
    End Function

    ' MPEG-2 video
    ' streamed in two passes so >2 GB sources work:
    '   pass 1: scan start codes -> AU boundaries, picture types, picture coding extension, sequence headers
    '   pass 2: read each AU into its own Byte()
    ' AU boundaries: a sequence header / GOP header belongs to the picture that follows it
    '                everything before the first picture belongs to the first AU

    Private Class M2vPicture
        Public AuStart As Long
        Public CodingType As Integer = 1
        Public PictureStructure As Integer = 3       ' 1 top field, 2 bottom field, 3 frame
        Public TopFieldFirst As Boolean
        Public RepeatFirstField As Boolean
        Public HasSequenceHeader As Boolean
        Public Units As Integer
        Public DecodeUnits As Long
        Public DisplayUnits As Long
    End Class

    Private Sub RegisterAndQueueM2v(mux As PamfMuxer, f As MuxInput, overridePstdKb As Integer, ps2MinPictures As Integer)
        ' 1) peek at file head to parse the first sequence_header
        Dim fi As New FileInfo(f.Path)
        Dim fileLen As Long = fi.Length
        Dim peekLen As Integer = CInt(Math.Min(CLng(1024 * 1024), fileLen))
        Dim peek(peekLen - 1) As Byte
        Using pfs As FileStream = File.OpenRead(f.Path)
            Dim total As Integer = 0
            While total < peekLen
                Dim n As Integer = pfs.Read(peek, total, peekLen - total)
                If n <= 0 Then Exit While
                total += n
            End While
        End Using
        Dim seq As M2vSequenceInfo = MpegSequenceHeaderParser.ParseFirstSequenceHeader(
            peek, 0, peek.Length)
        If seq Is Nothing Then
            Throw New InvalidDataException("Could not parse sequence_header from " & f.Path)
        End If
        Dim rate As FrameRate = M2vFrameRate(seq)
        If rate.Num = 0 Then
            rate = MakeRate(30000, 1001)
            Console.WriteLine("  [warn] MPEG-2 frame_rate_code is invalid, assuming 29.97 fps")
        End If
        Dim fps As Double = rate.Value

        If _firstVideoIsHighProfile Is Nothing Then _firstVideoIsHighProfile = False

        Dim frameRateCode As Integer = M2vFrameRateCodeFromFps(fps)
        If frameRateCode < 0 Then
            Console.WriteLine($"  [warn] MPEG-2 frame rate {fps:F3} fps is not a PAMF frame rate; header says 29.97")
            frameRateCode = 4
        End If
        Dim prog As Byte = CByte(If(seq.ProgressiveSequence, 1, 0))
        ' no colour_description (or no sequence_display_extension): 2/2/2 (unspecified)
        Dim colourPrim As Byte = If(seq.HasColourDescription, seq.ColourPrimaries, CByte(2))
        Dim transferCh As Byte = If(seq.HasColourDescription, seq.TransferCharacteristics, CByte(2))
        Dim matrixCoeff As Byte = If(seq.HasColourDescription, seq.MatrixCoefficients, CByte(2))
        Dim ps As PamfMuxStream = mux.AddM2vStream(
            profileLevel:=seq.ProfileAndLevel,
            progressive:=prog,
            frameRateCode:=CByte(frameRateCode),
            widthPx:=seq.WidthPixels, heightPx:=seq.HeightPixels,
            colourPrimaries:=colourPrim,
            transferChars:=transferCh,
            matrixCoeffs:=matrixCoeff)

        ' 2) pass 1: scan the whole file for start codes
        Dim pics As New List(Of M2vPicture)()
        Dim pendingStart As Long = -1           ' first sequence / GOP header since the last picture
        Dim seqPending As Boolean = False
        Dim lastVbv As Integer = -1
        Dim maxVbv As Integer = 0
        Dim progressiveSeq As Boolean = seq.ProgressiveSequence
        Const ScanBufSize As Integer = 1024 * 1024
        Const Overlap As Integer = 16

        Using fs As FileStream = File.OpenRead(f.Path)
            Dim buf(ScanBufSize + Overlap - 1) As Byte
            Dim carry As Integer = 0
            Dim baseOff As Long = 0     ' file offset that buf[0] represents
            Dim atEof As Boolean = False

            Do
                Dim toRead As Integer = buf.Length - carry
                Dim n As Integer = fs.Read(buf, carry, toRead)
                If n = 0 Then atEof = True
                Dim totalInBuf As Integer = carry + n

                ' leave `Overlap` bytes at the tail for the next round unless this is the last chunk
                Dim scanEnd As Integer = If(atEof, totalInBuf - 3, totalInBuf - Overlap)
                Dim i As Integer = 0
                While i < scanEnd
                    If buf(i) = 0 AndAlso buf(i + 1) = 0 AndAlso buf(i + 2) = 1 Then
                        Dim sc As Byte = buf(i + 3)
                        Dim off As Long = baseOff + i
                        Dim have As Integer = totalInBuf - i      ' bytes available from the start code on
                        Select Case sc
                            Case &H0
                                Dim pic As New M2vPicture()
                                If pics.Count = 0 Then
                                    pic.AuStart = 0
                                    pic.HasSequenceHeader = seqPending
                                Else
                                    pic.AuStart = If(pendingStart >= 0, pendingStart, off)
                                    pic.HasSequenceHeader = seqPending
                                End If
                                If have >= 6 Then pic.CodingType = (buf(i + 5) >> 3) And 7
                                pics.Add(pic)
                                pendingStart = -1
                                seqPending = False
                            Case &HB3
                                If pendingStart < 0 Then pendingStart = off
                                seqPending = True
                                If have >= 12 Then
                                    Dim w As Long = (CLng(buf(i + 8)) << 24) Or (CLng(buf(i + 9)) << 16) Or
                                                    (CLng(buf(i + 10)) << 8) Or buf(i + 11)
                                    lastVbv = CInt((w >> 3) And &H3FFL)
                                    maxVbv = Math.Max(maxVbv, lastVbv)
                                End If
                            Case &HB8
                                If pendingStart < 0 Then pendingStart = off
                            Case &HB5
                                If have >= 9 Then
                                    Dim extId As Integer = buf(i + 4) >> 4
                                    If extId = 1 Then
                                        progressiveSeq = ((buf(i + 5) >> 3) And 1) <> 0
                                        If lastVbv >= 0 Then
                                            maxVbv = Math.Max(maxVbv, lastVbv Or (CInt(buf(i + 8)) << 10))
                                        End If
                                    ElseIf extId = 8 AndAlso pics.Count > 0 Then
                                        Dim pic As M2vPicture = pics(pics.Count - 1)
                                        pic.PictureStructure = buf(i + 6) And 3
                                        pic.TopFieldFirst = (buf(i + 7) And &H80) <> 0
                                        pic.RepeatFirstField = (buf(i + 7) And &H2) <> 0
                                    End If
                                End If
                        End Select
                        i += 4
                        Continue While
                    End If
                    i += 1
                End While

                ' carry the last (totalInBuf - i) bytes to the front for the next read
                If atEof Then
                    Exit Do
                End If
                Dim remaining As Integer = totalInBuf - i
                If remaining > 0 Then
                    Array.Copy(buf, i, buf, 0, remaining)
                End If
                baseOff += i
                carry = remaining
            Loop
        End Using

        If pics.Count = 0 Then
            Throw New InvalidDataException("No picture_start_code found in " & f.Path)
        End If

        ' P-STD: ceil((vbv_buffer_size * 2048 + 40960) / 1024) KB (vbv_buffer_size in 16 kbit units)
        If maxVbv > 0 Then
            ps.PstdBufferSize = CInt(Math.Min(CLng(Mpeg2PsMuxer.MaxPstdKb), (CLng(maxVbv) * 2048L + 40960L + 1023L) \ 1024L))
            Console.WriteLine($"  MPEG-2 vbv_buffer_size {maxVbv} -> P-STD {ps.PstdBufferSize} KB")
        End If
        If overridePstdKb > 0 Then ps.PstdBufferSize = overridePstdKb

        ' display order: B pictures are shown as they are decoded,
        ' an I/P picture (frame or field pair) when the next I/P arrives
        ' positions in field units (frame = 2, + repeat_first_field)
        Dim nPic As Integer = pics.Count
        Dim dec As Long = 0
        For k As Integer = 0 To nPic - 1
            Dim p As M2vPicture = pics(k)
            If p.PictureStructure <> 3 AndAlso p.PictureStructure <> 0 Then
                p.Units = 1
            ElseIf progressiveSeq Then
                p.Units = 2 * (1 + If(p.RepeatFirstField, 1, 0) + If(p.RepeatFirstField AndAlso p.TopFieldFirst, 1, 0))
            Else
                p.Units = 2 + If(p.RepeatFirstField, 1, 0)
            End If
            p.DecodeUnits = dec
            dec += p.Units
        Next
        Dim totalUnits As Long = dec
        Dim dispOrder As New List(Of Integer)()
        Dim pending As New List(Of Integer)()
        Dim k2 As Integer = 0
        While k2 < nPic
            Dim group As New List(Of Integer) From {k2}
            If pics(k2).PictureStructure <> 3 AndAlso pics(k2).PictureStructure <> 0 AndAlso k2 + 1 < nPic AndAlso
               pics(k2 + 1).PictureStructure <> 3 AndAlso pics(k2 + 1).PictureStructure <> 0 AndAlso
               pics(k2 + 1).PictureStructure <> pics(k2).PictureStructure Then
                group.Add(k2 + 1)
            End If
            If pics(k2).CodingType = 3 Then
                dispOrder.AddRange(group)
            Else
                dispOrder.AddRange(pending)
                pending = group
            End If
            k2 += group.Count
        End While
        dispOrder.AddRange(pending)
        Dim disp As Long = 0
        For Each k In dispOrder
            pics(k).DisplayUnits = disp
            disp += pics(k).Units
        Next
        Dim delay As Long = 2
        For Each p In pics
            delay = Math.Max(delay, p.DecodeUnits - p.DisplayUnits)
        Next

        ' 3) pass 2: read each AU and queue it
        Dim blocks As Integer = 0, longBlocks As Integer = 0
        Dim blockStartK As Integer = 0
        Dim blockStartT27 As Long = 0
        Using fs As FileStream = File.OpenRead(f.Path)
            For i As Integer = 0 To nPic - 1
                Dim startOff As Long = pics(i).AuStart
                Dim endOff As Long = If(i + 1 < nPic, pics(i + 1).AuStart, fileLen)
                Dim sz As Long = endOff - startOff
                If sz <= 0 Then Throw New InvalidDataException($"MPEG-2 picture {i} has no data")
                If sz > Int32.MaxValue Then
                    Throw New InvalidDataException(
                        "Picture " & i & " byte range exceeds 2 GiB (" & sz & " bytes); refusing to allocate.")
                End If
                Dim au(CInt(sz) - 1) As Byte
                fs.Position = startOff
                Dim total As Integer = 0
                While total < au.Length
                    Dim n As Integer = fs.Read(au, total, au.Length - total)
                    If n <= 0 Then Exit While
                    total += n
                End While

                Dim p As M2vPicture = pics(i)
                ' random access point = I picture whose AU carries a sequence header (Sony: one ps2 block per GOP)
                Dim isRap As Boolean = (p.CodingType = 1) AndAlso p.HasSequenceHeader
                Dim a As New AccessUnit() With {
                    .Data = au,
                    .IsRandomAccessPoint = isRap,
                    .IsReferenceFrame = (p.CodingType = 1 OrElse p.CodingType = 2),
                    .PictureCodingType = CByte(p.CodingType And 7)
                }
                SetVideoTimes(a, p.DisplayUnits, p.DecodeUnits, delay, rate)
                Dim starts As Boolean = (i = 0) OrElse (isRap AndAlso (ps2MinPictures <= 0 OrElse i - blockStartK >= ps2MinPictures))
                If starts Then
                    If blocks > 0 AndAlso a.DecodeTime27 - blockStartT27 > LongBlockTicks * 300L Then longBlocks += 1
                    blocks += 1
                    blockStartK = i
                    blockStartT27 = a.DecodeTime27
                End If
                a.StartsBlock = starts
                mux.QueueAu(ps, a)
            Next
        End Using
        _videoTimings.Add(New VideoTiming() With {.Rate = rate, .TotalUnits = totalUnits,
                                                  .EndPts90 = 90000L + UnitsToTicks(totalUnits, rate, 90000L)})
        Console.WriteLine($"  MPEG-2: {nPic:N0} pictures, {fps:F3} fps, DTS offset {delay / 2.0:0.#} frame(s), {blocks:N0} ps2 blocks")
        If Not ((pics(0).CodingType = 1) AndAlso pics(0).HasSequenceHeader) Then
            Console.WriteLine("  [warn] the stream does not start with an I picture preceded by a sequence header")
        End If
        If longBlocks > 0 Then Console.WriteLine($"  [warn] {longBlocks} ps2 blocks (GOPs) are longer than 1.1 s")
    End Sub

    ' frame rate = frame_rate_code base * (frame_rate_extension_n + 1) / (frame_rate_extension_d + 1)
    Private Function M2vFrameRate(seq As M2vSequenceInfo) As FrameRate
        Dim num As Long, den As Long
        Select Case seq.FrameRateCode
            Case 1 : num = 24000 : den = 1001
            Case 2 : num = 24 : den = 1
            Case 3 : num = 25 : den = 1
            Case 4 : num = 30000 : den = 1001
            Case 5 : num = 30 : den = 1
            Case 6 : num = 50 : den = 1
            Case 7 : num = 60000 : den = 1001
            Case 8 : num = 60 : den = 1
            Case Else : Return New FrameRate() With {.Num = 0, .Den = 1}
        End Select
        Return MakeRate(num * (seq.FrameRateExtN + 1L), den * (seq.FrameRateExtD + 1L))
    End Function

    ' ATRAC3plus: read .at3 RIFF, rebuild 8-byte ATS header per frame
    '
    ' Each PAMF ATRAC3plus AU on the wire is:
    ' [ 8-byte ATS header ] [ raw_data_frame ]
    '
    ' our extractor strips the ATS header and stores "extra_config_data" in custom "atsc" RIFF chunk so we can round-trip it here
    '
    ' fallbacks (in order):
    '  * -noatsc CLI flag on the mux            : always emit zeros for bytes 4-7
    '  * .at3 has atsc chunk of size N*4        : use per-frame (preferred)
    '  * .at3 has legacy atsc chunk of size 4   : broadcast that single value to every frame
    '  * .at3 has no atsc chunk at all          : zero bytes 4-7
    '
    ' sync/params bytes (0..3) are always synthesized from channels + sample_rate + block_align

    Private Sub RegisterAndQueueAt3p(mux As PamfMuxer, f As MuxInput, noAtsc As Boolean)
        Dim chs As Integer = 0
        Dim sr As Integer = 0
        Dim frameSize As Integer = 0
        Dim perFrameExtra As List(Of Byte()) = Nothing
        Dim frames As List(Of Byte()) = ReadAt3Frames(f.Path, chs, sr, frameSize, perFrameExtra)
        Dim ps As PamfMuxStream = mux.AddAtrac3plusStream(numChannels:=CByte(chs))

        ' ATS header params field layout:
        '  [15:13]  sample_rate_idx  : 1 = 44100, 2 = 48000
        '  [12:10]  ch_config_idx    : 1 = 1ch, 2 = 2ch, 5 = 6ch (5.1), 7 = 8ch (7.1)
        '                              (3, 4, 6 = 3/4/7 channels; docs list 1/2/6/8 as the "official" set)
        '  [9:0]    nbytes_encoded   : (raw_data_frame_size / 8) - 1
        Dim srIdx As Integer
        Select Case sr
            Case 44100 : srIdx = 1
            Case 48000 : srIdx = 2
            Case Else
                Throw New InvalidDataException(
                    $"ATRAC3plus sample rate {sr} Hz is unsupported by PAMF (44100 / 48000 only).")
        End Select

        Dim chIdx As Integer
        Select Case chs
            Case 1 : chIdx = 1
            Case 2 : chIdx = 2
            Case 3 : chIdx = 3
            Case 4 : chIdx = 4
            Case 6 : chIdx = 5
            Case 7 : chIdx = 6
            Case 8 : chIdx = 7
            Case Else
                Throw New InvalidDataException(
                    $"ATRAC3plus channel count {chs} isn't representable in an ATS header.")
        End Select

        If (frameSize Mod 8) <> 0 OrElse frameSize <= 0 OrElse frameSize > &H200 * 8 Then
            Throw New InvalidDataException(
                $"ATRAC3plus block_align {frameSize} out of range - must be a positive multiple of 8, at most 4096.")
        End If
        Dim nbytesEnc As Integer = (frameSize \ 8) - 1
        Dim params As Integer = (srIdx << 13) Or (chIdx << 10) Or (nbytesEnc And &H3FF)

        ' one AU per frame, PTS spaced by samples-per-frame (2048) at sample rate converted to 90 kHz
        ' (3840 ticks at 48 kHz; exact fraction rounded down per frame otherwise)
        ' AT3+ has a decoder priming delay of 2416 samples (encoder look-ahead + filterbank/QMF warm-up)
        ' first N samples the decoder emits are silence / pre-roll
        '
        ' audio waveform actual t=0 lands at DTS = ptsBase - priming
        ' for 48 kHz AT3+, first AU PTS = video_first_pts - 2416 * 90000 / 48000 = video_first_pts - 4530
        '
        Const At3PlusPrimingSamples As Long = 2416L
        Dim primingTicks As Long = At3PlusPrimingSamples * 90000L \ CLng(sr)
        Dim ptsBase As Long = 90000L - primingTicks
        Dim zeros(3) As Byte
        For i As Integer = 0 To frames.Count - 1
            Dim raw As Byte() = frames(i)
            Dim ec As Byte()
            If noAtsc OrElse perFrameExtra Is Nothing OrElse i >= perFrameExtra.Count Then
                ec = zeros
            Else
                ec = perFrameExtra(i)
            End If
            ' build the PAMF AU: 8-byte ATS header + raw_data_frame
            Dim pamfAu(8 + raw.Length - 1) As Byte
            pamfAu(0) = &HF                            ' sync high byte
            pamfAu(1) = &HD0                           ' sync low byte
            pamfAu(2) = CByte((params >> 8) And &HFF)  ' params high byte
            pamfAu(3) = CByte(params And &HFF)         ' params low byte
            pamfAu(4) = ec(0)
            pamfAu(5) = ec(1)
            pamfAu(6) = ec(2)
            pamfAu(7) = ec(3)
            Array.Copy(raw, 0, pamfAu, 8, raw.Length)
            Dim pts As Long = ptsBase + (CLng(i) * 2048L * 90000L) \ CLng(sr)
            mux.QueueAu(ps, New AccessUnit() With {
                .Data = pamfAu, .Pts = pts, .Dts = pts, .DecodeTime27 = pts * 300L,
                .IsRandomAccessPoint = False
            })
        Next
    End Sub

    Private Function ReadAt3Frames(path As String, ByRef channels As Integer,
                                   ByRef sampleRate As Integer,
                                   ByRef frameSize As Integer,
                                   ByRef perFrameExtra As List(Of Byte())) As List(Of Byte())
        ' parse RIFF/WAVE-style .at3 file:
        '  WAVE_FORMAT_EXTENSIBLE (0xFFFE) with AT3+ SubFormat GUID
        '  nBlockAlign = raw_data_frame size (ATS header stripped)
        '  optional "atsc" chunk of size N*4 OR 4 (see fallbacks below)
        '
        ' 'perFrameExtra' is set to a list of 4-byte ATS extra_config_data entries
        Dim data As Byte() = File.ReadAllBytes(path)
        If data.Length < 12 Then Throw New InvalidDataException("Tiny .at3 file: " & path)
        If data(0) <> &H52 OrElse data(1) <> &H49 _
        OrElse data(2) <> &H46 OrElse data(3) <> &H46 Then
            Throw New InvalidDataException("Not a RIFF file: " & path)
        End If
        Dim pos As Integer = 12
        Dim dataOff As Integer = -1
        Dim dataLen As Integer = 0
        Dim atscOff As Integer = -1
        Dim atscSize As Integer = 0
        ' walk ALL chunks (not just those before data), because our extractor emits atsc AFTER data chunk now
        While pos < data.Length - 8
            Dim chunkId As String = System.Text.Encoding.ASCII.GetString(data, pos, 4)
            Dim chunkSize As Integer = BitConverter.ToInt32(data, pos + 4)
            If chunkId = "fmt " Then
                ' offsets within fmt chunk: +2 channels, +4 sampleRate, +12 blockAlign
                channels = BitConverter.ToUInt16(data, pos + 8 + 2)
                sampleRate = BitConverter.ToInt32(data, pos + 8 + 4)
                frameSize = BitConverter.ToUInt16(data, pos + 8 + 12)
            ElseIf chunkId = "atsc" Then
                atscOff = pos + 8
                atscSize = chunkSize
            ElseIf chunkId = "data" Then
                dataOff = pos + 8
                dataLen = chunkSize
            End If
            pos += 8 + chunkSize
            If (chunkSize And 1) = 1 Then pos += 1
        End While
        If dataOff < 0 Then Throw New InvalidDataException("No data chunk in .at3: " & path)
        If frameSize <= 0 Then frameSize = 688

        Dim frames As New List(Of Byte())()
        Dim n As Integer = dataLen \ frameSize
        For i As Integer = 0 To n - 1
            Dim frame(frameSize - 1) As Byte
            Array.Copy(data, dataOff + i * frameSize, frame, 0, frameSize)
            frames.Add(frame)
        Next

        ' resolve extra_config_data source:
        '  - atsc chunk size == n * 4  -> per-frame list (preferred)
        '  - atsc chunk size == 4      -> broadcast to every frame (legacy)
        '  - anything else / no chunk  -> nothing (caller zero-fills)
        perFrameExtra = Nothing
        If atscOff >= 0 Then
            If atscSize = n * 4 AndAlso n > 0 Then
                perFrameExtra = New List(Of Byte())()
                For i As Integer = 0 To n - 1
                    Dim ec(3) As Byte
                    Array.Copy(data, atscOff + i * 4, ec, 0, 4)
                    perFrameExtra.Add(ec)
                Next
            ElseIf atscSize = 4 AndAlso n > 0 Then
                ' legacy per-stream atsc. Broadcast.
                perFrameExtra = New List(Of Byte())()
                Dim broadcast(3) As Byte
                Array.Copy(data, atscOff, broadcast, 0, 4)
                For i As Integer = 0 To n - 1
                    perFrameExtra.Add(broadcast)
                Next
            End If
        End If

        Return frames
    End Function

    ' AC-3: walk syncframes by their coded size (A/52 table 5.18), one AU per syncframe
    ' (searching for every 0x0B77 also splits frames where payload happens to contain these two bytes)

    Private Sub RegisterAndQueueAc3(mux As PamfMuxer, f As MuxInput)
        Dim data As Byte() = File.ReadAllBytes(f.Path)
        Dim firstSync As Integer = FindAc3Sync(data, 0)
        If firstSync < 0 Then
            Throw New InvalidDataException("No AC-3 sync word found in " & f.Path)
        End If
        ' syncinfo: syncword (16), crc1 (16), fscod (2) + frmsizecod (6) at byte+4
        ' bsi: bsid (5) + bsmod (3) at byte+5, acmod (3) at the top of byte+6, then optional mix levels and lfeon
        If (data(firstSync + 4) >> 6) <> 0 Then
            Console.WriteLine("  [warn] AC-3 stream is not 48 kHz; PAMF audio is 48 kHz and PTS assume 1536-sample frames at 48 kHz")
        End If
        Dim acmod As Integer = (data(firstSync + 6) >> 5) And 7
        Dim chTable As Integer() = {2, 1, 2, 3, 3, 4, 4, 5}    ' acmod 0..7 full-bandwidth channels
        Dim channels As Integer = chTable(acmod) + Ac3LfeOn(data, firstSync)

        Dim ps As PamfMuxStream = mux.AddAc3Stream(numChannels:=CByte(channels))

        ' 1536 samples per syncframe = 2880 ticks at 48 kHz, first PTS 90000 (??)
        Dim ptsStep As Long = 1536L * 90000L \ 48000L
        Dim ptsBase As Long = 90000L

        ' walk frame by frame, resync on the next valid frame if a header is broken
        Dim auStarts As New List(Of Integer)()
        Dim p As Integer = firstSync
        Dim resyncs As Integer = 0
        While p >= 0 AndAlso p + 5 <= data.Length
            Dim size As Integer = 0
            If data(p) = &HB AndAlso data(p + 1) = &H77 Then size = Ac3FrameBytes(data(p + 4))
            If size > 0 Then
                auStarts.Add(p)
                p += size                       ' a truncated last frame keeps the remaining bytes
            Else
                p = FindAc3Sync(data, p + 1)
                resyncs += 1
            End If
        End While
        If resyncs > 0 Then
            Console.WriteLine($"  [warn] AC-3: {resyncs} broken frame header(s), resynced on the next valid frame")
        End If
        For i As Integer = 0 To auStarts.Count - 1
            Dim s As Integer = auStarts(i)
            Dim e As Integer = If(i + 1 < auStarts.Count, auStarts(i + 1), data.Length)
            Dim au(e - s - 1) As Byte
            Array.Copy(data, s, au, 0, e - s)
            Dim pts As Long = ptsBase + CLng(i) * ptsStep
            mux.QueueAu(ps, New AccessUnit() With {
                .Data = au, .Pts = pts, .Dts = pts, .DecodeTime27 = pts * 300L,
                .IsRandomAccessPoint = False
            })
        Next
    End Sub

    ' AC-3 syncframe size in bytes from the fscod/frmsizecod byte (A/52 table 5.18), 0 = invalid
    Private Function Ac3FrameBytes(fscodFrmsizecod As Byte) As Integer
        Dim fscod As Integer = fscodFrmsizecod >> 6
        Dim code As Integer = fscodFrmsizecod And &H3F
        If code > 37 Then Return 0
        Dim kbps As Integer = {32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320,
                               384, 448, 512, 576, 640}(code >> 1)
        Select Case fscod
            Case 0 : Return kbps * 4                                       ' 48 kHz
            Case 1 : Return ((kbps * 96000) \ 44100 + (code And 1)) * 2     ' 44.1 kHz (odd codes carry one extra word)
            Case 2 : Return kbps * 6                                       ' 32 kHz
            Case Else : Return 0
        End Select
    End Function

    ' first offset >= startAt holding a real syncframe: valid header, and the next frame (or EOF) right after it
    Private Function FindAc3Sync(data As Byte(), startAt As Integer) As Integer
        For i As Integer = Math.Max(0, startAt) To data.Length - 7
            If data(i) <> &HB OrElse data(i + 1) <> &H77 Then Continue For
            Dim size As Integer = Ac3FrameBytes(data(i + 4))
            If size = 0 OrElse (data(i + 5) >> 3) > 10 Then Continue For      ' bsid > 10 = E-AC-3 / not AC-3
            Dim nxt As Integer = i + size
            If nxt >= data.Length - 1 OrElse (data(nxt) = &HB AndAlso data(nxt + 1) = &H77) Then Return i
        Next
        Return -1
    End Function

    ' lfeon bit of the bsi: after acmod come cmixlev (3-front modes), surmixlev (surround modes), dsurmod (2/0)
    Private Function Ac3LfeOn(data As Byte(), frame As Integer) As Integer
        Dim bits As Integer = (CInt(data(frame + 6)) << 8) Or data(frame + 7)
        Dim acmod As Integer = (bits >> 13) And 7
        Dim pos As Integer = 3
        If (acmod And 1) <> 0 AndAlso acmod <> 1 Then pos += 2
        If (acmod And 4) <> 0 Then pos += 2
        If acmod = 2 Then pos += 2
        Return (bits >> (15 - pos)) And 1
    End Function

    ' LPCM: read WAV, swap LE -> BE, emit fixed-size AUs

    Private Sub RegisterAndQueueLpcm(mux As PamfMuxer, f As MuxInput)
        ' parse WAV header
        Dim data As Byte() = File.ReadAllBytes(f.Path)
        Dim sampleRate As Integer = 48000
        Dim channels As Integer = 2
        Dim bitsPerSample As Integer = 16
        Dim dataOff As Integer = -1
        Dim dataLen As Integer = 0
        Dim pos As Integer = 12
        While pos < data.Length - 8
            Dim chunkId As String = System.Text.Encoding.ASCII.GetString(data, pos, 4)
            Dim chunkSize As Integer = BitConverter.ToInt32(data, pos + 4)
            If chunkId = "fmt " Then
                channels = BitConverter.ToUInt16(data, pos + 8 + 2)
                sampleRate = BitConverter.ToInt32(data, pos + 8 + 4)
                bitsPerSample = BitConverter.ToUInt16(data, pos + 8 + 14)
            ElseIf chunkId = "data" Then
                dataOff = pos + 8
                dataLen = chunkSize
                Exit While
            End If
            pos += 8 + chunkSize
            If (chunkSize And 1) = 1 Then pos += 1
        End While
        If dataOff < 0 Then
            Throw New InvalidDataException("No data chunk in WAV: " & f.Path)
        End If

        Dim ps As PamfMuxStream = mux.AddLpcmStream(
            sampleRate:=sampleRate, numChannels:=CByte(channels),
            bitsPerSample:=bitsPerSample)

        ' PAMF LPCM AU is one 5 ms audio frame, 240 samples at 48 kHz
        ' au_size = sample_rate * bits_per_sample / 8 * padded_channels / 200
        '
        ' where padded_channels = channels + (channels & 1).
        '
        Dim bps As Integer = bitsPerSample \ 8
        Dim paddedChannels As Integer = channels + (channels And 1)
        Dim samplesPerAu As Integer = sampleRate \ 200
        Dim wavStride As Integer = channels * bps         ' bytes per sample-time in the input WAV
        Dim wireStride As Integer = paddedChannels * bps  ' bytes per sample-time on the PAMF wire
        Dim auBytes As Integer = samplesPerAu * wireStride
        Dim ptsBase As Long = 90000L
        Dim auIndex As Integer = 0
        Dim p As Integer = dataOff
        Dim [end] As Integer = dataOff + dataLen

        While p < [end]
            ' emit at most `samplesPerAu`, final AU may be shorter if WAV doesn't end on an AU boundary
            Dim wavRemain As Integer = [end] - p
            Dim samplesInAu As Integer = Math.Min(samplesPerAu, wavRemain \ wavStride)
            If samplesInAu <= 0 Then Exit While   ' partial trailing sample: drop

            Dim wireBytes As Integer = samplesInAu * wireStride
            Dim au(wireBytes - 1) As Byte
            Dim src As Integer = p
            Dim dst As Integer = 0
            For n As Integer = 0 To samplesInAu - 1
                ' real channels: byte-swap each channel sample WAV LE -> wire BE
                For ch As Integer = 0 To channels - 1
                    For b As Integer = 0 To bps - 1
                        au(dst + b) = data(src + bps - 1 - b)
                    Next
                    src += bps
                    dst += bps
                Next
                ' padded silence channel: array is already zero-initialised, just advance dst
                dst += (paddedChannels - channels) * bps
            Next
            p += samplesInAu * wavStride

            Dim pts As Long = ptsBase + (CLng(auIndex) * samplesPerAu * 90000L) \ CLng(sampleRate)
            mux.QueueAu(ps, New AccessUnit() With {
                .Data = au, .Pts = pts, .Dts = pts, .DecodeTime27 = pts * 300L,
                .IsRandomAccessPoint = False
            })
            auIndex += 1
        End While
    End Sub

End Module
