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
    ' readers for input files, AU payloads are read from them while packs are written
    Private ReadOnly _sources As New List(Of AuPayloadSource)()

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
                   endSector As Boolean,
                   sonyClock As Boolean)
        If positional.Count < 2 Then
            Console.Error.WriteLine("Usage: PAMFtool -mux <inputDir> <output.pamf> [-noep] [-noatsc] [-deblock | -nodeblock] [-pstd <KB>] [-mmb <n>] [-ps2-block <N>] [-muxrate <kbps>] [-std-delay <ticks>] [-initial-scr <ticks>] [-endsector] [-pamfmux]")
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
        If sonyClock Then Console.WriteLine("  -pamfmux: SCR origin 0 and mid-block video timestamps as pamfmux does")
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
                           overrideStdDelayTicks, overrideInitialScr, endSector, sonyClock)
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
                                   endSector As Boolean,
                                   sonyClock As Boolean)
        Try
            BuildPamf(files, outPath, noEp, forceDeblock, forceNoDeblock, noAtsc,
                      overridePstdKb, overrideMmb, ps2MinPictures, overrideMuxRateBps,
                      overrideStdDelayTicks, overrideInitialScr, endSector, sonyClock)
        Finally
            For Each src In _sources
                src.Dispose()
            Next
            _sources.Clear()
        End Try
    End Sub

    Private Sub BuildPamf(files As List(Of MuxInput), outPath As String,
                          noEp As Boolean, forceDeblock As Boolean,
                          forceNoDeblock As Boolean, noAtsc As Boolean,
                          overridePstdKb As Integer,
                          overrideMmb As Integer,
                          ps2MinPictures As Integer,
                          overrideMuxRateBps As Integer,
                          overrideStdDelayTicks As Integer,
                          overrideInitialScr As Long,
                          endSector As Boolean,
                          sonyClock As Boolean)
        Dim mux As New PamfMuxer()
        mux.SkipEpTable = noEp
        mux.PsMuxer.WriteEndSector = endSector
        mux.PsMuxer.SonyClock = sonyClock
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
        ' -pamfmux: 0 for every stream type, as Sony's command-line muxer (its clock starts at 0)
        ' -initial-scr keeps accepting the old pack-0 values: 30 -> origin 0, 30030 -> origin 30000, and wins over -pamfmux
        Dim origin As Long = If(sonyClock OrElse _firstVideoIsHighProfile.GetValueOrDefault(False), 0L, 30000L)
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
        ' one streamed pass: AU boundaries + per-AU analysis, AU bytes stay in file until packs are written
        Dim scan As AvcFileScan = ScanAvcFile(f.Path)
        Dim auStarts As List(Of Long) = scan.AuStarts
        Dim an As AvcStreamAnalysis = scan.Analysis
        Dim fileLen As Long = scan.FileLength
        Dim src As New FilePayloadSource(f.Path)
        _sources.Add(src)
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
            Dim e As Long = If(i + 1 < auStarts.Count, auStarts(i + 1), fileLen)
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
            Dim pic As AvcPictureInfo = an.Pictures(i)
            Dim a As New AccessUnit() With {
                .Source = src,
                .SourceOffset = auStarts(i),
                .SourceLength = CInt(auSizes(i)),
                .AnchorOffset = scan.Anchors(i),
                .IsRandomAccessPoint = pic.IsRap,
                .IsReferenceFrame = pic.IsReference,
                .PictureStructure = CByte(If(pic.FieldPic, If(pic.BottomField, 2, 1), 0))
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

    ' AVC file scan: AU boundaries (AuStarts, the last AU ends at FileLength), per-AU analysis and video anchors
    Private Class AvcFileScan
        Public ReadOnly AuStarts As New List(Of Long)()
        Public ReadOnly Anchors As New List(Of Integer)()
        Public Analysis As AvcStreamAnalysis
        Public FileLength As Long
    End Class

    ' H.264 AU boundary detection, streamed so >2 GB sources work (only the current AU and a read-ahead are in memory)
    '
    ' reliable boundary signals:
    '  - NAL type 9 (Access Unit Delimiter)
    '  - first VCL NAL of a picture has first_mb_in_slice == 0
    ' a non-VCL NAL after the VCL NALs of a picture starts next AU, if that AU's first VCL / AUD follows
    ' bytes before the first start code belong to no AU
    Private Function ScanAvcFile(path As String) As AvcFileScan
        Dim r As New AvcFileScan()
        Dim st As New AvcAnalyzer.AvcAnalyzerState()
        Using w As New FileWindow(path)
            Dim fileLen As Long = w.Length
            r.FileLength = fileLen
            Dim limit As Long = fileLen - 4             ' start code search positions are < limit
            Dim nalCount As Long = 0
            Dim curAuStart As Long = 0
            Dim haveVclInCurrentAu As Boolean = False
            Dim pendingNonVclStart As Long = -1
            Dim i As Long = 0

            While i < limit
                ' next 00 00 01 / 00 00 00 01 at or after i -> NAL header position
                ' window keeps current AU
                If nalCount > 0 Then w.KeepFrom = curAuStart
                Dim nalHdr As Long = -1
                Dim j As Long = i
                While j < limit
                    If j + 4 > w.Base + w.Count Then
                        If nalCount = 0 Then w.KeepFrom = Math.Max(0L, j - 8L)
                        w.Fill(j + 4)
                    End If
                    Dim buf As Byte() = w.Buf
                    Dim q As Integer = CInt(j - w.Base)
                    Dim qEnd As Integer = CInt(Math.Min(limit, w.Base + w.Count - 3) - w.Base)
                    While q < qEnd
                        If buf(q) = 0 AndAlso buf(q + 1) = 0 Then
                            If buf(q + 2) = 1 Then
                                nalHdr = w.Base + q + 3
                                Exit While
                            End If
                            If buf(q + 2) = 0 AndAlso buf(q + 3) = 1 Then
                                nalHdr = w.Base + q + 4
                                Exit While
                            End If
                        End If
                        q += 1
                    End While
                    If nalHdr >= 0 Then Exit While
                    j = w.Base + q
                End While
                If nalHdr < 0 Then Exit While

                ' start code position: 00 00 00 01 when 4 bytes before the header are that, else 00 00 01
                Dim startCodeAt As Long = nalHdr - 3
                If nalHdr >= 4 AndAlso w.At(nalHdr - 4) = 0 AndAlso w.At(nalHdr - 3) = 0 AndAlso
                   w.At(nalHdr - 2) = 0 AndAlso w.At(nalHdr - 1) = 1 Then
                    startCodeAt = nalHdr - 4
                End If
                Dim nt As Integer = w.At(nalHdr) And &H1F
                Dim isVcl As Boolean = (nt = 1 OrElse nt = 5)
                Dim isFirstSlice As Boolean = isVcl AndAlso
                    (nalHdr + 1 < fileLen) AndAlso ((w.At(nalHdr + 1) And &H80) <> 0)

                If nalCount = 0 Then
                    ' first AU starts at the first NAL
                    curAuStart = startCodeAt
                    haveVclInCurrentAu = isVcl
                Else
                    Dim isAud As Boolean = (nt = 9)
                    Dim startsNewPicture As Boolean =
                        (isAud AndAlso haveVclInCurrentAu) OrElse
                        (isVcl AndAlso isFirstSlice AndAlso haveVclInCurrentAu)
                    If startsNewPicture Then
                        Dim auStart As Long = If(pendingNonVclStart >= 0, pendingNonVclStart, startCodeAt)
                        AddScannedAvcAu(r, st, w, curAuStart, auStart)
                        curAuStart = auStart
                        haveVclInCurrentAu = isVcl
                        pendingNonVclStart = -1
                    ElseIf isVcl Then
                        haveVclInCurrentAu = True
                        pendingNonVclStart = -1
                    ElseIf haveVclInCurrentAu Then
                        ' non-VCL after a VCL, provisional next-AU start
                        If pendingNonVclStart < 0 Then pendingNonVclStart = startCodeAt
                    End If
                End If
                nalCount += 1
                i = nalHdr + 1
            End While

            ' a file without any NAL has no SPS either
            If nalCount = 0 Then
                Throw New InvalidDataException(
                    "Could not parse SPS from " & path & "; bitstream may not be Annex-B H.264.")
            End If
            w.KeepFrom = curAuStart
            AddScannedAvcAu(r, st, w, curAuStart, fileLen)
        End Using
        r.Analysis = AvcAnalyzer.Finish(st)
        Return r
    End Function

    ' AU = file bytes [a, b): analysis and anchor from bytes in the window
    Private Sub AddScannedAvcAu(r As AvcFileScan, st As AvcAnalyzer.AvcAnalyzerState, w As FileWindow, a As Long, b As Long)
        If b - a > Array.MaxLength Then
            Throw New InvalidDataException($"AVC access unit at byte {a:N0} is {b - a:N0} bytes long (over 2 GiB); refusing to allocate.")
        End If
        w.Fill(b)
        Dim off As Integer = CInt(a - w.Base)
        Dim len As Integer = CInt(b - a)
        AvcAnalyzer.AddAccessUnit(st, w.Buf, off, off + len)
        r.AuStarts.Add(a)
        r.Anchors.Add(Mpeg2PsMuxer.VideoAnchorOffset(w.Buf, off, len))
    End Sub

    ' read window over an input file: bytes from KeepFrom on stay in Buf, Buf(0) is file offset Base
    Private NotInheritable Class FileWindow
        Implements IDisposable

        Public Buf As Byte()
        Public Base As Long
        Public Count As Integer
        Public KeepFrom As Long
        Public ReadOnly Length As Long
        Private ReadOnly _fs As FileStream

        Public Sub New(path As String)
            _fs = New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16)
            Length = _fs.Length
            Buf = New Byte(CInt(Math.Max(16L, Math.Min(Length, 16L << 20))) - 1) {}
        End Sub

        ' make file bytes up to upTo available, reading ahead as far as Buf allows
        Public Sub Fill(upTo As Long)
            upTo = Math.Min(upTo, Length)
            If upTo <= Base + Count Then Return
            If KeepFrom < Base Then Throw New InvalidOperationException("read window: kept bytes were already dropped")
            If KeepFrom >= Base + Count Then
                Base = KeepFrom
                Count = 0
            ElseIf KeepFrom > Base Then
                Dim drop As Integer = CInt(KeepFrom - Base)
                Buffer.BlockCopy(Buf, drop, Buf, 0, Count - drop)
                Base = KeepFrom
                Count -= drop
            End If
            Dim need As Long = upTo - Base
            If need > Buf.Length Then
                If need > Array.MaxLength Then Throw New InvalidDataException("read window over 2 GiB (an access unit is too large)")
                Dim nb(CInt(Math.Min(CLng(Array.MaxLength), Math.Max(need, 2L * Buf.Length))) - 1) As Byte
                Buffer.BlockCopy(Buf, 0, nb, 0, Count)
                Buf = nb
            End If
            If _fs.Position <> Base + Count Then _fs.Position = Base + Count
            While Base + Count < upTo
                Dim want As Integer = CInt(Math.Min(CLng(Buf.Length - Count), Length - (Base + Count)))
                Dim got As Integer = _fs.Read(Buf, Count, want)
                If got <= 0 Then Throw New EndOfStreamException("input file shrank while it was read")
                Count += got
            End While
        End Sub

        Public Function At(pos As Long) As Byte
            If pos >= Length Then Throw New EndOfStreamException("read past the end of the input file")
            If pos >= Base + Count Then Fill(pos + 1)
            If pos < Base Then Throw New InvalidOperationException("read window: position already dropped")
            Return Buf(CInt(pos - Base))
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            _fs.Dispose()
        End Sub
    End Class

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
    '   pass 2: queue the AUs, their bytes are read from the file when the packs are written
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
    End Class

    ' MPEG-2 decode / presentation times (times in quarter frames, frame = 4):
    '   display time   : field 1/2 frame; frame 1, with repeat_first_field 1.5 (interlaced) or 2 / 3 (progressive_sequence, 3 with top_field_first)
    '   decode step    : B picture its own display timen I/P picture the display time of previous I/P picture
    '                    (interlaced I/P field: 1/2, or 1 when it follows the extra field of a repeat_first_field I/P frame)
    '   PTS            : B = its DTS; I/P = DTS of the next I/P picture, once at least a frame's worth of fields came since the
    '                    last one (so the two fields of an I/P pair are reordered together, the second at + the first's time)
    Private Sub M2vSonyTiming(pics As List(Of M2vPicture), progressiveSeq As Boolean, dtsQ As Long(), ptsQ As Long())
        Const D As Long = 4
        Dim n As Integer = pics.Count
        Dim dispQ(n - 1) As Long
        Dim prevStruct As Integer = 0, prevTff As Boolean = False, prevRff As Boolean = False   ' last saved I/P picture
        Dim lastFrame As Boolean = True
        Dim fields As Integer = 0
        Dim pend As New List(Of Integer)()
        Dim t As Long = 0
        For i As Integer = 0 To n - 1
            Dim p As M2vPicture = pics(i)
            Dim anchor As Boolean = (p.CodingType <> 3)
            Dim fieldPic As Boolean = (p.PictureStructure = 1 OrElse p.PictureStructure = 2)
            If anchor AndAlso fields > 1 Then
                M2vReorder(pend, t, ptsQ, dispQ, pics)
                pend.Clear()
                fields = fields Mod 2
            End If
            Dim disp As Long
            If progressiveSeq Then
                disp = If(p.RepeatFirstField, If(p.TopFieldFirst, 3 * D, 2 * D), D)
            ElseIf fieldPic Then
                disp = D \ 2
            Else
                disp = If(p.RepeatFirstField, D * 3 \ 2, D)
            End If
            Dim inc As Long
            If progressiveSeq Then
                If anchor Then
                    inc = If(prevRff, If(prevTff, 3 * D, 2 * D), D)
                Else
                    inc = If(p.RepeatFirstField, If(p.TopFieldFirst, 3 * D, 2 * D), D)
                End If
            ElseIf anchor Then
                If fieldPic Then
                    inc = If(Not lastFrame AndAlso prevStruct = 3 AndAlso prevRff, D, D \ 2)
                Else
                    inc = If(prevRff, D * 3 \ 2, D)
                End If
            Else
                inc = If(fieldPic, D \ 2, If(p.RepeatFirstField, D * 3 \ 2, D))
            End If
            dtsQ(i) = t
            ptsQ(i) = t
            dispQ(i) = disp
            t += inc
            pend.Add(i)
            If anchor AndAlso (p.PictureStructure <> 1 OrElse Not lastFrame) Then
                prevStruct = If(fieldPic, p.PictureStructure, 3)
                prevTff = p.TopFieldFirst
                prevRff = p.RepeatFirstField
            End If
            If progressiveSeq Then
                fields += 2
                lastFrame = True
            ElseIf Not fieldPic Then
                fields += If(p.RepeatFirstField, 3, 2)
                lastFrame = True
            Else
                fields += 1
                lastFrame = False
            End If
        Next
        M2vReorder(pend, t, ptsQ, dispQ, pics)
    End Sub

    ' I/P picture(s) at head of the pending list are shown at t (second field of a pair, right after the first)
    Private Sub M2vReorder(pend As List(Of Integer), t As Long, ptsQ As Long(), dispQ As Long(), pics As List(Of M2vPicture))
        If pend.Count >= 1 AndAlso pics(pend(0)).CodingType <> 3 Then ptsQ(pend(0)) = t
        If pend.Count >= 2 AndAlso pics(pend(1)).CodingType <> 3 Then ptsQ(pend(1)) = t + dispQ(pend(0))
    End Sub

    ' quarter frames -> 90 kHz ticks, rounded down (exact fraction)
    Private Function QuartersToTicks(q As Long, rate As FrameRate) As Long
        Return FloorDiv(CDec(q) * 90000D * CDec(rate.Den), 4D * CDec(rate.Num))
    End Function

    Private Sub RegisterAndQueueM2v(mux As PamfMuxer, f As MuxInput, overridePstdKb As Integer, ps2MinPictures As Integer)
        ' 1) peek at file head to parse the first sequence_header
        Dim fileLen As Long
        Using lfs As FileStream = File.OpenRead(f.Path)     ' length of file itself (FileInfo gives a symlink own size)
            fileLen = lfs.Length
        End Using
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

        ' picture timing as M2vSonyTiming, in quarter frames
        ' PTS of first picture shown = 90000
        ' field units (frame = 2, + repeat_first_field) give stream duration
        Dim nPic As Integer = pics.Count
        Dim totalUnits As Long = 0
        For k As Integer = 0 To nPic - 1
            Dim p As M2vPicture = pics(k)
            If p.PictureStructure <> 3 AndAlso p.PictureStructure <> 0 Then
                p.Units = 1
            ElseIf progressiveSeq Then
                p.Units = 2 * (1 + If(p.RepeatFirstField, 1, 0) + If(p.RepeatFirstField AndAlso p.TopFieldFirst, 1, 0))
            Else
                p.Units = 2 + If(p.RepeatFirstField, 1, 0)
            End If
            totalUnits += p.Units
        Next
        Dim dtsQ(nPic - 1) As Long
        Dim ptsQ(nPic - 1) As Long
        M2vSonyTiming(pics, progressiveSeq, dtsQ, ptsQ)
        Dim minPtsQ As Long = Long.MaxValue
        For k As Integer = 0 To nPic - 1
            minPtsQ = Math.Min(minPtsQ, ptsQ(k))
        Next

        ' 3) pass 2: queue each AU, bytes stay in file until packs are written
        ' anchor (first 00 00 01 of AU): every AU but the first starts with a start code, the first one is in the peek
        Dim blocks As Integer = 0, longBlocks As Integer = 0
        Dim blockStartK As Integer = 0
        Dim blockStartT27 As Long = 0
        Dim src As New FilePayloadSource(f.Path)
        _sources.Add(src)
        For i As Integer = 0 To nPic - 1
            Dim startOff As Long = pics(i).AuStart
            Dim endOff As Long = If(i + 1 < nPic, pics(i + 1).AuStart, fileLen)
            Dim sz As Long = endOff - startOff
            If sz <= 0 Then Throw New InvalidDataException($"MPEG-2 picture {i} has no data")
            If sz > Int32.MaxValue Then
                Throw New InvalidDataException(
                    "Picture " & i & " byte range exceeds 2 GiB (" & sz & " bytes); refusing to allocate.")
            End If

            Dim p As M2vPicture = pics(i)
            ' random access point = I picture whose AU carries a sequence header (Sony: one ps2 block per GOP)
            Dim isRap As Boolean = (p.CodingType = 1) AndAlso p.HasSequenceHeader
            Dim a As New AccessUnit() With {
                .Source = src,
                .SourceOffset = startOff,
                .SourceLength = CInt(sz),
                .AnchorOffset = If(i = 0, Mpeg2PsMuxer.VideoAnchorOffset(peek, 0, CInt(Math.Min(sz, CLng(peek.Length)))), 0),
                .IsRandomAccessPoint = isRap,
                .IsReferenceFrame = (p.CodingType = 1 OrElse p.CodingType = 2),
                .PictureCodingType = CByte(p.CodingType And 7),
                .PictureStructure = CByte(p.PictureStructure And 3)
            }
            a.Pts = 90000L + QuartersToTicks(ptsQ(i) - minPtsQ, rate)
            a.Dts = 90000L + QuartersToTicks(dtsQ(i) - minPtsQ, rate)
            a.DecodeTime27 = a.Dts * 300L
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
        _videoTimings.Add(New VideoTiming() With {.Rate = rate, .TotalUnits = totalUnits,
                                                  .EndPts90 = 90000L + UnitsToTicks(totalUnits, rate, 90000L)})
        Console.WriteLine($"  MPEG-2: {nPic:N0} pictures, {fps:F3} fps, DTS offset {(minPtsQ - dtsQ(0)) / 4.0:0.##} frame(s), {blocks:N0} ps2 blocks")
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
        Dim lay As At3Layout = ReadAt3Layout(f.Path, chs, sr, frameSize)
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
        ' each PAMF AU = 8-byte ATS header + raw_data_frame, built from file when the packs are written
        If noAtsc Then lay.Atsc = Nothing
        Dim src As New At3PayloadSource(f.Path, lay, frameSize, params)
        _sources.Add(src)
        For i As Long = 0 To lay.FrameCount - 1
            Dim pts As Long = ptsBase + (i * 2048L * 90000L) \ CLng(sr)
            mux.QueueAu(ps, New AccessUnit() With {
                .Source = src, .SourceOffset = lay.DataOffset + i * frameSize, .SourceLength = 8 + frameSize,
                .Pts = pts, .Dts = pts, .DecodeTime27 = pts * 300L,
                .IsRandomAccessPoint = False
            })
        Next
    End Sub

    ' PAMF ATRAC3plus AU: ATS header (sync 0x0FD0, codec parameters, frame extra_config_data) + raw frame
    ' SourceOffset = file offset of the frame
    Private NotInheritable Class At3PayloadSource
        Inherits AuPayloadSource

        Private ReadOnly _path As String
        Private ReadOnly _lay As At3Layout
        Private ReadOnly _frameSize As Integer
        Private ReadOnly _params As Integer
        Private _fs As FileStream

        Public Sub New(path As String, lay As At3Layout, frameSize As Integer, params As Integer)
            _path = path
            _lay = lay
            _frameSize = frameSize
            _params = params
        End Sub

        Protected Overrides Sub Load(au As AccessUnit, buf As Byte())
            buf(0) = &HF                               ' sync high byte
            buf(1) = &HD0                              ' sync low byte
            buf(2) = CByte((_params >> 8) And &HFF)    ' params high byte
            buf(3) = CByte(_params And &HFF)           ' params low byte
            ' extra_config_data: per frame, one value for all frames (legacy atsc), or zeros
            Dim k As Long = (au.SourceOffset - _lay.DataOffset) \ _frameSize
            Dim ecAt As Long = -1
            If _lay.Atsc IsNot Nothing Then ecAt = If(_lay.AtscBroadcast, 0L, k * 4L)
            For b As Integer = 0 To 3
                buf(4 + b) = If(ecAt >= 0, _lay.Atsc(CInt(ecAt) + b), CByte(0))
            Next
            If _fs Is Nothing Then _fs = New FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20)
            If _fs.Position <> au.SourceOffset Then _fs.Position = au.SourceOffset
            _fs.ReadExactly(buf, 8, _frameSize)
        End Sub

        Public Overrides Sub Dispose()
            If _fs IsNot Nothing Then _fs.Dispose()
            _fs = Nothing
        End Sub
    End Class

    ' .at3 layout: data chunk position, whole frames in it, the ATS extra_config_data source
    Private Class At3Layout
        Public DataOffset As Long
        Public FrameCount As Long
        Public Atsc As Byte()              ' Nothing = zeros
        Public AtscBroadcast As Boolean    ' one 4-byte value for every frame
    End Class

    Private Function ReadAt3Layout(path As String, ByRef channels As Integer,
                                   ByRef sampleRate As Integer,
                                   ByRef frameSize As Integer) As At3Layout
        ' parse RIFF/WAVE-style .at3 file:
        '  WAVE_FORMAT_EXTENSIBLE (0xFFFE) with AT3+ SubFormat GUID
        '  nBlockAlign = raw_data_frame size (ATS header stripped)
        '  optional "atsc" chunk of size N*4 OR 4 (see fallbacks below)
        ' only the chunk headers, fmt and atsc are read here, the frames are read when the packs are written
        ' chunk sizes are unsigned 32-bit, a data chunk that claims more than the file holds ends at the end of the file
        Using fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16)
            Dim fileLen As Long = fs.Length
            If fileLen < 12 Then Throw New InvalidDataException("Tiny .at3 file: " & path)
            Dim hdr(15) As Byte
            fs.ReadExactly(hdr, 0, 4)
            If hdr(0) <> &H52 OrElse hdr(1) <> &H49 _
            OrElse hdr(2) <> &H46 OrElse hdr(3) <> &H46 Then
                Throw New InvalidDataException("Not a RIFF file: " & path)
            End If
            Dim pos As Long = 12
            Dim dataOff As Long = -1
            Dim dataLen As Long = 0
            Dim atscOff As Long = -1
            Dim atscSize As Long = 0
            ' walk ALL chunks: the extractor writes atsc before the data chunk, older versions wrote it after
            While pos < fileLen - 8
                fs.Position = pos
                fs.ReadExactly(hdr, 0, 8)
                Dim chunkId As String = System.Text.Encoding.ASCII.GetString(hdr, 0, 4)
                Dim chunkSize As Long = BitConverter.ToUInt32(hdr, 4)
                If chunkId = "fmt " Then
                    ' offsets within fmt chunk: +2 channels, +4 sampleRate, +12 blockAlign
                    Dim fmt(15) As Byte
                    If ReadUpTo(fs, fmt, 16) < 14 Then Throw New InvalidDataException("Truncated fmt chunk in .at3: " & path)
                    channels = BitConverter.ToUInt16(fmt, 2)
                    sampleRate = BitConverter.ToInt32(fmt, 4)
                    frameSize = BitConverter.ToUInt16(fmt, 12)
                ElseIf chunkId = "atsc" Then
                    atscOff = pos + 8
                    atscSize = chunkSize
                ElseIf chunkId = "data" Then
                    dataOff = pos + 8
                    dataLen = Math.Min(chunkSize, fileLen - dataOff)
                End If
                pos += 8 + chunkSize
                If (chunkSize And 1) = 1 Then pos += 1
            End While
            If dataOff < 0 Then Throw New InvalidDataException("No data chunk in .at3: " & path)
            If frameSize <= 0 Then frameSize = 688

            Dim lay As New At3Layout() With {.DataOffset = dataOff, .FrameCount = dataLen \ frameSize}
            Dim n As Long = lay.FrameCount

            ' resolve extra_config_data source:
            '  - atsc chunk size == n * 4  -> per-frame list (preferred)
            '  - atsc chunk size == 4      -> broadcast to every frame (legacy)
            '  - anything else / no chunk  -> nothing (caller zero-fills)
            If atscOff >= 0 AndAlso n > 0 AndAlso atscOff + atscSize <= fileLen Then
                If atscSize = n * 4 OrElse atscSize = 4 Then
                    If atscSize > Array.MaxLength Then Throw New InvalidDataException("atsc chunk over 2 GiB in .at3: " & path)
                    lay.Atsc = New Byte(CInt(atscSize) - 1) {}
                    lay.AtscBroadcast = (atscSize = 4 AndAlso n <> 1)
                    fs.Position = atscOff
                    fs.ReadExactly(lay.Atsc, 0, lay.Atsc.Length)
                End If
            End If
            Return lay
        End Using
    End Function

    ' read up to count bytes (fewer only at end of file), return number
    Private Function ReadUpTo(fs As Stream, buf As Byte(), count As Integer) As Integer
        Dim total As Integer = 0
        While total < count
            Dim n As Integer = fs.Read(buf, total, count - total)
            If n <= 0 Then Exit While
            total += n
        End While
        Return total
    End Function

    ' AC-3: walk syncframes by their coded size (A/52 table 5.18), one AU per syncframe
    ' (searching for every 0x0B77 also splits frames where payload happens to contain these two bytes)

    Private Sub RegisterAndQueueAc3(mux As PamfMuxer, f As MuxInput)
        ' streamed: frame walk reads only around frame headers, frames are read when packs are written
        ' an AU is one syncframe of the size its header gives (A/52 table)
        ' bytes after the last frame or between frames are not muxed, only a truncated last frame keeps what is left of it
        Dim auStarts As New List(Of Long)()
        Dim auSizes As New List(Of Integer)()
        Dim fileLen As Long
        Dim channels As Integer
        Dim resyncs As Integer = 0
        Using w As New FileWindow(f.Path)
            fileLen = w.Length
            ' AC-3 with a 16-byte timecode header in front of every syncframe (the file starts with 0x01, each header too)
            ' strip the headers, so AUs are the syncframes after them
            Dim timecoded As Boolean = fileLen >= 21 AndAlso w.At(0) = &H1 AndAlso w.At(16) = &HB AndAlso w.At(17) = &H77
            Dim firstSync As Long
            If timecoded Then
                Dim q As Long = 0
                While q < fileLen AndAlso w.At(q) = &H1
                    w.KeepFrom = q
                    Dim fr As Long = q + 16
                    Dim size As Integer = 0
                    If fr + 5 <= fileLen AndAlso w.At(fr) = &HB AndAlso w.At(fr + 1) = &H77 Then size = Ac3FrameBytes(w.At(fr + 4))
                    If size = 0 Then
                        Throw New InvalidDataException($"AC-3 with timecode headers: no valid syncframe after the header at byte {q:N0} of {f.Path}")
                    End If
                    auStarts.Add(fr)
                    auSizes.Add(CInt(Math.Min(CLng(size), fileLen - fr)))
                    q = fr + size
                End While
                firstSync = auStarts(0)
                Console.WriteLine($"  AC-3 with timecode headers: {auStarts.Count:N0} syncframes, the 16-byte headers are left out")
            Else
                firstSync = FindAc3Sync(w, 0)
            End If
            If firstSync < 0 Then
                Throw New InvalidDataException("No AC-3 sync word found in " & f.Path)
            End If
            ' syncinfo: syncword (16), crc1 (16), fscod (2) + frmsizecod (6) at byte+4
            ' bsi: bsid (5) + bsmod (3) at byte+5, acmod (3) at the top of byte+6, then optional mix levels and lfeon
            w.KeepFrom = firstSync
            If (w.At(firstSync + 4) >> 6) <> 0 Then
                Console.WriteLine("  [warn] AC-3 stream is not 48 kHz; PAMF audio is 48 kHz and PTS assume 1536-sample frames at 48 kHz")
            End If
            Dim acmod As Integer = (w.At(firstSync + 6) >> 5) And 7
            Dim chTable As Integer() = {2, 1, 2, 3, 3, 4, 4, 5}    ' acmod 0..7 full-bandwidth channels
            channels = chTable(acmod) + Ac3LfeOn(w, firstSync)

            ' walk frame by frame, resync on the next valid frame if a header is broken
            Dim p As Long = If(timecoded, -1L, firstSync)
            While p >= 0 AndAlso p + 5 <= fileLen
                w.KeepFrom = p
                Dim size As Integer = 0
                If w.At(p) = &HB AndAlso w.At(p + 1) = &H77 Then size = Ac3FrameBytes(w.At(p + 4))
                If size > 0 Then
                    auStarts.Add(p)
                    auSizes.Add(CInt(Math.Min(CLng(size), fileLen - p)))   ' a truncated last frame keeps remaining bytes
                    p += size
                Else
                    p = FindAc3Sync(w, p + 1)
                    resyncs += 1
                End If
            End While
        End Using

        Dim ps As PamfMuxStream = mux.AddAc3Stream(numChannels:=CByte(channels))

        ' 1536 samples per syncframe = 2880 ticks at 48 kHz, first PTS 90000 (??)
        Dim ptsStep As Long = 1536L * 90000L \ 48000L
        Dim ptsBase As Long = 90000L

        If resyncs > 0 Then
            Console.WriteLine($"  [warn] AC-3: {resyncs} broken frame header(s), resynced on the next valid frame")
        End If
        Dim src As New FilePayloadSource(f.Path)
        _sources.Add(src)
        For i As Integer = 0 To auStarts.Count - 1
            Dim pts As Long = ptsBase + CLng(i) * ptsStep
            mux.QueueAu(ps, New AccessUnit() With {
                .Source = src, .SourceOffset = auStarts(i), .SourceLength = auSizes(i),
                .Pts = pts, .Dts = pts, .DecodeTime27 = pts * 300L,
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
    Private Function FindAc3Sync(w As FileWindow, startAt As Long) As Long
        Dim len As Long = w.Length
        For i As Long = Math.Max(0L, startAt) To len - 7
            w.KeepFrom = i
            If w.At(i) <> &HB OrElse w.At(i + 1) <> &H77 Then Continue For
            Dim size As Integer = Ac3FrameBytes(w.At(i + 4))
            If size = 0 OrElse (w.At(i + 5) >> 3) > 10 Then Continue For      ' bsid > 10 = E-AC-3 / not AC-3
            Dim nxt As Long = i + size
            If nxt >= len - 1 OrElse (w.At(nxt) = &HB AndAlso w.At(nxt + 1) = &H77) Then Return i
        Next
        Return -1
    End Function

    ' lfeon bit of the bsi: after acmod come cmixlev (3-front modes), surmixlev (surround modes), dsurmod (2/0)
    Private Function Ac3LfeOn(w As FileWindow, frame As Long) As Integer
        Dim bits As Integer = (CInt(w.At(frame + 6)) << 8) Or w.At(frame + 7)
        Dim acmod As Integer = (bits >> 13) And 7
        Dim pos As Integer = 3
        If (acmod And 1) <> 0 AndAlso acmod <> 1 Then pos += 2
        If (acmod And 4) <> 0 Then pos += 2
        If acmod = 2 Then pos += 2
        Return (bits >> (15 - pos)) And 1
    End Function

    ' LPCM: read WAV, swap LE -> BE, PS3 channel order, emit fixed-size AUs
    '
    ' WAV is streamed, only header is parsed here, every AU is converted when packs are written
    ' RIFF chunk sizes are unsigned 32-bit, RF64 (ds64) is read for WAVs over 4 GB,
    ' data chunk that claims more than the file holds (or 0xFFFFFFFF) goes to end of file

    Private Sub RegisterAndQueueLpcm(mux As PamfMuxer, f As MuxInput)
        ' parse WAV header
        Dim sampleRate As Integer = 48000
        Dim channels As Integer = 2
        Dim bitsPerSample As Integer = 16
        Dim formatTag As Integer = 1
        Dim pcmSubFormat As Boolean = True
        Dim dataOff As Long = -1
        Dim dataLen As Long = 0
        Dim fileLen As Long
        Using fs As New FileStream(f.Path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16)
            fileLen = fs.Length
            Dim ds64DataSize As Long = -1
            Dim hdr(7) As Byte
            Dim pos As Long = 12
            While pos < fileLen - 8
                fs.Position = pos
                fs.ReadExactly(hdr, 0, 8)
                Dim chunkId As String = System.Text.Encoding.ASCII.GetString(hdr, 0, 4)
                Dim chunkSize As Long = BitConverter.ToUInt32(hdr, 4)
                If chunkId = "fmt " Then
                    Dim fmt(39) As Byte
                    Dim got As Integer = ReadUpTo(fs, fmt, CInt(Math.Min(chunkSize, 40L)))
                    If got < 16 Then Throw New InvalidDataException("Truncated fmt chunk in WAV: " & f.Path)
                    formatTag = BitConverter.ToUInt16(fmt, 0)
                    channels = BitConverter.ToUInt16(fmt, 2)
                    sampleRate = BitConverter.ToInt32(fmt, 4)
                    bitsPerSample = BitConverter.ToUInt16(fmt, 14)
                    ' WAVE_FORMAT_EXTENSIBLE: SubFormat GUID at +24, KSDATAFORMAT_SUBTYPE_PCM start with 01 00 00 00
                    If formatTag = &HFFFE Then
                        pcmSubFormat = got >= 40 AndAlso BitConverter.ToUInt32(fmt, 24) = 1UI
                    End If
                ElseIf chunkId = "ds64" Then
                    ' RF64: u64 riffSize, u64 dataSize, u64 sampleCount, ...
                    Dim ds(15) As Byte
                    If ReadUpTo(fs, ds, 16) = 16 Then ds64DataSize = BitConverter.ToInt64(ds, 8)
                ElseIf chunkId = "data" Then
                    dataOff = pos + 8
                    dataLen = chunkSize
                    If chunkSize = &HFFFFFFFFL AndAlso ds64DataSize >= 0 Then dataLen = ds64DataSize
                    If dataLen > fileLen - dataOff Then
                        If chunkSize <> &HFFFFFFFFL Then
                            Console.WriteLine($"  [warn] WAV data chunk says {dataLen:N0} bytes, the file holds {fileLen - dataOff:N0}; using what is there")
                        End If
                        dataLen = fileLen - dataOff
                    End If
                    Exit While
                End If
                pos += 8 + chunkSize
                If (chunkSize And 1) = 1 Then pos += 1
            End While
        End Using
        If dataOff < 0 Then
            Throw New InvalidDataException("No data chunk in WAV: " & f.Path)
        End If

        ' PAMF LPCM
        If Not (formatTag = 1 OrElse (formatTag = &HFFFE AndAlso pcmSubFormat)) Then
            Throw New InvalidDataException($"WAV {Path.GetFileName(f.Path)} is not integer PCM (format 0x{formatTag:X4}); PAMF LPCM needs 16- or 24-bit integer PCM.")
        End If
        If sampleRate <> 48000 Then
            Throw New InvalidDataException($"WAV {Path.GetFileName(f.Path)} is {sampleRate} Hz; PAMF LPCM is 48 kHz only.")
        End If
        If bitsPerSample <> 16 AndAlso bitsPerSample <> 24 Then
            Throw New InvalidDataException($"WAV {Path.GetFileName(f.Path)} has {bitsPerSample}-bit samples; PAMF LPCM is 16- or 24-bit.")
        End If
        If channels <> 1 AndAlso channels <> 2 AndAlso channels <> 6 AndAlso channels <> 8 Then
            Throw New InvalidDataException($"WAV {Path.GetFileName(f.Path)} has {channels} channels; PAMF LPCM is mono, stereo, 5.1 or 7.1.")
        End If

        Dim ps As PamfMuxStream = mux.AddLpcmStream(
            sampleRate:=sampleRate, numChannels:=CByte(channels),
            bitsPerSample:=bitsPerSample)

        ' PAMF LPCM AU is one 5 ms audio frame, 240 samples at 48 kHz
        ' au_size = sample_rate * bits_per_sample / 8 * padded_channels / 200
        '
        ' where padded_channels = channels + (channels & 1).
        '
        ' a WAV that does not end on an AU boundary gets its last AU filled up with silence,
        ' as Sony's encoder pads its output to whole AUs (a partial trailing sample is dropped)
        Dim bps As Integer = bitsPerSample \ 8
        Dim paddedChannels As Integer = channels + (channels And 1)
        Dim samplesPerAu As Integer = sampleRate \ 200
        Dim wavStride As Integer = channels * bps         ' bytes per sample-time in the input WAV
        Dim wireStride As Integer = paddedChannels * bps  ' bytes per sample-time on the PAMF wire
        Dim auBytes As Integer = samplesPerAu * wireStride
        Dim totalSamples As Long = dataLen \ wavStride
        Dim auCount As Long = (totalSamples + samplesPerAu - 1) \ samplesPerAu
        Dim ptsBase As Long = 90000L

        Dim src As New LpcmPayloadSource(f.Path, dataOff, dataOff + totalSamples * wavStride, channels, bps)
        _sources.Add(src)
        For auIndex As Long = 0 To auCount - 1
            Dim pts As Long = ptsBase + (auIndex * samplesPerAu * 90000L) \ CLng(sampleRate)
            mux.QueueAu(ps, New AccessUnit() With {
                .Source = src, .SourceOffset = dataOff + auIndex * samplesPerAu * wavStride, .SourceLength = auBytes,
                .Pts = pts, .Dts = pts, .DecodeTime27 = pts * 300L,
                .IsRandomAccessPoint = False
            })
        Next
    End Sub

    ' one LPCM AU from WAV samples at SourceOffset:
    '   samples byte-swapped to big-endian
    '   channels in PS3 order
    '   mono padded with silent second channel, samples past end of WAV data are silence
    Private NotInheritable Class LpcmPayloadSource
        Inherits AuPayloadSource

        Private ReadOnly _path As String
        Private ReadOnly _dataOff As Long
        Private ReadOnly _dataEnd As Long
        Private ReadOnly _channels As Integer
        Private ReadOnly _bps As Integer
        Private ReadOnly _map As Integer()      ' WAV channel of each wire channel, -1 = padding
        Private _fs As FileStream
        Private _wav As Byte() = Array.Empty(Of Byte)()

        Public Shared Function WireFromWav(channels As Integer) As Integer()
            Select Case channels
                Case 1 : Return {0, -1}
                Case 6 : Return {0, 1, 2, 4, 5, 3}
                Case 8 : Return {0, 1, 2, 6, 4, 5, 7, 3}
                Case Else
                    Dim m(channels - 1) As Integer
                    For k As Integer = 0 To channels - 1
                        m(k) = k
                    Next
                    Return m
            End Select
        End Function

        Public Sub New(path As String, dataOff As Long, dataEnd As Long, channels As Integer, bytesPerSample As Integer)
            _path = path
            _dataOff = dataOff
            _dataEnd = dataEnd
            _channels = channels
            _bps = bytesPerSample
            _map = WireFromWav(channels)
        End Sub

        Protected Overrides Sub Load(au As AccessUnit, buf As Byte())
            Dim wavStride As Integer = _channels * _bps
            Dim wireStride As Integer = _map.Length * _bps
            Dim samples As Integer = au.SourceLength \ wireStride
            Dim avail As Integer = CInt(Math.Max(0L, Math.Min(CLng(samples), (_dataEnd - au.SourceOffset) \ wavStride)))
            Dim wavBytes As Integer = avail * wavStride
            If _wav.Length < wavBytes Then _wav = New Byte(wavBytes - 1) {}
            If wavBytes > 0 Then
                If _fs Is Nothing Then _fs = New FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20)
                If _fs.Position <> au.SourceOffset Then _fs.Position = au.SourceOffset
                _fs.ReadExactly(_wav, 0, wavBytes)
            End If
            Dim dst As Integer = 0
            For n As Integer = 0 To avail - 1
                Dim srcBase As Integer = n * wavStride
                For k As Integer = 0 To _map.Length - 1
                    Dim c As Integer = _map(k)
                    If c < 0 Then
                        For b As Integer = 0 To _bps - 1
                            buf(dst + b) = 0
                        Next
                    Else
                        Dim src As Integer = srcBase + c * _bps
                        For b As Integer = 0 To _bps - 1
                            buf(dst + b) = _wav(src + _bps - 1 - b)
                        Next
                    End If
                    dst += _bps
                Next
            Next
            ' silence after end of WAV data (last AU)
            If dst < au.SourceLength Then Array.Clear(buf, dst, au.SourceLength - dst)
        End Sub

        Public Overrides Sub Dispose()
            If _fs IsNot Nothing Then _fs.Dispose()
            _fs = Nothing
        End Sub
    End Class

End Module
