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

    Public Sub Run(positional As List(Of String),
                   noEp As Boolean,
                   forceDeblock As Boolean,
                   forceNoDeblock As Boolean,
                   noAtsc As Boolean,
                   paceMbps As Double,
                   overridePstdKb As Integer,
                   overrideMmb As Integer,
                   ps2FramesPerBlock As Integer,
                   overrideMuxRateBps As Integer,
                   overrideStdDelayTicks As Integer,
                   overrideInitialScr As Long)
        If positional.Count < 2 Then
            Console.Error.WriteLine("Usage: PamfExtractor -mux <inputDir> <output.pamf> [-noep] [-noatsc] [-deblock | -nodeblock] [-pace <Mbps|auto|off>] [-pstd <KB>] [-mmb <n>] [-ps2-block <N>] [-muxrate <kbps>]")
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
        If overridePstdKb > 0 Then Console.WriteLine($"  -pstd {overridePstdKb}: overriding AVC P-STD buffer size to {overridePstdKb} KB")
        If overrideMmb >= 0 Then Console.WriteLine($"  -mmb {overrideMmb}: overriding max_mean_bitrate byte to {overrideMmb}")
        If overrideMuxRateBps > 0 Then Console.WriteLine($"  -muxrate {overrideMuxRateBps \ 1000} kbps: overriding pack_header mux_rate (Sony spec: 48000/24000/12000)")
        If overrideInitialScr >= 0 Then Console.WriteLine($"  -initial-scr {overrideInitialScr}: SCR at pack 0 (Sony uses 30 for logos, 30030 for game content)")
        If ps2FramesPerBlock = 0 Then
            Console.WriteLine("  -ps2-block 0: emitting legacy 4-byte ps2 marker at every IDR")
        ElseIf ps2FramesPerBlock < 0 Then
            Console.WriteLine("  ps2 block N: auto-detect from AVC SPS interval (fallback 12)")
        Else
            Console.WriteLine($"  ps2 cadence: {ps2FramesPerBlock}-frame block, {18 + 4 * ps2FramesPerBlock}-byte payload every {ps2FramesPerBlock} AUs")
        End If
        If paceMbps > 0 Then
            Console.WriteLine($"  -pace {paceMbps} Mbps: SCR pacing at fixed rate (mux_rate field unchanged)")
        ElseIf paceMbps < 0 Then
            Console.WriteLine("  -pace off (default): SCR advances at mux_rate")
        Else
            Console.WriteLine("  -pace auto: SCR pacing derived from measured content bitrate")
        End If

        Console.WriteLine("Muxing to: " & outPath)
        BuildPamfFromFiles(files, outPath, noEp, forceDeblock, forceNoDeblock, noAtsc, paceMbps,
                           overridePstdKb, overrideMmb, ps2FramesPerBlock, overrideMuxRateBps,
                           overrideStdDelayTicks, overrideInitialScr)
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
                                   paceMbps As Double,
                                   overridePstdKb As Integer,
                                   overrideMmb As Integer,
                                   ps2FramesPerBlock As Integer,
                                   overrideMuxRateBps As Integer,
                                   overrideStdDelayTicks As Integer,
                                   overrideInitialScr As Long)
        Dim mux As New PamfMuxer()
        mux.SkipEpTable = noEp
        mux.PsMuxer.Ps2FramesPerBlock = ps2FramesPerBlock
        If overrideMuxRateBps > 0 Then
            mux.PsMuxer.MuxRateBps = overrideMuxRateBps
            ' Also default EffectiveDeliveryBps to the same so SCR advances match the
            ' new mux_rate unless the user separately specifies -pace.
            mux.PsMuxer.EffectiveDeliveryBps = overrideMuxRateBps
        End If
        If overrideStdDelayTicks > 0 Then
            ' User-supplied override wins over any HRD-based auto-detect that happens
            ' inside RegisterAndQueueAvc.
            mux.StdDelayBoundTicks = overrideStdDelayTicks
        End If
        If overrideInitialScr >= 0L Then
            mux.PsMuxer.InitialScr = overrideInitialScr
        End If

        Dim userSetInitialScr As Boolean = (overrideInitialScr >= 0L)

        ' for each input, parse codec metadata, register the stream, then slice into AUs and queue
        For Each f In files
            Select Case f.Kind
                Case "avc" : RegisterAndQueueAvc(mux, f, forceDeblock, forceNoDeblock, overridePstdKb, overrideMmb, userSetInitialScr)
                Case "mpeg2" : RegisterAndQueueM2v(mux, f, userSetInitialScr)
                Case "atrac3plus" : RegisterAndQueueAt3p(mux, f, noAtsc)
                Case "ac3" : RegisterAndQueueAc3(mux, f)
                Case "lpcm" : RegisterAndQueueLpcm(mux, f)
            End Select
        Next

        ' resolve the SCR pacing rate, 3 modes:
        '  paceMbps > 0  : fixed user value (e.g. -pace 7 -> 7 Mbps)
        '  paceMbps == 0 : "auto" - derive from measured content bitrate over playback duration
        '  paceMbps < 0  : "off"  - advance SCR at mux_rate (legacy front-loaded delivery)
        ' whatever we resolve here is passed to PsMuxer.EffectiveDeliveryBps
        Dim effectiveBps As Integer
        If paceMbps < 0 Then
            effectiveBps = mux.PsMuxer.MuxRateBps
        ElseIf paceMbps > 0 Then
            effectiveBps = CInt(paceMbps * 1_000_000)
        Else
            effectiveBps = ComputeAutoPaceBps(mux)
            Console.WriteLine($"  auto pace: {effectiveBps / 1_000_000.0:F2} Mbps (from measured content bitrate)")
        End If
        ' Clamp to a sensible floor - never pace slower than 1 Mbps or SCR span
        ' explodes into the multi-hour range for short clips, and hardware may not
        ' cope with unrealistic delivery clocks.
        If effectiveBps < 1_000_000 Then effectiveBps = 1_000_000
        If effectiveBps > mux.PsMuxer.MuxRateBps Then effectiveBps = mux.PsMuxer.MuxRateBps
        mux.PsMuxer.EffectiveDeliveryBps = effectiveBps

        Using outFs As New FileStream(outPath, FileMode.Create, FileAccess.Write)
            mux.WritePamfFile(outFs)
        End Using
    End Sub

    ' compute an SCR pacing rate
    ' rate is derived from the QUEUED AU payload bytes divided by playback duration
    Private Function ComputeAutoPaceBps(mux As PamfMuxer) As Integer
        Dim totalPayloadBytes As Long = 0
        Dim maxPts As Long = 0
        Dim minPts As Long = Long.MaxValue
        For Each s In mux.PsMuxer.Streams
            For Each au In s.AuQueue
                totalPayloadBytes += au.Data.LongLength
                If au.Pts < minPts Then minPts = au.Pts
                If au.Pts > maxPts Then maxPts = au.Pts
            Next
        Next
        If maxPts <= minPts OrElse totalPayloadBytes <= 0 Then
            Return mux.PsMuxer.MuxRateBps
        End If
        Dim durationSec As Double = (maxPts - minPts) / 90000.0
        If durationSec <= 0.1 Then
            Return mux.PsMuxer.MuxRateBps
        End If
        ' bytes per second from content alone, no overhead multiplier
        Return CInt(Math.Ceiling((totalPayloadBytes * 8.0) / durationSec))
    End Function

    ' AVC: parse SPS, slice into per-frame AU, generate PTS/DTS

    Private Sub RegisterAndQueueAvc(mux As PamfMuxer, f As MuxInput,
                                    forceDeblock As Boolean,
                                    forceNoDeblock As Boolean,
                                    overridePstdKb As Integer,
                                    overrideMmb As Integer,
                                    userSetInitialScr As Boolean)
        Dim bytes As Byte() = File.ReadAllBytes(f.Path)
        Dim sps As H264SpsInfo = H264SpsParser.ParseFirstSps(bytes, 0, bytes.Length)
        If sps Is Nothing Then
            Throw New InvalidDataException(
                "Could not parse SPS from " & f.Path & "; bitstream may not be Annex-B H.264.")
        End If
        Dim pps As H264PpsInfo = H264PpsParser.ParseFirstPps(bytes, 0, bytes.Length)
        Dim fps As Double = sps.FrameRate
        If fps <= 0.0 Then fps = 30000.0 / 1001.0   ' fallback 29.97

        Dim frameRateCode As Byte = AvcFrameRateCodeFromFps(fps)
        Dim mbsOnly As Byte = CByte(If(sps.FrameMbsOnlyFlag, 1, 0))
        Dim arIdc As Byte = If(sps.AspectRatioIdc > 0, sps.AspectRatioIdc, CByte(1))
        Dim hasVsi As Byte = CByte(If(sps.HasVideoSignalInfo, 1, 0))
        Dim videoFormat As Byte = If(sps.HasVideoSignalInfo, sps.VideoFormat, CByte(5))
        Dim fullRange As Byte = If(sps.HasVideoSignalInfo, sps.VideoFullRangeFlag, CByte(0))
        Dim colourPrim As Byte = If(sps.ColourPrimaries > 0, sps.ColourPrimaries, CByte(1))
        Dim transferCh As Byte = If(sps.TransferCharacteristics > 0, sps.TransferCharacteristics, CByte(1))
        Dim matrixCoeff As Byte = If(sps.MatrixCoefficients > 0, sps.MatrixCoefficients, CByte(1))
        Dim cabac As Byte = If(pps IsNot Nothing AndAlso pps.CabacFlag, CByte(1), CByte(0))
        ' PAMF codec_info deblockingFilterFlag byte reports whether stream uses in-loop deblocking filter
        ' try to read from first slice, fall back to 0 if we can't
        Dim deblock As Byte = InferAvcDeblockFilterEnabled(bytes, sps, pps)
        If forceDeblock Then deblock = 1
        If forceNoDeblock Then deblock = 0

        ' adaptive std_delay
        ' - if SPS carried HRD parameters, we know the encoder-declared peak bit rate
        ' - absent HRD, keep the safe default 90000 (1 s)
        If sps.HasHrdParameters AndAlso mux.StdDelayBoundTicks = 90000 Then
            If sps.HrdPeakBitrateBps > 30_000_000L Then
                mux.StdDelayBoundTicks = 67500
                Console.WriteLine($"  AVC HRD peak bit rate {sps.HrdPeakBitrateBps / 1_000_000.0:F1} Mbps > 30 Mbps -> std_delay = 0.75 s (67500 ticks) per Sony spec")
            End If
        End If

        ' auto-detect ps2 block N from SPS interval when user didn't pass -ps2-block
        ' Sony inserts SPS at each seek-block boundary and ps2 marker N field matches that interval
        ' fall back to 12 if stream doesn't have consistent SPS cadence
        If mux.PsMuxer.Ps2FramesPerBlock < 0 Then
            Dim detected As Integer = DetectAvcSpsInterval(bytes)
            If detected > 0 Then
                mux.PsMuxer.Ps2FramesPerBlock = detected
                Console.WriteLine($"  ps2 block N auto-detected from AVC SPS interval: {detected}")
            Else
                mux.PsMuxer.Ps2FramesPerBlock = 12
                Console.WriteLine("  ps2 block N: could not detect SPS cadence, defaulting to 12")
            End If
        End If

        ' auto-detect InitialScr (SCR value at pack 0) from AVC profile
        ' AVC High profile varies but 30 is verified in byte-identical files
        ' so High profile keeps 30 unconditionally
        If Not userSetInitialScr Then
            Dim isHighProfileFamily As Boolean =
                sps.ProfileIdc = 100 OrElse sps.ProfileIdc = 110 OrElse sps.ProfileIdc = 122 OrElse
                sps.ProfileIdc = 244 OrElse sps.ProfileIdc = 44 OrElse sps.ProfileIdc = 83 OrElse
                sps.ProfileIdc = 86 OrElse sps.ProfileIdc = 118 OrElse sps.ProfileIdc = 128 OrElse
                sps.ProfileIdc = 138 OrElse sps.ProfileIdc = 139 OrElse sps.ProfileIdc = 134 OrElse
                sps.ProfileIdc = 135
            If Not isHighProfileFamily Then
                mux.PsMuxer.InitialScr = 30030L
                Console.WriteLine("  AVC Main/Baseline/Extended profile detected -> InitialScr = 30030 (Sony's game-content SCR0 policy)")
            End If
            ' High profile: leave InitialScr at whatever it already is (30 by default)
        End If
        ' default max_mean_bitrate byte
        ' observed values:
        '   L3.1 (720p CAVLC) :      5
        '   L4.1 (1080p CABAC):     11
        '   L4.1 SD (720x480) :      1  (bitrate class overrides level tier)
        ' derive from AVC level, or user override via -mmb
        ' value 0 (previous default) is legal but suboptimal
        Dim mmb As Byte = DefaultMaxMeanBitrateForAvcLevel(sps.LevelIdc, sps.WidthPixels, sps.HeightPixels)
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

        ' apply user overrides
        If overridePstdKb > 0 Then
            ps.PstdBufferSize = overridePstdKb
            mux.HeaderWriter.OverrideLastAvcPstd(overridePstdKb)
        End If
        ' MMB: overrideMmb >= 0 means user passed -mmb, use that value
        ' itherwise use level-derived default we computed above
        If overrideMmb >= 0 Then
            mux.HeaderWriter.OverrideLastAvcMaxMeanBitrate(CByte(overrideMmb And &HFF))
        Else
            mux.HeaderWriter.OverrideLastAvcMaxMeanBitrate(mmb)
        End If

        ' slice into per-frame AU at VCL NAL boundaries (types 1, 5)
        ' each AU starts at a NALU, ends just before the next VCL NALU
        Dim auStarts As List(Of Integer) = FindAvcAuStarts(bytes)
        Dim tickPerFrame As Long = CLng(90000.0 / fps)
        Dim ptsBase As Long = 90000L
        ' DTS lags PTS by exactly one frame
        For i As Integer = 0 To auStarts.Count - 1
            Dim s As Integer = auStarts(i)
            Dim e As Integer = If(i + 1 < auStarts.Count, auStarts(i + 1), bytes.Length)
            Dim au(e - s - 1) As Byte
            Array.Copy(bytes, s, au, 0, e - s)
            Dim pts As Long = ptsBase + CLng(i) * tickPerFrame
            Dim dts As Long = pts - tickPerFrame
            mux.QueueAu(ps, New AccessUnit() With {
                .Data = au, .Pts = pts, .Dts = dts,
                .IsRandomAccessPoint = AvcAuContainsIdr(au),
                .IsReferenceFrame = AvcAuFirstVclIsReference(au)
            })
        Next
    End Sub

    ' AVC "reference" detection: examine the FIRST VCL NAL unit in the AU
    '
    ' (nal_unit_type 1 = non-IDR slice or 5 = IDR slice)
    '
    ' read its nal_ref_idc field, which lives in bits 5-6 of the NAL header byte
    ' nal_ref_idc == 0 means the picture is not used as a reference for later frames
    ' nonzero means it is a reference (I,P, or reference B)
    Private Function AvcAuFirstVclIsReference(au As Byte()) As Boolean
        Dim i As Integer = 0
        While i < au.Length - 4
            If au(i) = 0 AndAlso au(i + 1) = 0 Then
                Dim hdr As Integer = -1
                If au(i + 2) = 1 Then hdr = i + 3
                If au(i + 2) = 0 AndAlso (i + 3) < au.Length AndAlso au(i + 3) = 1 Then hdr = i + 4
                If hdr >= 0 AndAlso hdr < au.Length Then
                    Dim nt As Integer = au(hdr) And &H1F
                    If nt = 1 OrElse nt = 5 Then
                        ' VCL NAL - inspect nal_ref_idc
                        Return ((au(hdr) >> 5) And &H3) <> 0
                    End If
                End If
            End If
            i += 1
        End While
        Return True
    End Function

    ' check if AVC stream uses in-loop deblocking filter
    ' oarse the first slice NAL disable_deblocking_filter_idc:
    '   idc = 0 = filter enabled                      = return 1
    '   idc = 1 = filter disabled                     = return 0
    '   idc = 2 = enabled but not at slice boundaries = return 1

    ' if PPS deblocking_filter_control_present_flag is 0, field isn't coded and its inferred value is 0 (filter enabled)
    ' if we can't parse header cleanly, fall back to 0
    ' only the FIRST slice is checked because codec_info deblock byte is a stream-level property
    Private Function InferAvcDeblockFilterEnabled(bytes As Byte(),
                                                  sps As H264SpsInfo,
                                                  pps As H264PpsInfo) As Byte
        If pps Is Nothing Then Return 0
        ' When per-slice deblock params are not coded, the value is inferred to 0
        ' (filter enabled) per H.264 §7.4.3
        If Not pps.DeblockingFilterControlPresent Then Return 1

        Dim idc As Integer = ParseFirstSliceDisableDeblockIdc(bytes, sps, pps)
        If idc < 0 Then Return 0     ' parse failure -> safe Sony default
        Return If(idc = 1, CByte(0), CByte(1))
    End Function

    ' - walk NAL units
    ' - find the first VCL slice (nal_unit_type 1 or 5)
    ' - parse its slice_header() up to disable_deblocking_filter_idc
    ' - return -1 on any parse failure so caller can fall back
    ' (handle H.264 slice header layout for I/P/B/SI/SP slices)
    Private Function ParseFirstSliceDisableDeblockIdc(bytes As Byte(),
                                                      sps As H264SpsInfo,
                                                      pps As H264PpsInfo) As Integer
        ' find first VCL NAL (type 1 = non-IDR slice, 5 = IDR slice)
        Dim nalStart As Integer = -1
        Dim nalType As Integer = 0
        Dim i As Integer = 0
        While i < bytes.Length - 4
            Dim sc As Integer = 0
            If bytes(i) = 0 AndAlso bytes(i + 1) = 0 AndAlso bytes(i + 2) = 1 Then
                sc = 3
            ElseIf i + 3 < bytes.Length AndAlso bytes(i) = 0 AndAlso bytes(i + 1) = 0 _
                   AndAlso bytes(i + 2) = 0 AndAlso bytes(i + 3) = 1 Then
                sc = 4
            End If
            If sc > 0 Then
                Dim hdr As Integer = i + sc
                If hdr < bytes.Length Then
                    Dim nt As Integer = bytes(hdr) And &H1F
                    If nt = 1 OrElse nt = 5 Then
                        nalStart = hdr + 1
                        nalType = nt
                        Exit While
                    End If
                End If
                i += sc
            Else
                i += 1
            End If
        End While
        If nalStart < 0 Then Return -1

        ' extract RBSP (drop 0x03 emulation prevention bytes) into a slice-header buffer
        ' we don't need whole slice, header is at most a few hundred bytes
        Dim maxHeaderBytes As Integer = Math.Min(bytes.Length - nalStart, 512)
        Dim rbsp As New List(Of Byte)(maxHeaderBytes)
        Dim zeroRun As Integer = 0
        For k As Integer = 0 To maxHeaderBytes - 1
            Dim b As Byte = bytes(nalStart + k)
            If zeroRun = 2 AndAlso b = 3 Then
                zeroRun = 0
                Continue For
            End If
            rbsp.Add(b)
            zeroRun = If(b = 0, zeroRun + 1, 0)
        Next
        If rbsp.Count = 0 Then Return -1

        Try
            Dim br As New SliceBitReader(rbsp.ToArray())
            br.Ue()                                                            ' first_mb_in_slice
            Dim sliceTypeRaw As UInteger = br.Ue()                             ' slice_type
            Dim sliceType As Integer = CInt(sliceTypeRaw Mod 5UI)              ' 0=P 1=B 2=I 3=SP 4=SI
            br.Ue()                                                            ' pic_parameter_set_id
            If sps.SeparateColourPlaneFlag Then br.U(2)                        ' colour_plane_id
            br.U(sps.Log2MaxFrameNumMinus4 + 4)                                ' frame_num
            Dim fieldPicFlag As UInteger = 0UI
            If Not sps.FrameMbsOnlyFlag Then
                fieldPicFlag = br.U(1)
                If fieldPicFlag <> 0UI Then br.U(1)                            ' bottom_field_flag
            End If
            If nalType = 5 Then br.Ue()                                        ' idr_pic_id
            If sps.PicOrderCntType = 0 Then
                br.U(sps.Log2MaxPicOrderCntLsbMinus4 + 4)                      ' pic_order_cnt_lsb
                If pps.BottomFieldPicOrderInFramePresentFlag AndAlso fieldPicFlag = 0UI Then
                    br.Se()                                                    ' delta_pic_order_cnt_bottom
                End If
            End If
            If sps.PicOrderCntType = 1 AndAlso Not sps.DeltaPicOrderAlwaysZeroFlag Then
                br.Se()                                                        ' delta_pic_order_cnt[0]
                If pps.BottomFieldPicOrderInFramePresentFlag AndAlso fieldPicFlag = 0UI Then
                    br.Se()                                                    ' delta_pic_order_cnt[1]
                End If
            End If
            If pps.RedundantPicCntPresentFlag Then br.Ue()                     ' redundant_pic_cnt

            ' for I/SI slices (2/4)
            ' P/B-specific fields aren't coded and neither is pred_weight_table (only applies to weighted P/B)
            ' typical case for the first slice (IDR = I-slice), so we can parse the rest cleanly
            ' reject non-I first slices to keep the parser bounded
            If sliceType <> 2 AndAlso sliceType <> 4 Then Return -1

            ' no ref_pic_list_modification for I/SI
            ' no pred_weight_table for I/SI

            ' dec_ref_pic_marking: coded iff nal_ref_idc != 0
            '
            ' - IDR (type 5): u(1) no_output_of_prior_pics + u(1) long_term_reference
            ' - non-IDR I-slices with nal_ref_idc != 0: adaptive_ref_pic_marking_mode_flag + optional MMCO loop
            ' first slice is usually IDR so we take that branch
            '
            Dim nalRefIdc As Integer = (bytes(nalStart - 1) >> 5) And &H3
            If nalRefIdc <> 0 Then
                If nalType = 5 Then
                    br.U(1)                                                    ' no_output_of_prior_pics_flag
                    br.U(1)                                                    ' long_term_reference_flag
                Else
                    ' non-IDR I-slice with references: MMCO loop, bail rather than parse
                    Return -1
                End If
            End If

            ' cabac_init_idc: coded iff entropy_coding_mode_flag && slice_type != I/SI
            ' we're on an I/SI slice so this is skipped
            br.Se()                                                            ' slice_qp_delta
            ' no slice_qs_delta (only SP/SI)
            If pps.DeblockingFilterControlPresent Then
                Return CInt(br.Ue() And &H3UI)                                 ' disable_deblocking_filter_idc
            End If
            Return 0
        Catch ex As Exception
            Return -1
        End Try
    End Function

    ' Minimal Exp-Golomb bit reader for slice-header parsing.
    Private Class SliceBitReader
        Private ReadOnly _data As Byte()
        Private _pos As Integer
        Public Sub New(data As Byte())
            _data = data
        End Sub
        Public Function U(n As Integer) As UInteger
            Dim v As UInteger = 0UI
            For k As Integer = 0 To n - 1
                If _pos >= _data.Length * 8 Then Throw New InvalidOperationException("Past end")
                Dim bit As Integer = (_data(_pos >> 3) >> (7 - (_pos And 7))) And 1
                v = (v << 1) Or CUInt(bit)
                _pos += 1
            Next
            Return v
        End Function
        Public Function Ue() As UInteger
            Dim zeros As Integer = 0
            While _pos < _data.Length * 8 AndAlso U(1) = 0UI
                zeros += 1
                If zeros > 31 Then Throw New InvalidOperationException("ue() codeword too long")
            End While
            If zeros = 0 Then Return 0UI
            Dim suffix As UInteger = U(zeros)
            Return (1UI << zeros) - 1UI + suffix
        End Function
        Public Function Se() As Integer
            Dim v As UInteger = Ue()
            If (v And 1UI) = 0UI Then
                Return -CInt(v >> 1)
            Else
                Return CInt((v + 1UI) >> 1)
            End If
        End Function
    End Class

    ' max_mean_bitrate byte per AVC level+resolution class
    ' observed:
    '  L3.1 720p CAVLC:     mmb = 5
    '  L4.1 1080p CABAC:    mmb = 11
    '  L4.1 720x480 (SD):   mmb = 1   (SD overrides level tier)
    ' !!! could be a bitrate class ID rather than a literal Mbps count !!!
    Private Function DefaultMaxMeanBitrateForAvcLevel(levelIdc As Integer, widthPx As Integer, heightPx As Integer) As Byte
        ' SD resolutions (≤ 720x576) use a lower mmb regardless of level
        If widthPx <= 720 AndAlso heightPx <= 576 Then Return 1
        If levelIdc >= 41 Then Return 11    ' L4.1 and up (1080p, high-bitrate 720p)
        If levelIdc >= 31 Then Return 5     ' L3.1 (720p CAVLC)
        If levelIdc >= 30 Then Return 3     ' L3.0 (SD-ish, extrapolated)
        Return 1                            ' L2.1 and below
    End Function

    ' detect frame-interval between successive SPS insertions in an Annex-B AVC stream
    '
    ' Sony PAMF encoder inserts one SPS at the start of each "seek block"
    ' ps2 marker payload N field (bytes 16-17) encodes exactly that block size
    '
    ' return 0 if we can't reliably determine the interval:
    '  - fewer than 3 SPS NALs (need at least 2 intervals for consistency check)
    '  - intervals are inconsistent (varying by more than 1 across samples)
    '  - result out of the [1, 64] range
    '
    ' callers should fall back to default (12) when this returns 0
    Private Function DetectAvcSpsInterval(bytes As Byte()) As Integer
        Dim frameCount As Integer = 0
        Dim spsFrames As New List(Of Integer)()
        Dim haveVclInCurrentFrame As Boolean = False
        Dim spsSeenInCurrentFrame As Boolean = False

        Dim i As Integer = 0
        While i < bytes.Length - 4
            Dim nalHdr As Integer = FindAnnexBNal(bytes, i)
            If nalHdr < 0 Then Exit While
            Dim nt As Integer = bytes(nalHdr) And &H1F
            Dim isVcl As Boolean = (nt = 1 OrElse nt = 5)
            Dim isFirstSlice As Boolean = isVcl AndAlso
                (nalHdr + 1 < bytes.Length) AndAlso ((bytes(nalHdr + 1) And &H80) <> 0)

            If nt = 9 AndAlso haveVclInCurrentFrame Then
                ' AUD after a VCL → previous frame complete, new frame starts here
                frameCount += 1
                haveVclInCurrentFrame = False
                spsSeenInCurrentFrame = False
            ElseIf isVcl AndAlso isFirstSlice AndAlso haveVclInCurrentFrame Then
                ' first-slice of new picture without an intervening AUD
                frameCount += 1
                haveVclInCurrentFrame = False
                spsSeenInCurrentFrame = False
            End If

            If nt = 7 AndAlso Not spsSeenInCurrentFrame Then     ' SPS
                spsFrames.Add(frameCount)
                spsSeenInCurrentFrame = True
            ElseIf isVcl Then
                haveVclInCurrentFrame = True
            End If
            i = nalHdr + 1

            ' bail early once we've gathered 8 SPS samples
            If spsFrames.Count >= 8 Then Exit While
        End While

        If spsFrames.Count < 3 Then Return 0

        ' check the last few intervals for consistency (skip the first, which is often anomalous for streams that start with a burst of RAPs at frames 0,3,5,6)
        Dim intervals As New List(Of Integer)()
        For k As Integer = 2 To spsFrames.Count - 1
            intervals.Add(spsFrames(k) - spsFrames(k - 1))
        Next
        If intervals.Count = 0 Then Return 0
        Dim first As Integer = intervals(0)
        For Each v In intervals
            If Math.Abs(v - first) > 1 Then Return 0    ' inconsistent
        Next
        If first < 1 OrElse first > 64 Then Return 0
        Return first
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

    Private Function AvcAuContainsIdr(au As Byte()) As Boolean
        ' RAP detection, AU is a random access point if it contains either:
        '   - IDR slice (NAL type 5), classic closed-GOP entry
        '   - SPS NAL (NAL type 7)
        Dim i As Integer = 0
        While i < au.Length - 4
            If au(i) = 0 AndAlso au(i + 1) = 0 Then
                Dim hdr As Integer = -1
                If au(i + 2) = 1 Then hdr = i + 3
                If au(i + 2) = 0 AndAlso (i + 3) < au.Length AndAlso au(i + 3) = 1 Then hdr = i + 4
                If hdr >= 0 AndAlso hdr < au.Length Then
                    Dim nt As Integer = au(hdr) And &H1F
                    If nt = 5 OrElse nt = 7 Then Return True
                End If
            End If
            i += 1
        End While
        Return False
    End Function

    Private Function AvcFrameRateCodeFromFps(fps As Double) As Byte
        If Math.Abs(fps - 24000.0 / 1001.0) < 0.01 Then Return 0
        If Math.Abs(fps - 24.0) < 0.01 Then Return 1
        If Math.Abs(fps - 25.0) < 0.01 Then Return 2
        If Math.Abs(fps - 30000.0 / 1001.0) < 0.01 Then Return 3
        If Math.Abs(fps - 30.0) < 0.01 Then Return 4
        If Math.Abs(fps - 50.0) < 0.01 Then Return 5
        If Math.Abs(fps - 60000.0 / 1001.0) < 0.01 Then Return 6
        Return 3   ' default 29.97
    End Function

    ' MPEG-2 Video: parse sequence_header, slice at picture boundaries.
    ' stream the input file for >2 GB M2V sources
    '   pass 1: scan sequentially, picture_start_code offset (Int64), sequence_header offset (for RAP marking), and vbv_delay
    '   pass 2: seek to each picture, read into per-picture Byte(), queue as AccessUnit

    Private Sub RegisterAndQueueM2v(mux As PamfMuxer, f As MuxInput, userSetInitialScr As Boolean)
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
        Dim fps As Double = seq.FrameRate
        If fps <= 0.0 Then fps = 30000.0 / 1001.0

        ' auto-detect InitialScr for M2V, unless override
        ' every M2V reference so faruses SCR0=30030
        If Not userSetInitialScr Then
            mux.PsMuxer.InitialScr = 30030L
        End If

        Dim frameRateCode As Byte = M2vFrameRateCodeFromFps(fps)
        Dim prog As Byte = CByte(If(seq.ProgressiveSequence, 1, 0))
        Dim colourPrim As Byte = If(seq.HasColourDescription, seq.ColourPrimaries, CByte(1))
        Dim transferCh As Byte = If(seq.HasColourDescription, seq.TransferCharacteristics, CByte(1))
        Dim matrixCoeff As Byte = If(seq.HasColourDescription, seq.MatrixCoefficients, CByte(1))
        Dim ps As PamfMuxStream = mux.AddM2vStream(
            profileLevel:=seq.ProfileAndLevel,
            progressive:=prog,
            frameRateCode:=frameRateCode,
            widthPx:=seq.WidthPixels, heightPx:=seq.HeightPixels,
            colourPrimaries:=colourPrim,
            transferChars:=transferCh,
            matrixCoeffs:=matrixCoeff)

        Dim tickPerFrame As Long = CLng(90000.0 / fps)
        Dim ptsBase As Long = 90000L

        ' 2) pass 1: scan the whole file for start codes
        Dim picOffs As New List(Of Long)()      ' file offset of each picture_start_code
        Dim picIsRap As New List(Of Boolean)()  ' whether a sequence_header preceded that picture
        Dim picVbvs As New List(Of UShort)()    ' vbv_delay parsed from picture_header
        Dim seqHdrPendingRap As Boolean = False
        Const ScanBufSize As Integer = 1024 * 1024

        Using fs As FileStream = File.OpenRead(f.Path)
            Const Overlap As Integer = 7
            Dim buf(ScanBufSize + Overlap - 1) As Byte
            Dim carry As Integer = 0
            Dim baseOff As Long = 0     ' file offset that buf[0] represents
            Dim atEof As Boolean = False

            Do
                Dim toRead As Integer = buf.Length - carry
                Dim n As Integer = fs.Read(buf, carry, toRead)
                If n = 0 Then atEof = True
                Dim totalInBuf As Integer = carry + n

                ' end of scan region: leave `Overlap` bytes at the tail unless this
                ' is the very last chunk, in which case scan all the way to the end
                Dim scanEnd As Integer = If(atEof, totalInBuf - 3, totalInBuf - Overlap)
                Dim i As Integer = 0
                While i < scanEnd
                    If buf(i) = 0 AndAlso buf(i + 1) = 0 AndAlso buf(i + 2) = 1 Then
                        Dim sc As Byte = buf(i + 3)
                        If sc = 0 Then
                            picOffs.Add(baseOff + i)
                            picIsRap.Add(seqHdrPendingRap)
                            seqHdrPendingRap = False
                            ' vbv_delay lives in the picture_header bytes at (start_code)+4..+7:
                            '   byte 5 (relative) = temporal_reference[1:0] | picture_coding_type[2:0] | vbv_delay[15:13]
                            '   byte 6 (relative) = vbv_delay[12:5]
                            '   byte 7 (relative) = vbv_delay[4:0] | ...
                            Dim vbv As UShort = 0
                            If i + 7 < totalInBuf Then
                                Dim b1 As Integer = buf(i + 5)
                                Dim b2 As Integer = buf(i + 6)
                                Dim b3 As Integer = buf(i + 7)
                                vbv = CUShort(((b1 And &H7) << 13) Or (b2 << 5) Or (b3 >> 3))
                            End If
                            picVbvs.Add(vbv)
                            i += 4
                            Continue While
                        ElseIf sc = &HB3 Then
                            seqHdrPendingRap = True
                            i += 4
                            Continue While
                        End If
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

        ' 3) pass 2: read each picture bytes and queue as an AU
        If picOffs.Count = 0 Then
            If fileLen > Int32.MaxValue Then
                Throw New InvalidDataException(
                    "No picture_start_code found and file exceeds 2 GiB - can't queue as one AU: " & f.Path)
            End If
            Dim au(CInt(fileLen) - 1) As Byte
            Using fs As FileStream = File.OpenRead(f.Path)
                Dim total As Integer = 0
                While total < au.Length
                    Dim n As Integer = fs.Read(au, total, au.Length - total)
                    If n <= 0 Then Exit While
                    total += n
                End While
            End Using
            mux.QueueAu(ps, New AccessUnit() With {
                .Data = au, .Pts = 90000L, .Dts = 90000L,
                .IsRandomAccessPoint = True
            })
            Return
        End If

        Using fs As FileStream = File.OpenRead(f.Path)
            For i As Integer = 0 To picOffs.Count - 1
                Dim startOff As Long = If(i = 0, 0L, picOffs(i))   ' first AU includes preamble
                Dim endOff As Long = If(i + 1 < picOffs.Count, picOffs(i + 1), fileLen)
                Dim sz As Long = endOff - startOff
                If sz <= 0 Then Continue For
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

                Dim pts As Long = ptsBase + CLng(i) * tickPerFrame
                Dim dts As Long = pts - tickPerFrame  ' one-frame reorder delay, see RegisterAndQueueAvc
                ' RAP if this AU starts with sequence_header (0x000001B3)
                Dim isRap As Boolean = picIsRap(i) OrElse (au.Length >= 4 AndAlso
                    au(0) = 0 AndAlso au(1) = 0 AndAlso au(2) = 1 AndAlso au(3) = &HB3)
                mux.QueueAu(ps, New AccessUnit() With {
                    .Data = au, .Pts = pts, .Dts = dts,
                    .IsRandomAccessPoint = isRap,
                    .VideoPictureIndex = i,
                    .VideoVbvDelay = picVbvs(i)
                })
            Next
        End Using
    End Sub

    Private Function FindStartCodes(bytes As Byte(), codeByte As Byte) As List(Of Integer)
        Dim r As New List(Of Integer)()
        Dim i As Integer = 0
        While i < bytes.Length - 4
            If bytes(i) = 0 AndAlso bytes(i + 1) = 0 AndAlso bytes(i + 2) = 1 _
            AndAlso bytes(i + 3) = codeByte Then
                r.Add(i)
                i += 4
            Else
                i += 1
            End If
        End While
        Return r
    End Function

    Private Function M2vFrameRateCodeFromFps(fps As Double) As Byte
        If Math.Abs(fps - 24000.0 / 1001.0) < 0.01 Then Return 1
        If Math.Abs(fps - 24.0) < 0.01 Then Return 2
        If Math.Abs(fps - 25.0) < 0.01 Then Return 3
        If Math.Abs(fps - 30000.0 / 1001.0) < 0.01 Then Return 4
        If Math.Abs(fps - 30.0) < 0.01 Then Return 5
        If Math.Abs(fps - 50.0) < 0.01 Then Return 6
        If Math.Abs(fps - 60000.0 / 1001.0) < 0.01 Then Return 7
        Return 4
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
        ' pts_step = 2048 * 90000 / sample_rate
        Dim ptsStep As Long = CLng(2048L * 90000L \ CLng(sr))
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
            mux.QueueAu(ps, New AccessUnit() With {
                .Data = pamfAu, .Pts = ptsBase + CLng(i) * ptsStep,
                .Dts = ptsBase + CLng(i) * ptsStep,
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

    ' AC-3: scan for 0x0B77 sync frames, one AU per frame

    Private Sub RegisterAndQueueAc3(mux As PamfMuxer, f As MuxInput)
        Dim data As Byte() = File.ReadAllBytes(f.Path)
        ' decode bsid and frame-size from the first sync frame to get channel info
        Dim firstSync As Integer = -1
        For i As Integer = 0 To data.Length - 8
            If data(i) = &HB AndAlso data(i + 1) = &H77 Then
                firstSync = i : Exit For
            End If
        Next
        If firstSync < 0 Then
            Throw New InvalidDataException("No AC-3 sync word found in " & f.Path)
        End If
        ' fscod (2 bits at byte+4 top), frmsizecod (6 bits same byte), bsid (5 bits at byte+5 top), bsmod (3 bits), acmod (3 bits at byte+6 top)
        Dim acmod As Integer = (data(firstSync + 6) >> 5) And 7
        Dim chTable As Integer() = {2, 1, 2, 3, 3, 4, 4, 5}    ' acmod 0..7 channel count
        Dim channels As Integer = chTable(acmod)
        ' LFE bit follows acmod-specific layout, conservatively add 1 if a likely 5.1 stream
        If channels = 5 Then channels = 6

        Dim ps As PamfMuxStream = mux.AddAc3Stream(numChannels:=CByte(channels))

        ' sample rate for AC-3: 48 kHz typical, ptsStep per AC-3 frame = 1536 samples
        Dim ptsStep As Long = CLng(1536L * 90000L \ 48000L)   ' = 2880 ticks
        Dim ptsBase As Long = 90000L

        ' slice at every sync word
        Dim auStarts As New List(Of Integer)()
        Dim p As Integer = firstSync
        While p < data.Length - 2
            If data(p) = &HB AndAlso data(p + 1) = &H77 Then
                auStarts.Add(p)
                p += 2
            Else
                p += 1
            End If
        End While
        For i As Integer = 0 To auStarts.Count - 1
            Dim s As Integer = auStarts(i)
            Dim e As Integer = If(i + 1 < auStarts.Count, auStarts(i + 1), data.Length)
            Dim au(e - s - 1) As Byte
            Array.Copy(data, s, au, 0, e - s)
            mux.QueueAu(ps, New AccessUnit() With {
                .Data = au, .Pts = ptsBase + CLng(i) * ptsStep,
                .Dts = ptsBase + CLng(i) * ptsStep,
                .IsRandomAccessPoint = False
            })
        Next
    End Sub

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
        Dim ptsStep As Long = CLng(samplesPerAu * 90000L \ CLng(sampleRate))
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

            mux.QueueAu(ps, New AccessUnit() With {
                .Data = au,
                .Pts = ptsBase + CLng(auIndex) * ptsStep,
                .Dts = ptsBase + CLng(auIndex) * ptsStep,
                .IsRandomAccessPoint = False
            })
            auIndex += 1
        End While
    End Sub

End Module