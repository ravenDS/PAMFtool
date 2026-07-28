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
                   forceNoDeblock As Boolean)
        If positional.Count < 2 Then
            Console.Error.WriteLine("Usage: PamfExtractor -mux <inputDir> <output.pamf> [-noep] [-deblock | -nodeblock]")
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

        Console.WriteLine("Muxing to: " & outPath)
        BuildPamfFromFiles(files, outPath, noEp, forceDeblock, forceNoDeblock)
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
                                   forceNoDeblock As Boolean)
        Dim mux As New PamfMuxer()
        mux.SkipEpTable = noEp

        ' for each input, parse codec metadata, register the stream, then slice into AUs and queue
        For Each f In files
            Select Case f.Kind
                Case "avc" : RegisterAndQueueAvc(mux, f, forceDeblock, forceNoDeblock)
                Case "mpeg2" : RegisterAndQueueM2v(mux, f)
                Case "atrac3plus" : RegisterAndQueueAt3p(mux, f)
                Case "ac3" : RegisterAndQueueAc3(mux, f)
                Case "lpcm" : RegisterAndQueueLpcm(mux, f)
            End Select
        Next

        Using outFs As New FileStream(outPath, FileMode.Create, FileAccess.Write)
            mux.WritePamfFile(outFs)
        End Using
    End Sub

    ' AVC: parse SPS, slice into per-frame AU, generate PTS/DTS

    Private Sub RegisterAndQueueAvc(mux As PamfMuxer, f As MuxInput,
                                    forceDeblock As Boolean,
                                    forceNoDeblock As Boolean)
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
        Dim deblock As Byte = If(pps IsNot Nothing AndAlso pps.DeblockingFilterControlPresent, CByte(1), CByte(0))
        If forceDeblock Then deblock = 1
        If forceNoDeblock Then deblock = 0
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

        ' slice into per-frame AU at VCL NAL boundaries (types 1, 5)
        ' each AU starts at a NALU, ends just before the next VCL NALU
        Dim auStarts As List(Of Integer) = FindAvcAuStarts(bytes)
        Dim tickPerFrame As Long = CLng(90000.0 / fps)
        Dim ptsBase As Long = 90000L
        For i As Integer = 0 To auStarts.Count - 1
            Dim s As Integer = auStarts(i)
            Dim e As Integer = If(i + 1 < auStarts.Count, auStarts(i + 1), bytes.Length)
            Dim au(e - s - 1) As Byte
            Array.Copy(bytes, s, au, 0, e - s)
            Dim pts As Long = ptsBase + CLng(i) * tickPerFrame
            mux.QueueAu(ps, New AccessUnit() With {
                .Data = au, .Pts = pts, .Dts = pts,
                .IsRandomAccessPoint = AvcAuContainsIdr(au)
            })
        Next
    End Sub

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

    Private Sub RegisterAndQueueM2v(mux As PamfMuxer, f As MuxInput)
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
                ' RAP if this AU starts with sequence_header (0x000001B3)
                Dim isRap As Boolean = picIsRap(i) OrElse (au.Length >= 4 AndAlso
                    au(0) = 0 AndAlso au(1) = 0 AndAlso au(2) = 1 AndAlso au(3) = &HB3)
                mux.QueueAu(ps, New AccessUnit() With {
                    .Data = au, .Pts = pts, .Dts = pts,
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

    ' ATRAC3plus: read .at3 RIFF, re-add 8-byte PAMF prefix per frame

    Private Sub RegisterAndQueueAt3p(mux As PamfMuxer, f As MuxInput)
        Dim chs As Integer = 0
        Dim sr As Integer = 0
        Dim frameSize As Integer = 0
        Dim frames As List(Of Byte()) = ReadAt3Frames(f.Path, chs, sr, frameSize)
        Dim ps As PamfMuxStream = mux.AddAtrac3plusStream(numChannels:=CByte(chs))

        ' one AU per frame, PTS spaced by samples-per-frame (2048) at sample rate converted to 90 kHz
        ' pts_step = 2048 * 90000 / sample_rate
        Dim ptsStep As Long = CLng(2048L * 90000L \ CLng(If(sr > 0, sr, 48000)))
        Dim ptsBase As Long = 90000L
        For i As Integer = 0 To frames.Count - 1
            Dim raw As Byte() = frames(i)
            ' build the 696-byte PAMF AU: 8-byte prefix + 688-byte frame
            Dim pamfAu(8 + raw.Length - 1) As Byte
            pamfAu(0) = &HF : pamfAu(1) = &HD0
            pamfAu(2) = &H48 : pamfAu(3) = &H55
            ' bytes 4..7 left as zero
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
                                   ByRef frameSize As Integer) As List(Of Byte())
        ' parse the RIFF/WAVE-style .at3 file
        ' (WAVE_FORMAT_EXTENSIBLE with Sony AT3+ GUID, blockAlign = 688).
        Dim data As Byte() = File.ReadAllBytes(path)
        If data.Length < 12 Then Throw New InvalidDataException("Tiny .at3 file: " & path)
        If data(0) <> &H52 OrElse data(1) <> &H49 _
        OrElse data(2) <> &H46 OrElse data(3) <> &H46 Then
            Throw New InvalidDataException("Not a RIFF file: " & path)
        End If
        Dim pos As Integer = 12
        Dim dataOff As Integer = -1
        Dim dataLen As Integer = 0
        While pos < data.Length - 8
            Dim chunkId As String = System.Text.Encoding.ASCII.GetString(data, pos, 4)
            Dim chunkSize As Integer = BitConverter.ToInt32(data, pos + 4)
            If chunkId = "fmt " Then
                ' offsets within fmt chunk: +2 channels, +4 sampleRate, +12 blockAlign
                channels = BitConverter.ToUInt16(data, pos + 8 + 2)
                sampleRate = BitConverter.ToInt32(data, pos + 8 + 4)
                frameSize = BitConverter.ToUInt16(data, pos + 8 + 12)
            ElseIf chunkId = "data" Then
                dataOff = pos + 8
                dataLen = chunkSize
                Exit While
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