' PamfExtractRunner.vb - github.com/ravenDS/PAMFtool

Imports System.IO
Imports System.Text
Imports PAMFtool.PamfMux

Friend Module PamfExtractRunner

    Public Sub Run(positional As List(Of String),
                   wantWav As Boolean, infoOnly As Boolean)
        Dim inPath As String = positional(0)
        Dim outDir As String = If(positional.Count >= 2,
                                  positional(1),
                                  Path.Combine(
                                      Path.GetDirectoryName(Path.GetFullPath(inPath)),
                                      Path.GetFileNameWithoutExtension(inPath) & "_extracted"))

        Dim pamf As New PamfFile()
        Console.WriteLine("Opening: " & inPath)
        pamf.Open(inPath)
        PrintHeaderInfo(pamf)

        If infoOnly Then Return

        Console.WriteLine()
        Console.WriteLine("Extracting to: " & outDir)
        pamf.ExtractAll(outDir)
        If wantWav Then
            ' implement audio to wav decode here
        End If
        Console.WriteLine("Done.")
    End Sub

    Public Sub PrintHeaderInfo(pamf As PamfFile)
        Console.WriteLine("  File size       : " & pamf.FileSize.ToString("N0") & " bytes")
        Console.WriteLine("  PAMF version    : " & pamf.Version)
        Console.WriteLine("  Header size     : 0x" & pamf.HeaderSize.ToString("X") &
                          " (" & pamf.HeaderSectors & " sector(s))")
        Console.WriteLine("  Stream offset   : 0x" & pamf.StreamOffset.ToString("X"))
        Console.WriteLine("  Stream size     : " & pamf.StreamSize.ToString("N0") &
                          " bytes (" & pamf.DataSectors & " packs)")
        Console.WriteLine("  psmf_marks      : offset=0x" & pamf.PsmfMarksOffset.ToString("X") &
                          " size=0x" & pamf.PsmfMarksSize.ToString("X"))
        Console.WriteLine("  unk chunk       : offset=0x" & pamf.UnkOffset.ToString("X") &
                          " size=0x" & pamf.UnkSize.ToString("X"))

        Console.WriteLine()
        Console.WriteLine("  PamfSequenceInfo @ 0x50:")
        Console.WriteLine("    size            : " & pamf.SeqInfoSize)
        Console.WriteLine("    start_pts       : 0x" & pamf.StartPts90.ToString("X") &
                          " (" & FormatPts(pamf.StartPts90) & ")")
        Console.WriteLine("    end_pts         : 0x" & pamf.EndPts90.ToString("X") &
                          " (" & FormatPts(pamf.EndPts90) & ")")
        Console.WriteLine("    duration        : " &
                          FormatPts(pamf.EndPts90 - pamf.StartPts90))
        Console.WriteLine("    mux_rate_bound  : 0x" & pamf.MuxRateBound.ToString("X") &
                          " (" & pamf.MuxRateBound.ToString("N0") & " units = " &
                          FormatBps(CLng(pamf.MuxRateBound) * 400L) & ")")
        Console.WriteLine("    std_delay_bound : 0x" & pamf.StdDelayBound.ToString("X") &
                          " (" & pamf.StdDelayBound.ToString("N0") & " ticks = " &
                          (pamf.StdDelayBound / 90000.0).ToString("0.000") & "s)")
        Console.WriteLine("    total_stream_num: " & pamf.TotalStreamNum)
        Console.WriteLine("    grouping_period#: " & pamf.GroupingPeriodNum)

        Console.WriteLine()
        Console.WriteLine("  PamfGroupingPeriod @ 0x70:")
        Console.WriteLine("    size            : " & pamf.GroupingPeriodSize)
        Console.WriteLine("    start_pts       : 0x" & pamf.GpStartPts90.ToString("X") &
                          " (" & FormatPts(pamf.GpStartPts90) & ")")
        Console.WriteLine("    end_pts         : 0x" & pamf.GpEndPts90.ToString("X") &
                          " (" & FormatPts(pamf.GpEndPts90) & ")")
        Console.WriteLine("    group_num       : " & pamf.GroupNum)

        Console.WriteLine()
        Console.WriteLine("  PamfGroup @ 0x82:")
        Console.WriteLine("    size            : " & pamf.GroupSize)
        Console.WriteLine("    stream_num      : " & pamf.Streams.Count)

        Console.WriteLine()
        If pamf.HasEpTable Then
            Console.WriteLine("  EP table        : PRESENT (" & pamf.TotalEpEntries &
                              " entries across " & pamf.Streams.Count & " streams)")
        Else
            Console.WriteLine("  EP table        : ABSENT (no stream has ep_num > 0)")
        End If

        Console.WriteLine()
        Console.WriteLine("  Streams         : " & pamf.Streams.Count)
        For Each s In pamf.Streams
            PrintStreamDetails(s)
        Next
    End Sub

    Private Sub PrintStreamDetails(s As PamfStreamInfo)
        Console.WriteLine()
        Console.WriteLine("  Stream #" & s.Index & " -- " & s.TypeName &
                          " (channel " & s.Channel & ")")
        Console.WriteLine("    coding_type    : 0x" & CByte(s.StreamType).ToString("X2"))
        Console.WriteLine("    stream_id      : 0x" & s.PesStreamId.ToString("X2") &
                          If(IsAudio(s.StreamType),
                             "  sub_stream_id: 0x" & s.SubStreamId.ToString("X2"),
                             ""))
        Console.WriteLine("    p_std_buffer   : 0x" & s.PStdBufferRaw.ToString("X4") &
                          "  (scale=" & s.PStdBufferScale &
                          ", size=" & s.PStdBufferSize & " -> " &
                          s.PStdBufferBytes.ToString("N0") & " bytes)")
        Console.WriteLine("    ep_offset      : 0x" & s.EpOffset.ToString("X") &
                          "   ep_num: " & s.EpNum)

        Select Case s.StreamType
            Case PamfStreamType.AVC : PrintAvcDetails(s)
            Case PamfStreamType.MPEG2Video : PrintM2vDetails(s)
            Case PamfStreamType.ATRAC3plus, PamfStreamType.AC3 : PrintAudioDetails(s)
            Case PamfStreamType.LPCM : PrintAudioDetails(s) : PrintLpcmExtras(s)
        End Select

        If s.CodecInfo IsNot Nothing Then
            Console.WriteLine("    codec_info raw : " & HexBytes(s.CodecInfo, 0, 16))
            Console.WriteLine("                     " & HexBytes(s.CodecInfo, 16, 16))
        End If
    End Sub

    Private Sub PrintAvcDetails(s As PamfStreamInfo)
        Console.WriteLine("    profile/level  : " & DescribeAvcProfile(s.ProfileIdc) &
                          " (0x" & s.ProfileIdc.ToString("X2") & ") / Level " &
                          DescribeAvcLevel(s.LevelIdc) & " (0x" & s.LevelIdc.ToString("X2") & ")")
        Console.WriteLine("    resolution     : " & s.Width & "x" & s.Height &
                          "  (MB-grid " & (s.Width \ 16) & "x" & (s.Height \ 16) & ")")
        Console.WriteLine("    frame_rate     : code " & s.FrameRateInfo & " (" &
                          DescribeAvcFrameRate(s.FrameRateInfo) & ")")
        Console.WriteLine("    frame_mbs_only : " & s.FrameMbsOnlyFlag &
                          " (" & If(s.FrameMbsOnlyFlag = 1, "progressive", "interlaced") & ")")
        Console.WriteLine("    aspect_ratio   : IDC " & s.AspectRatioIdc &
                          " (" & DescribeAspectRatio(s.AspectRatioIdc) & ")" &
                          If(s.AspectRatioIdc = 255, "  SAR " & s.SarWidth & ":" & s.SarHeight, ""))
        Console.WriteLine("    crop offsets   : L=" & s.FrameCropLeftOffset &
                          " R=" & s.FrameCropRightOffset &
                          " T=" & s.FrameCropTopOffset &
                          " B=" & s.FrameCropBottomOffset)
        Console.WriteLine("    video_signal   : flag=" & s.VideoSignalInfoFlag &
                          "  format=" & s.VideoFormat & " (" &
                          DescribeVideoFormat(s.VideoFormat) & ")  full_range=" &
                          s.VideoFullRangeFlag)
        Console.WriteLine("    colour         : primaries=" & s.ColourPrimaries &
                          " (" & DescribeColour(s.ColourPrimaries) & ")  transfer=" &
                          s.TransferCharacteristics &
                          " (" & DescribeColour(s.TransferCharacteristics) & ")  matrix=" &
                          s.MatrixCoefficients &
                          " (" & DescribeColour(s.MatrixCoefficients) & ")")
        Console.WriteLine("    entropy_coding : " & s.CabacFlag &
                          " (" & If(s.CabacFlag = 1, "CABAC", "CAVLC") & ")")
        Console.WriteLine("    deblock_flag   : " & s.DeblockingFilterFlag &
                          " (deblocking_filter_control_present_flag)")
        Console.WriteLine("    min_slice_idc  : " & s.MinNumSlicePerPictureIdc)
        Console.WriteLine("    nfw_idc        : " & s.NfwIdc)
        Console.WriteLine("    max_mean_bitrate: " & s.MaxMeanBitrate)
    End Sub

    Private Sub PrintM2vDetails(s As PamfStreamInfo)
        Console.WriteLine("    profile_level  : 0x" & s.ProfileIdc.ToString("X2") &
                          " (" & DescribeM2vProfileLevel(s.ProfileIdc) & ")")
        Console.WriteLine("    resolution     : " & s.Width & "x" & s.Height &
                          "  (explicit " & s.HorizontalSizeValue & "x" & s.VerticalSizeValue & ")")
        Console.WriteLine("    frame_rate     : code " & s.FrameRateInfo & " (" &
                          DescribeM2vFrameRate(s.FrameRateInfo) & ")")
        Console.WriteLine("    progressive_seq: " & s.ProgressiveSequence)
        Console.WriteLine("    aspect_ratio   : IDC " & s.AspectRatioIdc &
                          If(s.AspectRatioIdc = 255, "  SAR " & s.SarWidth & ":" & s.SarHeight, ""))
        Console.WriteLine("    video_signal   : flag=" & s.VideoSignalInfoFlag &
                          "  format=" & s.VideoFormat & "  full_range=" & s.VideoFullRangeFlag)
        Console.WriteLine("    colour         : primaries=" & s.ColourPrimaries &
                          " transfer=" & s.TransferCharacteristics &
                          " matrix=" & s.MatrixCoefficients)
    End Sub

    Private Sub PrintAudioDetails(s As PamfStreamInfo)
        Console.WriteLine("    channels       : " & s.NumChannels)
        Console.WriteLine("    sample_rate    : " & s.SampleRate & " Hz")
    End Sub

    Private Sub PrintLpcmExtras(s As PamfStreamInfo)
        Console.WriteLine("    bits_per_sample: " & s.BitsPerSample)
    End Sub

    Private Function FormatPts(ticks90 As Long) As String
        If ticks90 < 0L Then Return "<negative>"
        Dim seconds As Double = ticks90 / 90000.0
        Return seconds.ToString("0.0000") & "s"
    End Function

    Private Function FormatBps(bps As Long) As String
        If bps >= 1_000_000L Then
            Return (bps / 1_000_000.0).ToString("0.00") & " Mbps"
        ElseIf bps >= 1_000L Then
            Return (bps / 1_000.0).ToString("0.0") & " Kbps"
        End If
        Return bps & " bps"
    End Function

    Private Function HexBytes(b As Byte(), off As Integer, count As Integer) As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To count - 1
            If i > 0 Then sb.Append(" ")
            sb.Append(b(off + i).ToString("X2"))
        Next
        Return sb.ToString()
    End Function

    Private Function IsAudio(t As PamfStreamType) As Boolean
        Return t = PamfStreamType.ATRAC3plus _
            OrElse t = PamfStreamType.AC3 _
            OrElse t = PamfStreamType.LPCM _
            OrElse t = PamfStreamType.UserData
    End Function

    Private Function DescribeAvcProfile(idc As Byte) As String
        Select Case CInt(idc)
            Case 66 : Return "Baseline"
            Case 77 : Return "Main"
            Case 88 : Return "Extended"
            Case 100 : Return "High"
            Case 110 : Return "High 10"
            Case 122 : Return "High 4:2:2"
            Case 244 : Return "High 4:4:4"
            Case Else : Return "profile " & idc
        End Select
    End Function

    Private Function DescribeAvcLevel(idc As Byte) As String
        ' Level 3.1 = 0x1F = 31, etc
        ' Encoded as level * 10
        Dim major As Integer = idc \ 10
        Dim minor As Integer = idc Mod 10
        Return major & "." & minor
    End Function

    Private Function DescribeAvcFrameRate(code As Byte) As String
        Select Case CInt(code)
            Case 0 : Return "24000/1001 (~23.976)"
            Case 1 : Return "24"
            Case 2 : Return "25"
            Case 3 : Return "30000/1001 (~29.97)"
            Case 4 : Return "30"
            Case 5 : Return "50"
            Case 6 : Return "60000/1001 (~59.94)"
            Case Else : Return "unknown"
        End Select
    End Function

    Private Function DescribeM2vFrameRate(code As Byte) As String
        ' MPEG-2 frame_rate_code table (1-8), PAMF stores raw value
        Select Case CInt(code)
            Case 1 : Return "24000/1001 (~23.976)"
            Case 2 : Return "24"
            Case 3 : Return "25"
            Case 4 : Return "30000/1001 (~29.97)"
            Case 5 : Return "30"
            Case 6 : Return "50"
            Case 7 : Return "60000/1001 (~59.94)"
            Case 8 : Return "60"
            Case Else : Return "unknown"
        End Select
    End Function

    Private Function DescribeM2vProfileLevel(b As Byte) As String
        Select Case CInt(b)
            Case &H44 : Return "Main Profile @ High Level"
            Case &H48 : Return "Main Profile @ Main Level"
            Case &H4A : Return "Main Profile @ Low Level"
            Case Else : Return "code 0x" & b.ToString("X2")
        End Select
    End Function

    Private Function DescribeAspectRatio(idc As Byte) As String
        Select Case CInt(idc)
            Case 0 : Return "unspecified"
            Case 1 : Return "1:1 (square)"
            Case 2 : Return "12:11"
            Case 3 : Return "10:11"
            Case 4 : Return "16:11"
            Case 5 : Return "40:33"
            Case 6 : Return "24:11"
            Case 7 : Return "20:11"
            Case 8 : Return "32:11"
            Case 9 : Return "80:33"
            Case 10 : Return "18:11"
            Case 11 : Return "15:11"
            Case 12 : Return "64:33"
            Case 13 : Return "160:99"
            Case 14 : Return "4:3"
            Case 15 : Return "3:2"
            Case 16 : Return "2:1"
            Case 255 : Return "extended SAR"
            Case Else : Return "reserved"
        End Select
    End Function

    Private Function DescribeVideoFormat(v As Byte) As String
        Select Case CInt(v)
            Case 0 : Return "Component"
            Case 1 : Return "PAL"
            Case 2 : Return "NTSC"
            Case 3 : Return "SECAM"
            Case 4 : Return "MAC"
            Case 5 : Return "unspecified"
            Case Else : Return "reserved"
        End Select
    End Function

    Private Function DescribeColour(v As Byte) As String
        Select Case CInt(v)
            Case 1 : Return "BT.709"
            Case 2 : Return "unspecified"
            Case 4 : Return "BT.470M"
            Case 5 : Return "BT.470BG"
            Case 6 : Return "SMPTE 170M / BT.601"
            Case 7 : Return "SMPTE 240M"
            Case 8 : Return "linear / YCgCo"
            Case 9 : Return "BT.2020"
            Case Else : Return "code " & v
        End Select
    End Function

End Module