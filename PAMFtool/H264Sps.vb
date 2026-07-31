' H264Sps.vb - github.com/ravenDS/PAMFtool

' H.264 SPS (Sequence Parameter Set) parser
' decode the SPS including the VUI timing fields
' frame rate = time_scale / (2 * num_units_in_tick) -> ITU-T H.264 spec

Public Class H264SpsInfo
    Public Property WidthPixels As Integer
    Public Property HeightPixels As Integer
    Public Property ProfileIdc As Byte
    Public Property LevelIdc As Byte
    Public Property FrameMbsOnlyFlag As Boolean   ' True = progressive
    Public Property HasVuiTiming As Boolean
    Public Property NumUnitsInTick As UInteger
    Public Property TimeScale As UInteger
    Public Property FixedFrameRate As Boolean
    ' SPS frame-cropping rectangle (in chroma 4:2:0 units, as encoded in SPS)
    Public Property FrameCropLeftOffset As Integer
    Public Property FrameCropRightOffset As Integer
    Public Property FrameCropTopOffset As Integer
    Public Property FrameCropBottomOffset As Integer
    ' Aspect ratio: 0 = unspecified, 1 = 1:1, 14 = 4:3, etc (H.264 Table E-1)
    Public Property AspectRatioIdc As Byte
    Public Property SarWidth As Integer
    Public Property SarHeight As Integer
    ' VUI video_signal_type and colour_description (when present in SPS)
    Public Property HasVideoSignalInfo As Boolean
    Public Property VideoFormat As Byte    ' 0..5 (5 = unspecified)
    Public Property VideoFullRangeFlag As Byte
    Public Property ColourPrimaries As Byte    ' 1 = BT.709
    Public Property TransferCharacteristics As Byte  ' 1 = BT.709
    Public Property MatrixCoefficients As Byte    ' 1 = BT.709

    ' std_delay_bound is varying with peak video bit rate
    ' (1.00 s if peak <= 30 Mbps, 0.75 s if > 30 Mbps).
    ' peak comes from HRD when the encoder wrote a full VUI, otherwise callers fall back to conservative defaults
    Public Property HasHrdParameters As Boolean
    Public Property HrdPeakBitrateBps As Long    ' 0 if HRD absent

    ' fields the slice-header parser needs
    Public Property Log2MaxFrameNumMinus4 As Byte
    Public Property PicOrderCntType As Byte
    Public Property Log2MaxPicOrderCntLsbMinus4 As Byte
    Public Property DeltaPicOrderAlwaysZeroFlag As Boolean
    Public Property SeparateColourPlaneFlag As Boolean

    Public ReadOnly Property FrameRate As Double
        Get
            If Not HasVuiTiming OrElse NumUnitsInTick = 0 Then Return 0.0
            Return TimeScale / (2.0 * NumUnitsInTick)
        End Get
    End Property
End Class

Friend Module H264SpsParser

    ' Scan Annex-B byte range for the first NAL with nal_unit_type = 7
    ' Returns Nothing if no SPS is found or if parsing bails out
    Public Function ParseFirstSps(buf As Byte(), off As Integer, length As Integer) As H264SpsInfo
        Dim rbsp As Byte() = FindSpsRbsp(buf, off, length)
        If rbsp Is Nothing Then Return Nothing
        Dim unescaped() As Byte = StripEmulationPrevention(rbsp)
        Try
            Return ParseSpsRbsp(unescaped)
        Catch
            ' SPS contained something we couldn't parse, fall back to PAMF header
            Return Nothing
        End Try
    End Function

    ' Locate the SPS NAL payload, return bytes after 1-byte NAL header 
    Private Function FindSpsRbsp(buf As Byte(), off As Integer, length As Integer) As Byte()
        Dim [end] As Integer = Math.Min(buf.Length, off + length)
        Dim i As Integer = off
        While i < [end] - 4
            ' match 0x000001 or 0x00000001 then check nal_unit_type
            Dim nalHdrAt As Integer = -1
            If buf(i) = 0 AndAlso buf(i + 1) = 0 Then
                If buf(i + 2) = 1 Then
                    nalHdrAt = i + 3
                ElseIf buf(i + 2) = 0 AndAlso i + 3 < [end] AndAlso buf(i + 3) = 1 Then
                    nalHdrAt = i + 4
                End If
            End If

            If nalHdrAt >= 0 AndAlso nalHdrAt < [end] Then
                Dim nalType As Integer = buf(nalHdrAt) And &H1F
                If nalType = 7 Then
                    ' RBSP runs until the next start code (or EOB)
                    Dim s As Integer = nalHdrAt + 1
                    Dim e As Integer = s
                    While e < [end] - 2
                        If buf(e) = 0 AndAlso buf(e + 1) = 0 _
                        AndAlso (buf(e + 2) = 1 _
                                 OrElse (buf(e + 2) = 0 AndAlso e + 3 < [end] AndAlso buf(e + 3) = 1)) Then
                            Exit While
                        End If
                        e += 1
                    End While
                    Dim outBuf(e - s - 1) As Byte
                    Array.Copy(buf, s, outBuf, 0, e - s)
                    Return outBuf
                End If
                i = nalHdrAt + 1
            Else
                i += 1
            End If
        End While
        Return Nothing
    End Function

    ' strip 0x000000 / 0x000001 / 0x000002 / 0x000003 before bit-parsing
    Private Function StripEmulationPrevention(rbsp As Byte()) As Byte()
        Dim outList As New List(Of Byte)(rbsp.Length)
        Dim i As Integer = 0
        While i < rbsp.Length
            If i + 2 < rbsp.Length _
            AndAlso rbsp(i) = 0 AndAlso rbsp(i + 1) = 0 AndAlso rbsp(i + 2) = 3 Then
                outList.Add(0) : outList.Add(0)
                i += 3
            Else
                outList.Add(rbsp(i))
                i += 1
            End If
        End While
        Return outList.ToArray()
    End Function

    Private Function ParseSpsRbsp(rbsp As Byte()) As H264SpsInfo
        Dim br As New BitReader(rbsp)
        Dim sps As New H264SpsInfo()

        sps.ProfileIdc = CByte(br.U(8))
        br.U(8)                                              ' 6 constraint_set flags + 2 reserved
        sps.LevelIdc = CByte(br.U(8))
        br.Ue()                                              ' seq_parameter_set_id

        ' high-profile family has chroma_format_idc, bit-depth, scaling lists
        If sps.ProfileIdc = 100 OrElse sps.ProfileIdc = 110 _
        OrElse sps.ProfileIdc = 122 OrElse sps.ProfileIdc = 244 _
        OrElse sps.ProfileIdc = 44 OrElse sps.ProfileIdc = 83 _
        OrElse sps.ProfileIdc = 86 OrElse sps.ProfileIdc = 118 _
        OrElse sps.ProfileIdc = 128 OrElse sps.ProfileIdc = 138 _
        OrElse sps.ProfileIdc = 139 OrElse sps.ProfileIdc = 134 _
        OrElse sps.ProfileIdc = 135 Then
            Dim chromaFormatIdc As UInteger = br.Ue()
            If chromaFormatIdc = 3UI Then
                sps.SeparateColourPlaneFlag = (br.U(1) <> 0UI)             ' separate_colour_plane_flag
            End If
            br.Ue()                                           ' bit_depth_luma_minus8
            br.Ue()                                           ' bit_depth_chroma_minus8
            br.U(1)                                           ' qpprime_y_zero_transform_bypass_flag
            Dim seqScalingMatrixPresentFlag As UInteger = br.U(1)
            If seqScalingMatrixPresentFlag <> 0UI Then
                ' implement loop with se() deltas here
                Throw New NotSupportedException("SPS scaling lists present")
            End If
        End If

        sps.Log2MaxFrameNumMinus4 = CByte(br.Ue() And &HFUI)      ' log2_max_frame_num_minus4
        Dim picOrderCntType As UInteger = br.Ue()
        sps.PicOrderCntType = CByte(picOrderCntType And &HFFUI)
        If picOrderCntType = 0UI Then
            sps.Log2MaxPicOrderCntLsbMinus4 = CByte(br.Ue() And &HFUI)   ' log2_max_pic_order_cnt_lsb_minus4
        ElseIf picOrderCntType = 1UI Then
            sps.DeltaPicOrderAlwaysZeroFlag = (br.U(1) <> 0UI)   ' delta_pic_order_always_zero_flag
            br.Se()                                              ' offset_for_non_ref_pic
            br.Se()                                              ' offset_for_top_to_bottom_field
            Dim numRefInPicOrder As UInteger = br.Ue()           ' num_ref_frames_in_pic_order_cnt_cycle
            For i As Integer = 0 To CInt(numRefInPicOrder) - 1
                br.Se()                                          ' offset_for_ref_frame(i)
            Next
        End If

        br.Ue()                                               ' max_num_ref_frames
        br.U(1)                                               ' gaps_in_frame_num_value_allowed_flag
        Dim picWidthMbsMinus1 As UInteger = br.Ue()
        Dim picHeightMapUnitsMinus1 As UInteger = br.Ue()
        Dim frameMbsOnly As UInteger = br.U(1)
        sps.FrameMbsOnlyFlag = (frameMbsOnly <> 0UI)
        If frameMbsOnly = 0UI Then
            br.U(1)                                           ' mb_adaptive_frame_field_flag
        End If
        br.U(1)                                               ' direct_8x8_inference_flag

        Dim frameCropping As UInteger = br.U(1)
        Dim cropLeft As UInteger = 0
        Dim cropRight As UInteger = 0
        Dim cropTop As UInteger = 0
        Dim cropBottom As UInteger = 0
        If frameCropping <> 0UI Then
            cropLeft = br.Ue()
            cropRight = br.Ue()
            cropTop = br.Ue()
            cropBottom = br.Ue()
        End If
        sps.FrameCropLeftOffset = CInt(cropLeft)
        sps.FrameCropRightOffset = CInt(cropRight)
        sps.FrameCropTopOffset = CInt(cropTop)
        sps.FrameCropBottomOffset = CInt(cropBottom)

        ' pixel dimensions for crop, Chroma 4:2:0 assumed, for other chroma formats offsets scale differently
        Dim mbW As Integer = CInt(picWidthMbsMinus1 + 1UI) * 16
        Dim mbH As Integer = CInt(picHeightMapUnitsMinus1 + 1UI) * 16
        If Not sps.FrameMbsOnlyFlag Then mbH *= 2     ' field-coded: map units are pairs
        sps.WidthPixels = mbW - CInt((cropLeft + cropRight) * 2UI)
        sps.HeightPixels = mbH - CInt((cropTop + cropBottom) * 2UI _
                                      * (If(sps.FrameMbsOnlyFlag, 1UI, 2UI)))

        Dim vuiPresent As UInteger = br.U(1)
        If vuiPresent = 0UI Then Return sps     ' no VUI, no timing

        ' VUI parameters
        If br.U(1) <> 0UI Then                       ' aspect_ratio_info_present_flag
            Dim aspectIdc As UInteger = br.U(8)      ' aspect_ratio_idc
            sps.AspectRatioIdc = CByte(aspectIdc)
            If aspectIdc = 255UI Then                ' extended_SAR
                sps.SarWidth = CInt(br.U(16))
                sps.SarHeight = CInt(br.U(16))
            End If
        End If
        If br.U(1) <> 0UI Then br.U(1)               ' overscan_info_present + overscan_appropriate
        If br.U(1) <> 0UI Then                       ' video_signal_type_present_flag
            sps.HasVideoSignalInfo = True
            sps.VideoFormat = CByte(br.U(3))
            sps.VideoFullRangeFlag = CByte(br.U(1))
            If br.U(1) <> 0UI Then                   ' colour_description_present_flag
                sps.ColourPrimaries = CByte(br.U(8))
                sps.TransferCharacteristics = CByte(br.U(8))
                sps.MatrixCoefficients = CByte(br.U(8))
            End If
        End If
        If br.U(1) <> 0UI Then                       ' chroma_loc_info_present_flag
            br.Ue() : br.Ue()
        End If

        If br.U(1) <> 0UI Then                       ' timing_info_present_flag
            sps.HasVuiTiming = True
            sps.NumUnitsInTick = br.U(32)
            sps.TimeScale = br.U(32)
            sps.FixedFrameRate = (br.U(1) <> 0UI)
        End If

        ' HRD parameters
        ' we only need the peak bitrate so we skip cpb sizes and cbr flags
        Dim nalHrdPresent As Boolean = (br.U(1) <> 0UI)
        Dim peakBps As Long = 0
        If nalHrdPresent Then peakBps = Math.Max(peakBps, ParseHrdParameters(br))
        Dim vclHrdPresent As Boolean = (br.U(1) <> 0UI)
        If vclHrdPresent Then peakBps = Math.Max(peakBps, ParseHrdParameters(br))
        If nalHrdPresent OrElse vclHrdPresent Then
            sps.HasHrdParameters = True
            sps.HrdPeakBitrateBps = peakBps
        End If

        Return sps
    End Function

    ' parse hrd_parameters() and return the peak bitrate across all schedSelIdx
    ' peak formula:
    ' bitRate[i] = (bit_rate_value_minus1[i] + 1)* 2^(6 + bit_rate_scale) bps
    Private Function ParseHrdParameters(br As BitReader) As Long
        Dim cpbCntMinus1 As UInteger = br.Ue()
        Dim bitRateScale As UInteger = br.U(4)
        Dim cpbSizeScale As UInteger = br.U(4)
        Dim maxBps As Long = 0
        For i As UInteger = 0UI To cpbCntMinus1
            Dim brValue As UInteger = br.Ue()  ' bit_rate_value_minus1[i]
            Dim cpbSize As UInteger = br.Ue()  ' cpb_size_value_minus1[i] (unused)
            br.U(1)                            ' cbr_flag[i] (unused)
            Dim bps As Long = CLng(brValue + 1UI) << CInt(6 + bitRateScale)
            If bps > maxBps Then maxBps = bps
        Next
        br.U(5)     ' initial_cpb_removal_delay_length_minus1
        br.U(5)     ' cpb_removal_delay_length_minus1
        br.U(5)     ' dpb_output_delay_length_minus1
        br.U(5)     ' time_offset_length
        Return maxBps
    End Function

    Private Class BitReader
        Private ReadOnly _data As Byte()
        Private _pos As Integer
        Public Sub New(data As Byte())
            _data = data
            _pos = 0
        End Sub

        Public Function U(n As Integer) As UInteger
            Dim v As UInteger = 0UI
            For i As Integer = 0 To n - 1
                If _pos >= _data.Length * 8 Then
                    Throw New InvalidOperationException("BitReader past end of buffer")
                End If
                Dim bit As Integer = (_data(_pos >> 3) >> (7 - (_pos And 7))) And 1
                v = (v << 1) Or CUInt(bit)
                _pos += 1
            Next
            Return v
        End Function

        Public Function Ue() As UInteger
            Dim zeros As Integer = 0
            While U(1) = 0UI
                zeros += 1
                If zeros > 32 Then
                    Throw New InvalidOperationException("Exp-Golomb code too large")
                End If
            End While
            If zeros = 0 Then Return 0UI
            Return (CUInt(1) << zeros) - 1UI + U(zeros)
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

End Module

' PPS parsing (entropy_coding_mode + deblocking_filter_control)
Public Class H264PpsInfo
    Public Property CabacFlag As Boolean
    Public Property DeblockingFilterControlPresent As Boolean
    ' fields the slice-header parser needs
    Public Property BottomFieldPicOrderInFramePresentFlag As Boolean
    Public Property NumRefIdxL0DefaultActiveMinus1 As Byte
    Public Property NumRefIdxL1DefaultActiveMinus1 As Byte
    Public Property WeightedPredFlag As Boolean
    Public Property WeightedBipredIdc As Byte    ' 0..3
    Public Property RedundantPicCntPresentFlag As Boolean
End Class

Friend Module H264PpsParser

    Public Function ParseFirstPps(buf As Byte(), off As Integer, length As Integer) As H264PpsInfo
        Dim rbsp As Byte() = FindPpsRbsp(buf, off, length)
        If rbsp Is Nothing Then Return Nothing
        Dim unesc As Byte() = StripEpb(rbsp)
        Try
            Return ParsePpsRbsp(unesc)
        Catch
            Return Nothing
        End Try
    End Function

    Private Function FindPpsRbsp(buf As Byte(), off As Integer, length As Integer) As Byte()
        Dim [end] As Integer = Math.Min(buf.Length, off + length)
        Dim i As Integer = off
        While i < [end] - 4
            If buf(i) = 0 AndAlso buf(i + 1) = 0 Then
                Dim scLen As Integer = 0
                If buf(i + 2) = 1 Then
                    scLen = 3
                ElseIf buf(i + 2) = 0 AndAlso i + 3 < [end] AndAlso buf(i + 3) = 1 Then
                    scLen = 4
                End If
                If scLen > 0 Then
                    Dim nalStart As Integer = i + scLen
                    Dim nalType As Integer = buf(nalStart) And &H1F
                    If nalType = 8 Then    ' PPS
                        Dim nalEnd As Integer = nalStart + 1
                        While nalEnd < [end] - 2
                            If buf(nalEnd) = 0 AndAlso buf(nalEnd + 1) = 0 AndAlso
                               (buf(nalEnd + 2) = 1 OrElse
                                (buf(nalEnd + 2) = 0 AndAlso nalEnd + 3 < [end] AndAlso buf(nalEnd + 3) = 1)) Then
                                Exit While
                            End If
                            nalEnd += 1
                        End While
                        Dim rbsp(nalEnd - (nalStart + 1) - 1) As Byte
                        Array.Copy(buf, nalStart + 1, rbsp, 0, rbsp.Length)
                        Return rbsp
                    End If
                    i = nalStart
                    Continue While
                End If
            End If
            i += 1
        End While
        Return Nothing
    End Function

    Private Function StripEpb(rbsp As Byte()) As Byte()
        Dim outList As New List(Of Byte)(rbsp.Length)
        Dim i As Integer = 0
        While i < rbsp.Length
            If i + 2 < rbsp.Length AndAlso rbsp(i) = 0 AndAlso rbsp(i + 1) = 0 AndAlso rbsp(i + 2) = 3 Then
                outList.Add(0) : outList.Add(0)
                i += 3
            Else
                outList.Add(rbsp(i))
                i += 1
            End If
        End While
        Return outList.ToArray()
    End Function

    Private Function ParsePpsRbsp(rbsp As Byte()) As H264PpsInfo
        Dim br As New PpsBitReader(rbsp)
        Dim pps As New H264PpsInfo()

        ' ISO 14496-10 pic_parameter_set_rbsp():
        br.Ue()    ' pic_parameter_set_id
        br.Ue()    ' seq_parameter_set_id
        pps.CabacFlag = (br.U(1) <> 0UI)              ' entropy_coding_mode_flag
        pps.BottomFieldPicOrderInFramePresentFlag = (br.U(1) <> 0UI)

        Dim numSliceGroupsMinus1 As UInteger = br.Ue()
        If numSliceGroupsMinus1 > 0UI Then
            ' FMO
            Dim sliceGroupMapType As UInteger = br.Ue()
            Select Case CInt(sliceGroupMapType)
                Case 0
                    For iGroup As Integer = 0 To CInt(numSliceGroupsMinus1)
                        br.Ue()    ' run_length_minus1[iGroup]
                    Next
                Case 2
                    For iGroup As Integer = 0 To CInt(numSliceGroupsMinus1) - 1
                        br.Ue()    ' top_left[iGroup]
                        br.Ue()    ' bottom_right[iGroup]
                    Next
                Case 3, 4, 5
                    br.U(1)        ' slice_group_change_direction_flag
                    br.Ue()        ' slice_group_change_rate_minus1
                Case 6
                    Dim picSizeInMapUnitsMinus1 As UInteger = br.Ue()
                    ' slice_group_id(i) is u(v) where v = ceil(log2(numSG))
                    Dim numGroups As UInteger = numSliceGroupsMinus1 + 1UI
                    Dim v As Integer = 0
                    Dim tmp As UInteger = numGroups - 1UI
                    While tmp > 0UI
                        v += 1
                        tmp = tmp >> 1
                    End While
                    If v = 0 Then v = 1
                    For i As Integer = 0 To CInt(picSizeInMapUnitsMinus1)
                        br.U(v)    ' slice_group_id[i]
                    Next
            End Select
        End If

        pps.NumRefIdxL0DefaultActiveMinus1 = CByte(br.Ue() And &H3FUI)
        pps.NumRefIdxL1DefaultActiveMinus1 = CByte(br.Ue() And &H3FUI)
        pps.WeightedPredFlag = (br.U(1) <> 0UI)
        pps.WeightedBipredIdc = CByte(br.U(2) And 3UI)
        br.Se()    ' pic_init_qp_minus26
        br.Se()    ' pic_init_qs_minus26
        br.Se()    ' chroma_qp_index_offset

        pps.DeblockingFilterControlPresent = (br.U(1) <> 0UI)
        br.U(1)    ' constrained_intra_pred_flag
        pps.RedundantPicCntPresentFlag = (br.U(1) <> 0UI)
        Return pps
    End Function

    Private Class PpsBitReader
        Private ReadOnly _data As Byte()
        Private _pos As Integer
        Public Sub New(data As Byte())
            _data = data
            _pos = 0
        End Sub
        Public Function U(n As Integer) As UInteger
            Dim v As UInteger = 0UI
            For i As Integer = 0 To n - 1
                If _pos >= _data.Length * 8 Then Throw New InvalidOperationException("Past end")
                Dim bit As Integer = (_data(_pos >> 3) >> (7 - (_pos And 7))) And 1
                v = (v << 1) Or CUInt(bit)
                _pos += 1
            Next
            Return v
        End Function
        Public Function Ue() As UInteger
            Dim zeros As Integer = 0
            While U(1) = 0UI
                zeros += 1
                If zeros > 32 Then Throw New InvalidOperationException("Exp-Golomb too large")
            End While
            If zeros = 0 Then Return 0UI
            Return (CUInt(1) << zeros) - 1UI + U(zeros)
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

End Module