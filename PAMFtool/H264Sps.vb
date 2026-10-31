' H264Sps.vb - github.com/ravenDS/PAMFtool

' H.264 SPS/PPS parsers (refs: ITU-T H.264 7.3.2.1.1 / 7.3.2.2, VUI per Annex E)
' frame rate = time_scale / (2 * num_units_in_tick)

Public Class H264SpsInfo
    Public Property SpsId As Integer
    Public Property WidthPixels As Integer
    Public Property HeightPixels As Integer
    Public Property ProfileIdc As Byte
    Public Property LevelIdc As Byte
    Public Property ConstraintFlags As Byte
    Public Property FrameMbsOnlyFlag As Boolean   ' True = progressive
    Public Property MbAdaptiveFrameFieldFlag As Boolean
    Public Property HasVuiTiming As Boolean
    Public Property NumUnitsInTick As UInteger
    Public Property TimeScale As UInteger
    Public Property FixedFrameRate As Boolean
    ' SPS frame-cropping rectangle (in crop units, as encoded in SPS)
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

    ' HRD (when the encoder wrote one): peak bit rate and the coded picture buffer size
    ' std_delay_bound follows peak (1.00 s if peak <= 30 Mbps, 0.75 s above)
    ' P-STD buffer follows CPB size: ceil((cpb_size / 8 + 40960) / 1024) KB
    Public Property HasHrdParameters As Boolean
    Public Property HrdPeakBitrateBps As Long    ' 0 if HRD absent
    Public Property NalHrdCpbSizeBits As Long    ' 0 if NAL HRD absent (largest SchedSelIdx)
    Public Property VclHrdCpbSizeBits As Long    ' 0 if VCL HRD absent
    Public Property PicStructPresentFlag As Boolean
    Public Property HasBitstreamRestriction As Boolean
    Public Property MaxNumReorderFrames As Integer = -1
    Public Property MaxDecFrameBuffering As Integer = -1
    Public Property VuiParseFailed As Boolean

    ' fields the slice-header/POC code needs
    Public Property ChromaFormatIdc As Integer = 1
    Public Property SeparateColourPlaneFlag As Boolean
    Public Property BitDepthLumaMinus8 As Integer
    Public Property BitDepthChromaMinus8 As Integer
    Public Property Log2MaxFrameNumMinus4 As Byte
    Public Property PicOrderCntType As Byte
    Public Property Log2MaxPicOrderCntLsbMinus4 As Byte
    Public Property DeltaPicOrderAlwaysZeroFlag As Boolean
    Public Property OffsetForNonRefPic As Integer
    Public Property OffsetForTopToBottomField As Integer
    Public Property OffsetForRefFrame As Integer() = Array.Empty(Of Integer)()
    Public Property MaxNumRefFrames As Integer

    Public ReadOnly Property ChromaArrayType As Integer
        Get
            Return If(SeparateColourPlaneFlag, 0, ChromaFormatIdc)
        End Get
    End Property

    Public ReadOnly Property FrameRate As Double
        Get
            If Not HasVuiTiming OrElse NumUnitsInTick = 0 Then Return 0.0
            Return TimeScale / (2.0 * NumUnitsInTick)
        End Get
    End Property
End Class

Friend Module H264SpsParser

    ' high-profile family: these profile_idc values carry chroma_format_idc, bit depths and scaling lists
    Public Function IsHighProfileFamily(profileIdc As Integer) As Boolean
        Select Case profileIdc
            Case 100, 110, 122, 244, 44, 83, 86, 118, 128, 138, 139, 134, 135
                Return True
        End Select
        Return False
    End Function

    ' Scan Annex-B byte range for the first NAL with nal_unit_type = 7
    ' Returns Nothing if no SPS is found or if parsing bails out
    Public Function ParseFirstSps(buf As Byte(), off As Integer, length As Integer) As H264SpsInfo
        Dim [end] As Integer = Math.Min(buf.Length, off + length)
        Dim nal As NalSpan
        Dim p As Integer = off
        While NextNal(buf, p, [end], nal)
            If (buf(nal.Header) And &H1F) = 7 Then
                Return ParseSpsNal(buf, nal.Header, nal.PayloadEnd)
            End If
            p = nal.PayloadEnd
        End While
        Return Nothing
    End Function

    ' nalHeader = index of the NAL header byte, nalEnd = end of the NAL payload
    Public Function ParseSpsNal(buf As Byte(), nalHeader As Integer, nalEnd As Integer) As H264SpsInfo
        Try
            Return ParseSpsRbsp(Unescape(buf, nalHeader + 1, nalEnd))
        Catch
            Return Nothing
        End Try
    End Function

    Private Sub SkipScalingList(br As RbspReader, size As Integer)
        Dim lastScale As Integer = 8
        Dim nextScale As Integer = 8
        For j As Integer = 0 To size - 1
            If nextScale <> 0 Then
                Dim delta As Integer = br.Se()
                nextScale = (lastScale + delta + 256) Mod 256
            End If
            If nextScale <> 0 Then lastScale = nextScale
        Next
    End Sub

    Private Function ParseSpsRbsp(rbsp As Byte()) As H264SpsInfo
        Dim br As New RbspReader(rbsp)
        Dim sps As New H264SpsInfo()

        sps.ProfileIdc = CByte(br.U(8))
        sps.ConstraintFlags = CByte(br.U(8))                 ' 6 constraint_set flags + 2 reserved
        sps.LevelIdc = CByte(br.U(8))
        sps.SpsId = CInt(br.Ue())

        If IsHighProfileFamily(sps.ProfileIdc) Then
            sps.ChromaFormatIdc = CInt(br.Ue())
            If sps.ChromaFormatIdc = 3 Then
                sps.SeparateColourPlaneFlag = br.Flag()
            End If
            sps.BitDepthLumaMinus8 = CInt(br.Ue())
            sps.BitDepthChromaMinus8 = CInt(br.Ue())
            br.U(1)                                           ' qpprime_y_zero_transform_bypass_flag
            If br.Flag() Then                                 ' seq_scaling_matrix_present_flag
                Dim lists As Integer = If(sps.ChromaFormatIdc <> 3, 8, 12)
                For i As Integer = 0 To lists - 1
                    If br.Flag() Then                         ' seq_scaling_list_present_flag[i]
                        SkipScalingList(br, If(i < 6, 16, 64))
                    End If
                Next
            End If
        End If

        sps.Log2MaxFrameNumMinus4 = CByte(br.Ue() And &HFUI)
        Dim picOrderCntType As UInteger = br.Ue()
        sps.PicOrderCntType = CByte(picOrderCntType And &HFFUI)
        If picOrderCntType = 0UI Then
            sps.Log2MaxPicOrderCntLsbMinus4 = CByte(br.Ue() And &HFUI)
        ElseIf picOrderCntType = 1UI Then
            sps.DeltaPicOrderAlwaysZeroFlag = br.Flag()
            sps.OffsetForNonRefPic = br.Se()
            sps.OffsetForTopToBottomField = br.Se()
            Dim n As Integer = CInt(br.Ue())
            If n > 255 Then Throw New InvalidOperationException("num_ref_frames_in_pic_order_cnt_cycle out of range")
            Dim offs(n - 1) As Integer
            For i As Integer = 0 To n - 1
                offs(i) = br.Se()
            Next
            sps.OffsetForRefFrame = offs
        End If

        sps.MaxNumRefFrames = CInt(br.Ue())
        br.U(1)                                               ' gaps_in_frame_num_value_allowed_flag
        Dim picWidthMbsMinus1 As UInteger = br.Ue()
        Dim picHeightMapUnitsMinus1 As UInteger = br.Ue()
        sps.FrameMbsOnlyFlag = br.Flag()
        If Not sps.FrameMbsOnlyFlag Then
            sps.MbAdaptiveFrameFieldFlag = br.Flag()
        End If
        br.U(1)                                               ' direct_8x8_inference_flag

        Dim cropLeft As UInteger = 0
        Dim cropRight As UInteger = 0
        Dim cropTop As UInteger = 0
        Dim cropBottom As UInteger = 0
        If br.Flag() Then                                     ' frame_cropping_flag
            cropLeft = br.Ue()
            cropRight = br.Ue()
            cropTop = br.Ue()
            cropBottom = br.Ue()
        End If
        sps.FrameCropLeftOffset = CInt(cropLeft)
        sps.FrameCropRightOffset = CInt(cropRight)
        sps.FrameCropTopOffset = CInt(cropTop)
        sps.FrameCropBottomOffset = CInt(cropBottom)

        ' crop units: CropUnitX = SubWidthC, CropUnitY = SubHeightC * (2 - frame_mbs_only_flag)
        ' (1 and 2 - frame_mbs_only_flag when ChromaArrayType = 0)
        Dim subW As Integer = 1
        Dim subH As Integer = 1
        Select Case sps.ChromaArrayType
            Case 1 : subW = 2 : subH = 2
            Case 2 : subW = 2 : subH = 1
        End Select
        Dim cropUnitY As Integer = subH * If(sps.FrameMbsOnlyFlag, 1, 2)
        Dim mbW As Integer = CInt(picWidthMbsMinus1 + 1UI) * 16
        Dim mbH As Integer = CInt(picHeightMapUnitsMinus1 + 1UI) * 16
        If Not sps.FrameMbsOnlyFlag Then mbH *= 2     ' field-coded: map units are MB pairs
        sps.WidthPixels = mbW - CInt(cropLeft + cropRight) * subW
        sps.HeightPixels = mbH - CInt(cropTop + cropBottom) * cropUnitY

        If Not br.Flag() Then Return sps                      ' vui_parameters_present_flag

        ' a damaged / truncated VUI must not throw away the SPS fields above
        Try
            ParseVui(br, sps)
        Catch
            sps.VuiParseFailed = True
        End Try
        Return sps
    End Function

    Private Sub ParseVui(br As RbspReader, sps As H264SpsInfo)
        If br.Flag() Then                            ' aspect_ratio_info_present_flag
            Dim aspectIdc As UInteger = br.U(8)
            sps.AspectRatioIdc = CByte(aspectIdc)
            If aspectIdc = 255UI Then                ' Extended_SAR
                sps.SarWidth = CInt(br.U(16))
                sps.SarHeight = CInt(br.U(16))
            End If
        End If
        If br.Flag() Then br.U(1)                    ' overscan_info_present_flag + overscan_appropriate_flag
        If br.Flag() Then                            ' video_signal_type_present_flag
            sps.HasVideoSignalInfo = True
            sps.VideoFormat = CByte(br.U(3))
            sps.VideoFullRangeFlag = CByte(br.U(1))
            If br.Flag() Then                        ' colour_description_present_flag
                sps.ColourPrimaries = CByte(br.U(8))
                sps.TransferCharacteristics = CByte(br.U(8))
                sps.MatrixCoefficients = CByte(br.U(8))
            End If
        End If
        If br.Flag() Then                            ' chroma_loc_info_present_flag
            br.Ue() : br.Ue()
        End If
        If br.Flag() Then                            ' timing_info_present_flag
            sps.HasVuiTiming = True
            sps.NumUnitsInTick = br.U(32)
            sps.TimeScale = br.U(32)
            sps.FixedFrameRate = br.Flag()
        End If

        Dim nalHrd As Boolean = br.Flag()
        Dim peakBps As Long = 0
        If nalHrd Then
            Dim h As HrdInfo = ParseHrdParameters(br)
            peakBps = Math.Max(peakBps, h.MaxBitRate)
            sps.NalHrdCpbSizeBits = h.MaxCpbSizeBits
        End If
        Dim vclHrd As Boolean = br.Flag()
        If vclHrd Then
            Dim h As HrdInfo = ParseHrdParameters(br)
            peakBps = Math.Max(peakBps, h.MaxBitRate)
            sps.VclHrdCpbSizeBits = h.MaxCpbSizeBits
        End If
        If nalHrd OrElse vclHrd Then
            sps.HasHrdParameters = True
            sps.HrdPeakBitrateBps = peakBps
            br.U(1)                                  ' low_delay_hrd_flag
        End If
        sps.PicStructPresentFlag = br.Flag()
        If br.Flag() Then                            ' bitstream_restriction_flag
            sps.HasBitstreamRestriction = True
            br.U(1)                                  ' motion_vectors_over_pic_boundaries_flag
            br.Ue()                                  ' max_bytes_per_pic_denom
            br.Ue()                                  ' max_bits_per_mb_denom
            br.Ue()                                  ' log2_max_mv_length_horizontal
            br.Ue()                                  ' log2_max_mv_length_vertical
            sps.MaxNumReorderFrames = CInt(br.Ue())
            sps.MaxDecFrameBuffering = CInt(br.Ue())
        End If
    End Sub

    Private Structure HrdInfo
        Public MaxBitRate As Long
        Public MaxCpbSizeBits As Long
    End Structure

    ' hrd_parameters() (E.1.2):
    '   bitRate[i] = (bit_rate_value_minus1[i] + 1) * 2^(6 + bit_rate_scale)       bit/s
    '   cpbSize[i] = (cpb_size_value_minus1[i] + 1) * 2^(4 + cpb_size_scale)       bits
    Private Function ParseHrdParameters(br As RbspReader) As HrdInfo
        Dim cpbCntMinus1 As UInteger = br.Ue()
        If cpbCntMinus1 > 31UI Then Throw New InvalidOperationException("cpb_cnt_minus1 out of range")
        Dim bitRateScale As Integer = CInt(br.U(4))
        Dim cpbSizeScale As Integer = CInt(br.U(4))
        Dim r As New HrdInfo()
        For i As UInteger = 0UI To cpbCntMinus1
            Dim brValue As Long = CLng(br.Ue()) + 1L
            Dim cpbValue As Long = CLng(br.Ue()) + 1L
            br.U(1)                            ' cbr_flag[i]
            r.MaxBitRate = Math.Max(r.MaxBitRate, brValue << (6 + bitRateScale))
            r.MaxCpbSizeBits = Math.Max(r.MaxCpbSizeBits, cpbValue << (4 + cpbSizeScale))
        Next
        br.U(5)     ' initial_cpb_removal_delay_length_minus1
        br.U(5)     ' cpb_removal_delay_length_minus1
        br.U(5)     ' dpb_output_delay_length_minus1
        br.U(5)     ' time_offset_length
        Return r
    End Function

End Module

' PPS parsing (entropy_coding_mode + deblocking_filter_control + everything the slice header needs)
Public Class H264PpsInfo
    Public Property PpsId As Integer
    Public Property SpsId As Integer
    Public Property CabacFlag As Boolean
    Public Property DeblockingFilterControlPresent As Boolean
    ' fields the slice-header parser needs
    Public Property BottomFieldPicOrderInFramePresentFlag As Boolean
    Public Property NumSliceGroupsMinus1 As Integer
    Public Property SliceGroupMapType As Integer
    Public Property NumRefIdxL0DefaultActiveMinus1 As Byte
    Public Property NumRefIdxL1DefaultActiveMinus1 As Byte
    Public Property WeightedPredFlag As Boolean
    Public Property WeightedBipredIdc As Byte    ' 0..3
    Public Property RedundantPicCntPresentFlag As Boolean
End Class

Friend Module H264PpsParser

    Public Function ParseFirstPps(buf As Byte(), off As Integer, length As Integer) As H264PpsInfo
        Dim [end] As Integer = Math.Min(buf.Length, off + length)
        Dim nal As NalSpan
        Dim p As Integer = off
        While NextNal(buf, p, [end], nal)
            If (buf(nal.Header) And &H1F) = 8 Then
                Return ParsePpsNal(buf, nal.Header, nal.PayloadEnd)
            End If
            p = nal.PayloadEnd
        End While
        Return Nothing
    End Function

    Public Function ParsePpsNal(buf As Byte(), nalHeader As Integer, nalEnd As Integer) As H264PpsInfo
        Try
            Return ParsePpsRbsp(Unescape(buf, nalHeader + 1, nalEnd))
        Catch
            Return Nothing
        End Try
    End Function

    Private Function ParsePpsRbsp(rbsp As Byte()) As H264PpsInfo
        Dim br As New RbspReader(rbsp)
        Dim pps As New H264PpsInfo()

        ' ISO 14496-10 pic_parameter_set_rbsp():
        pps.PpsId = CInt(br.Ue())
        pps.SpsId = CInt(br.Ue())
        pps.CabacFlag = br.Flag()                         ' entropy_coding_mode_flag
        pps.BottomFieldPicOrderInFramePresentFlag = br.Flag()

        pps.NumSliceGroupsMinus1 = CInt(br.Ue())
        If pps.NumSliceGroupsMinus1 > 7 Then Throw New InvalidOperationException("num_slice_groups_minus1 out of range")
        If pps.NumSliceGroupsMinus1 > 0 Then
            ' FMO
            pps.SliceGroupMapType = CInt(br.Ue())
            Select Case pps.SliceGroupMapType
                Case 0
                    For iGroup As Integer = 0 To pps.NumSliceGroupsMinus1
                        br.Ue()    ' run_length_minus1[iGroup]
                    Next
                Case 2
                    For iGroup As Integer = 0 To pps.NumSliceGroupsMinus1 - 1
                        br.Ue()    ' top_left[iGroup]
                        br.Ue()    ' bottom_right[iGroup]
                    Next
                Case 3, 4, 5
                    br.U(1)        ' slice_group_change_direction_flag
                    br.Ue()        ' slice_group_change_rate_minus1
                Case 6
                    Dim picSizeInMapUnitsMinus1 As UInteger = br.Ue()
                    ' slice_group_id[i] is u(v) with v = Ceil(Log2(num_slice_groups_minus1 + 1))
                    Dim v As Integer = 0
                    While (1 << v) < pps.NumSliceGroupsMinus1 + 1
                        v += 1
                    End While
                    For i As Long = 0 To CLng(picSizeInMapUnitsMinus1)
                        br.U(v)
                    Next
            End Select
        End If

        pps.NumRefIdxL0DefaultActiveMinus1 = CByte(br.Ue() And &H3FUI)
        pps.NumRefIdxL1DefaultActiveMinus1 = CByte(br.Ue() And &H3FUI)
        pps.WeightedPredFlag = br.Flag()
        pps.WeightedBipredIdc = CByte(br.U(2) And 3UI)
        br.Se()    ' pic_init_qp_minus26
        br.Se()    ' pic_init_qs_minus26
        br.Se()    ' chroma_qp_index_offset

        pps.DeblockingFilterControlPresent = br.Flag()
        br.U(1)    ' constrained_intra_pred_flag
        pps.RedundantPicCntPresentFlag = br.Flag()
        Return pps
    End Function

End Module
