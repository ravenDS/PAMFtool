' AvcAnalyzer.vb - github.com/ravenDS/PAMFtool
'
' per-picture analysis of aH.264 stream for PAMF muxer:
'   - random access points (IDR, or an all-intra picture with a recovery point SEI, recovery_frame_cnt 0)
'   - SPS presence per AU (Sony starts a ps2 block at every RAP that carries an SPS)
'   - reference flag (nal_ref_idc)
'   - picture order count (8.2.1) -> display order -> exact PTS / DTS
'   - deblocking filter usage over every slice of the stream (codec_info deblockingFilterFlag)

Public Class AvcPictureInfo
    Public IsIdr As Boolean
    Public IsReference As Boolean
    Public AllIntra As Boolean
    Public HasRecoveryPoint0 As Boolean       ' recovery point SEI with recovery_frame_cnt = 0
    Public HasSps As Boolean
    Public HasVcl As Boolean
    Public FieldPic As Boolean
    Public BottomField As Boolean
    Public HasMmco5 As Boolean
    Public FrameNum As Integer
    Public Poc As Long                        ' PicOrderCnt (frame: min(top, bottom)), after an mmco5 reset
    Public Period As Integer                  ' POC period: increments at IDR and mmco5 pictures
    Public SliceHeaderOk As Boolean           ' first slice header parsed
    Public ReadOnly Property IsRap As Boolean
        Get
            Return IsIdr OrElse (AllIntra AndAlso HasRecoveryPoint0)
        End Get
    End Property
    ' timing, in field units (a frame AU = 2, a field AU = 1)
    Public DecodeUnits As Long                ' sum of the units of all AUs before this one in decode order
    Public DisplayUnits As Long               ' sum of the units of all AUs displayed before this one
    Public Units As Integer
End Class

Public Class AvcStreamAnalysis
    Public ReadOnly Pictures As New List(Of AvcPictureInfo)()
    Public FirstSps As H264SpsInfo
    Public FirstPps As H264PpsInfo
    Public SliceCount As Long
    Public SlicesDeblockOn As Long           ' disable_deblocking_filter_idc 0 or 2 (or not coded = 0)
    Public SlicesDeblockOff As Long          ' disable_deblocking_filter_idc 1
    Public SliceParseFailures As Long
    Public SpsParseFailures As Integer
    Public PpsParseFailures As Integer
    Public TotalUnits As Long                ' sum of all AU units
    Public DelayUnits As Long                ' DTS offset D in field units (>= 2 = one frame)
End Class

Friend Module AvcAnalyzer

    Private Class SliceHeader
        Public SliceType As Integer           ' 0 P, 1 B, 2 I, 3 SP, 4 SI
        Public PpsId As Integer
        Public FrameNum As Integer
        Public FieldPic As Boolean
        Public BottomField As Boolean
        Public PocLsb As Integer
        Public DeltaPocBottom As Integer
        Public DeltaPoc0 As Integer
        Public DeltaPoc1 As Integer
        Public HasMmco5 As Boolean
        Public DisableDeblockIdc As Integer
    End Class

    Public Function Analyze(bytes As Byte(), auStarts As List(Of Integer)) As AvcStreamAnalysis
        Dim r As New AvcStreamAnalysis()
        Dim spsById As New Dictionary(Of Integer, H264SpsInfo)()
        Dim ppsById As New Dictionary(Of Integer, H264PpsInfo)()
        Dim firstSliceOf As New List(Of SliceHeader)()
        Dim nalRefOf As New List(Of Integer)()
        Dim idrOf As New List(Of Boolean)()
        Dim spsOfPic As New List(Of H264SpsInfo)()

        For k As Integer = 0 To auStarts.Count - 1
            Dim a As Integer = auStarts(k)
            Dim b As Integer = If(k + 1 < auStarts.Count, auStarts(k + 1), bytes.Length)
            Dim pic As New AvcPictureInfo() With {.AllIntra = True}
            Dim first As SliceHeader = Nothing
            Dim firstRef As Integer = 0
            Dim firstSps As H264SpsInfo = Nothing
            Dim nal As NalSpan
            Dim p As Integer = a
            While NextNal(bytes, p, b, nal)
                p = nal.PayloadEnd
                Dim hb As Integer = bytes(nal.Header)
                Dim nt As Integer = hb And &H1F
                Dim nri As Integer = (hb >> 5) And 3
                Select Case nt
                    Case 7
                        pic.HasSps = True
                        Dim sps As H264SpsInfo = H264SpsParser.ParseSpsNal(bytes, nal.Header, nal.PayloadEnd)
                        If sps Is Nothing Then
                            r.SpsParseFailures += 1
                        Else
                            spsById(sps.SpsId) = sps
                            If r.FirstSps Is Nothing Then r.FirstSps = sps
                        End If
                    Case 8
                        Dim pps As H264PpsInfo = H264PpsParser.ParsePpsNal(bytes, nal.Header, nal.PayloadEnd)
                        If pps Is Nothing Then
                            r.PpsParseFailures += 1
                        Else
                            ppsById(pps.PpsId) = pps
                            If r.FirstPps Is Nothing Then r.FirstPps = pps
                        End If
                    Case 6
                        If SeiHasRecoveryPoint0(bytes, nal.Header, nal.PayloadEnd) Then pic.HasRecoveryPoint0 = True
                    Case 1, 5
                        pic.HasVcl = True
                        If nt = 5 Then pic.IsIdr = True
                        If nri <> 0 Then pic.IsReference = True
                        r.SliceCount += 1
                        Dim sh As SliceHeader = Nothing
                        Dim spsUsed As H264SpsInfo = Nothing
                        Try
                            sh = ParseSliceHeader(bytes, nal.Header, nal.PayloadEnd, nt, nri, spsById, ppsById, spsUsed)
                        Catch
                            sh = Nothing
                        End Try
                        If sh Is Nothing Then
                            r.SliceParseFailures += 1
                            pic.AllIntra = False         ' unknown slice type: never call it intra
                        Else
                            If sh.DisableDeblockIdc = 1 Then
                                r.SlicesDeblockOff += 1
                            Else
                                r.SlicesDeblockOn += 1
                            End If
                            If sh.SliceType <> 2 AndAlso sh.SliceType <> 4 Then pic.AllIntra = False
                            If first Is Nothing Then
                                first = sh
                                firstRef = nri
                                firstSps = spsUsed
                            End If
                        End If
                End Select
            End While
            If Not pic.HasVcl Then pic.AllIntra = False
            r.Pictures.Add(pic)
            firstSliceOf.Add(first)
            nalRefOf.Add(firstRef)
            idrOf.Add(pic.IsIdr)
            spsOfPic.Add(firstSps)
        Next

        ComputePoc(r, firstSliceOf, nalRefOf, spsOfPic)
        ComputeOrder(r)
        Return r
    End Function

    ' sei_rbsp(): walk sei_message()s, report a recovery point (payloadType 6) with recovery_frame_cnt = 0
    Private Function SeiHasRecoveryPoint0(buf As Byte(), nalHeader As Integer, nalEnd As Integer) As Boolean
        Dim pl As Byte()
        Try
            pl = Unescape(buf, nalHeader + 1, nalEnd)
        Catch
            Return False
        End Try
        Dim q As Integer = 0
        While q < pl.Length
            If pl(q) = &H80 AndAlso q = pl.Length - 1 Then Exit While     ' rbsp_trailing_bits
            Dim t As Integer = 0
            While q < pl.Length AndAlso pl(q) = &HFF
                t += 255 : q += 1
            End While
            If q >= pl.Length Then Exit While
            t += pl(q) : q += 1
            Dim sz As Integer = 0
            While q < pl.Length AndAlso pl(q) = &HFF
                sz += 255 : q += 1
            End While
            If q >= pl.Length Then Exit While
            sz += pl(q) : q += 1
            If q + sz > pl.Length Then Exit While
            If t = 6 AndAlso sz > 0 Then
                Dim payload(sz - 1) As Byte
                Array.Copy(pl, q, payload, 0, sz)
                Try
                    Dim br As New RbspReader(payload)
                    If br.Ue() = 0UI Then Return True              ' recovery_frame_cnt
                Catch
                End Try
            End If
            q += sz
        End While
        Return False
    End Function

    ' slice_header() up to disable_deblocking_filter_idc (7.3.3)
    Private Function ParseSliceHeader(buf As Byte(), nalHeader As Integer, nalEnd As Integer,
                                      nalType As Integer, nalRefIdc As Integer,
                                      spsById As Dictionary(Of Integer, H264SpsInfo),
                                      ppsById As Dictionary(Of Integer, H264PpsInfo),
                                      ByRef spsUsed As H264SpsInfo) As SliceHeader
        Dim rbsp As Byte() = Unescape(buf, nalHeader + 1, nalEnd, 2048)
        Dim br As New RbspReader(rbsp)
        Dim sh As New SliceHeader()
        br.Ue()                                                    ' first_mb_in_slice
        Dim st As UInteger = br.Ue()
        If st > 9UI Then Return Nothing
        sh.SliceType = CInt(st Mod 5UI)
        sh.PpsId = CInt(br.Ue())
        Dim pps As H264PpsInfo = Nothing
        If Not ppsById.TryGetValue(sh.PpsId, pps) Then Return Nothing
        Dim sps As H264SpsInfo = Nothing
        If Not spsById.TryGetValue(pps.SpsId, sps) Then Return Nothing
        spsUsed = sps

        If sps.SeparateColourPlaneFlag Then br.U(2)                ' colour_plane_id
        sh.FrameNum = CInt(br.U(sps.Log2MaxFrameNumMinus4 + 4))
        If Not sps.FrameMbsOnlyFlag Then
            sh.FieldPic = br.Flag()
            If sh.FieldPic Then sh.BottomField = br.Flag()
        End If
        If nalType = 5 Then br.Ue()                                ' idr_pic_id
        If sps.PicOrderCntType = 0 Then
            sh.PocLsb = CInt(br.U(sps.Log2MaxPicOrderCntLsbMinus4 + 4))
            If pps.BottomFieldPicOrderInFramePresentFlag AndAlso Not sh.FieldPic Then
                sh.DeltaPocBottom = br.Se()
            End If
        End If
        If sps.PicOrderCntType = 1 AndAlso Not sps.DeltaPicOrderAlwaysZeroFlag Then
            sh.DeltaPoc0 = br.Se()
            If pps.BottomFieldPicOrderInFramePresentFlag AndAlso Not sh.FieldPic Then
                sh.DeltaPoc1 = br.Se()
            End If
        End If
        If pps.RedundantPicCntPresentFlag Then br.Ue()             ' redundant_pic_cnt

        Dim isB As Boolean = (sh.SliceType = 1)
        Dim isP As Boolean = (sh.SliceType = 0 OrElse sh.SliceType = 3)
        If isB Then br.U(1)                                        ' direct_spatial_mv_pred_flag
        Dim numL0 As Integer = pps.NumRefIdxL0DefaultActiveMinus1
        Dim numL1 As Integer = pps.NumRefIdxL1DefaultActiveMinus1
        If isP OrElse isB Then
            If br.Flag() Then                                      ' num_ref_idx_active_override_flag
                numL0 = CInt(br.Ue())
                If isB Then numL1 = CInt(br.Ue())
            End If
        End If
        If numL0 > 31 OrElse numL1 > 31 Then Return Nothing

        ' ref_pic_list_modification()
        If sh.SliceType <> 2 AndAlso sh.SliceType <> 4 Then
            SkipRefPicListModification(br)
            If isB Then SkipRefPicListModification(br)
        End If

        If (pps.WeightedPredFlag AndAlso isP) OrElse (pps.WeightedBipredIdc = 1 AndAlso isB) Then
            ' pred_weight_table()
            br.Ue()                                                ' luma_log2_weight_denom
            Dim chroma As Boolean = (sps.ChromaArrayType <> 0)
            If chroma Then br.Ue()                                 ' chroma_log2_weight_denom
            SkipWeights(br, numL0, chroma)
            If isB Then SkipWeights(br, numL1, chroma)
        End If

        If nalRefIdc <> 0 Then
            ' dec_ref_pic_marking()
            If nalType = 5 Then
                br.U(1)                                            ' no_output_of_prior_pics_flag
                br.U(1)                                            ' long_term_reference_flag
            ElseIf br.Flag() Then                                  ' adaptive_ref_pic_marking_mode_flag
                Dim guard As Integer = 0
                Do
                    Dim mmco As UInteger = br.Ue()
                    If mmco = 0UI Then Exit Do
                    If mmco = 5UI Then sh.HasMmco5 = True
                    If mmco = 1UI OrElse mmco = 3UI Then br.Ue()    ' difference_of_pic_nums_minus1
                    If mmco = 2UI Then br.Ue()                      ' long_term_pic_num
                    If mmco = 3UI OrElse mmco = 6UI Then br.Ue()    ' long_term_frame_idx
                    If mmco = 4UI Then br.Ue()                      ' max_long_term_frame_idx_plus1
                    If mmco > 6UI Then Return Nothing
                    guard += 1
                    If guard > 66 Then Return Nothing
                Loop
            End If
        End If

        If pps.CabacFlag AndAlso sh.SliceType <> 2 AndAlso sh.SliceType <> 4 Then br.Ue()   ' cabac_init_idc
        br.Se()                                                    ' slice_qp_delta
        If sh.SliceType = 3 OrElse sh.SliceType = 4 Then
            If sh.SliceType = 3 Then br.U(1)                       ' sp_for_switch_flag
            br.Se()                                                ' slice_qs_delta
        End If
        sh.DisableDeblockIdc = 0                                   ' inferred 0 when not coded (filter on)
        If pps.DeblockingFilterControlPresent Then
            Dim idc As UInteger = br.Ue()
            If idc > 2UI Then Return Nothing
            sh.DisableDeblockIdc = CInt(idc)
        End If
        Return sh
    End Function

    Private Sub SkipRefPicListModification(br As RbspReader)
        If Not br.Flag() Then Return                               ' ref_pic_list_modification_flag_lX
        Dim guard As Integer = 0
        Do
            Dim idc As UInteger = br.Ue()
            If idc = 3UI Then Exit Do
            If idc > 5UI Then Throw New InvalidOperationException("modification_of_pic_nums_idc")
            br.Ue()                       ' abs_diff_pic_num_minus1 / long_term_pic_num / abs_diff_view_idx_minus1
            guard += 1
            If guard > 66 Then Throw New InvalidOperationException("ref_pic_list_modification loop")
        Loop
    End Sub

    Private Sub SkipWeights(br As RbspReader, numMinus1 As Integer, chroma As Boolean)
        For i As Integer = 0 To numMinus1
            If br.Flag() Then                                      ' luma_weight_lX_flag
                br.Se() : br.Se()
            End If
            If chroma AndAlso br.Flag() Then                       ' chroma_weight_lX_flag
                br.Se() : br.Se() : br.Se() : br.Se()
            End If
        Next
    End Sub

    ' decoding process for picture order count, types 0/1/2
    Private Sub ComputePoc(r As AvcStreamAnalysis, first As List(Of SliceHeader),
                           nalRef As List(Of Integer), spsOf As List(Of H264SpsInfo))
        Dim prevMsb As Long = 0, prevLsb As Long = 0               ' type 0 (from the previous reference picture)
        Dim prevFrameNumOffset As Long = 0, prevFrameNum As Long = 0
        Dim period As Integer = 0
        Dim lastPoc As Long = 0

        For k As Integer = 0 To r.Pictures.Count - 1
            Dim pic As AvcPictureInfo = r.Pictures(k)
            Dim sh As SliceHeader = first(k)
            Dim sps As H264SpsInfo = spsOf(k)
            pic.Units = 2
            If sh Is Nothing OrElse sps Is Nothing Then
                ' unparsable picture: keep it right after previous one in display order
                pic.SliceHeaderOk = False
                pic.Period = period
                pic.Poc = lastPoc
                Continue For
            End If
            pic.SliceHeaderOk = True
            pic.FieldPic = sh.FieldPic
            pic.BottomField = sh.BottomField
            pic.FrameNum = sh.FrameNum
            pic.HasMmco5 = sh.HasMmco5
            If sh.FieldPic Then pic.Units = 1
            Dim isRef As Boolean = nalRef(k) <> 0
            Dim maxFrameNum As Long = 1L << (sps.Log2MaxFrameNumMinus4 + 4)

            If pic.IsIdr Then period += 1
            Dim top As Long = 0, bottom As Long = 0
            Select Case sps.PicOrderCntType
                Case 0
                    If pic.IsIdr Then
                        prevMsb = 0 : prevLsb = 0
                    End If
                    Dim maxLsb As Long = 1L << (sps.Log2MaxPicOrderCntLsbMinus4 + 4)
                    Dim lsb As Long = sh.PocLsb
                    Dim msb As Long
                    If lsb < prevLsb AndAlso (prevLsb - lsb) >= maxLsb \ 2 Then
                        msb = prevMsb + maxLsb
                    ElseIf lsb > prevLsb AndAlso (lsb - prevLsb) > maxLsb \ 2 Then
                        msb = prevMsb - maxLsb
                    Else
                        msb = prevMsb
                    End If
                    If Not sh.FieldPic Then
                        top = msb + lsb
                        bottom = top + sh.DeltaPocBottom
                    ElseIf Not sh.BottomField Then
                        top = msb + lsb : bottom = top
                    Else
                        bottom = msb + lsb : top = bottom
                    End If
                    If isRef Then
                        If sh.HasMmco5 Then
                            ' 8.2.1: after mmco5 the next picture uses prevPicOrderCntMsb = 0 and
                            ' prevPicOrderCntLsb = TopFieldOrderCnt of this picture after reset (0 for a bottom field)
                            prevMsb = 0
                            prevLsb = If(sh.FieldPic AndAlso sh.BottomField, 0L, top - Math.Min(top, If(sh.FieldPic, top, bottom)))
                        Else
                            prevMsb = msb : prevLsb = lsb
                        End If
                    End If
                Case 1, 2
                    Dim frameNumOffset As Long
                    If pic.IsIdr Then
                        frameNumOffset = 0
                    ElseIf prevFrameNum > sh.FrameNum Then
                        frameNumOffset = prevFrameNumOffset + maxFrameNum
                    Else
                        frameNumOffset = prevFrameNumOffset
                    End If
                    If sps.PicOrderCntType = 1 Then
                        Dim n As Integer = sps.OffsetForRefFrame.Length
                        Dim absFrameNum As Long = If(n <> 0, frameNumOffset + sh.FrameNum, 0L)
                        If Not isRef AndAlso absFrameNum > 0 Then absFrameNum -= 1
                        Dim expected As Long = 0
                        If absFrameNum > 0 Then
                            Dim deltaPerCycle As Long = 0
                            For Each o In sps.OffsetForRefFrame
                                deltaPerCycle += o
                            Next
                            Dim cycleCnt As Long = (absFrameNum - 1) \ n
                            Dim inCycle As Long = (absFrameNum - 1) Mod n
                            expected = cycleCnt * deltaPerCycle
                            For i As Integer = 0 To CInt(inCycle)
                                expected += sps.OffsetForRefFrame(i)
                            Next
                        End If
                        If Not isRef Then expected += sps.OffsetForNonRefPic
                        If Not sh.FieldPic Then
                            top = expected + sh.DeltaPoc0
                            bottom = top + sps.OffsetForTopToBottomField + sh.DeltaPoc1
                        ElseIf Not sh.BottomField Then
                            top = expected + sh.DeltaPoc0 : bottom = top
                        Else
                            bottom = expected + sps.OffsetForTopToBottomField + sh.DeltaPoc0 : top = bottom
                        End If
                    Else
                        Dim temp As Long
                        If pic.IsIdr Then
                            temp = 0
                        ElseIf Not isRef Then
                            temp = 2 * (frameNumOffset + sh.FrameNum) - 1
                        Else
                            temp = 2 * (frameNumOffset + sh.FrameNum)
                        End If
                        top = temp : bottom = temp
                    End If
                    prevFrameNumOffset = If(sh.HasMmco5, 0L, frameNumOffset)
                    prevFrameNum = If(sh.HasMmco5, 0L, CLng(sh.FrameNum))
            End Select

            Dim poc As Long
            If Not sh.FieldPic Then
                poc = Math.Min(top, bottom)
            ElseIf sh.BottomField Then
                poc = bottom
            Else
                poc = top
            End If
            If sh.HasMmco5 Then
                ' the picture is output after everything before it, its POC restarts at 0
                period += 1
                poc = 0
            End If
            pic.Period = period
            pic.Poc = poc
            lastPoc = poc
        Next
    End Sub

    ' display order = (period, POC, decode index), PTS/DTS in field units
    Private Sub ComputeOrder(r As AvcStreamAnalysis)
        Dim n As Integer = r.Pictures.Count
        Dim dec As Long = 0
        For k As Integer = 0 To n - 1
            r.Pictures(k).DecodeUnits = dec
            dec += r.Pictures(k).Units
        Next
        r.TotalUnits = dec
        Dim order(n - 1) As Integer
        For k As Integer = 0 To n - 1
            order(k) = k
        Next
        Array.Sort(order, Function(x As Integer, y As Integer)
                              Dim px As AvcPictureInfo = r.Pictures(x)
                              Dim py As AvcPictureInfo = r.Pictures(y)
                              If px.Period <> py.Period Then Return px.Period.CompareTo(py.Period)
                              If px.Poc <> py.Poc Then Return px.Poc.CompareTo(py.Poc)
                              Return x.CompareTo(y)
                          End Function)
        Dim disp As Long = 0
        For Each k In order
            r.Pictures(k).DisplayUnits = disp
            disp += r.Pictures(k).Units
        Next
        ' DTS = PTS_origin - D + decode position
        ' D >= one frame (minimum, also for I/P-only streams)
        Dim d As Long = 2
        For k As Integer = 0 To n - 1
            Dim lag As Long = r.Pictures(k).DecodeUnits - r.Pictures(k).DisplayUnits
            If lag > d Then d = lag
        Next
        r.DelayUnits = d
    End Sub

End Module
