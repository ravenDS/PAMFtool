'  MpegSequenceHeader.vb - github.com/ravenDS/PAMFtool
'  MPEG-2 video sequence-header parser

Public Class M2vSequenceInfo
    Public Property WidthPixels As Integer
    Public Property HeightPixels As Integer
    Public Property AspectRatioInfo As Byte
    Public Property FrameRateCode As Byte
    Public Property ProfileAndLevel As Byte
    Public Property ProgressiveSequence As Boolean
    Public Property ChromaFormat As Byte
    Public Property FrameRateExtN As Byte
    Public Property FrameRateExtD As Byte
    Public Property HasExtension As Boolean

    ' PAMF muxer forwards them from bitstream when colour_description_flag = 1
    Public Property HasDisplayExtension As Boolean
    Public Property VideoFormat As Byte             ' 0..5 (5 = unspecified)
    Public Property HasColourDescription As Boolean
    Public Property ColourPrimaries As Byte         ' 1 = BT.709, 6 = SMPTE 170M (BT.601)
    Public Property TransferCharacteristics As Byte
    Public Property MatrixCoefficients As Byte

    ' Resolved fps per ISO/IEC 13818-2 §6.3.3:
    '   actual = base[frame_rate_code] × (frame_rate_extension_n + 1)
    '                                  / (frame_rate_extension_d + 1)
    Public ReadOnly Property FrameRate As Double
        Get
            Dim baseFps As Double = BaseFpsFromCode(FrameRateCode)
            If baseFps = 0.0 Then Return 0.0
            Return baseFps * (FrameRateExtN + 1) / (FrameRateExtD + 1)
        End Get
    End Property

    Private Shared Function BaseFpsFromCode(code As Byte) As Double
        Select Case code
            Case 1 : Return 24000.0 / 1001.0   ' 23.976
            Case 2 : Return 24.0
            Case 3 : Return 25.0
            Case 4 : Return 30000.0 / 1001.0   ' 29.97
            Case 5 : Return 30.0
            Case 6 : Return 50.0
            Case 7 : Return 60000.0 / 1001.0   ' 59.94
            Case 8 : Return 60.0
            Case Else : Return 0.0
        End Select
    End Function
End Class

Friend Module MpegSequenceHeaderParser

    Public Function ParseFirstSequenceHeader(buf As Byte(),
                                             off As Integer,
                                             length As Integer) As M2vSequenceInfo
        Dim [end] As Integer = Math.Min(buf.Length, off + length)

        Dim shAt As Integer = FindStartCode(buf, off, [end], &HB3)
        If shAt < 0 Then Return Nothing
        Dim info As M2vSequenceInfo = Nothing
        Try
            info = ParseSequenceHeader(buf, shAt + 4, [end])
        Catch
            Return Nothing
        End Try
        If info Is Nothing Then Return Nothing

        ' look for extension start codes (0x000001B5) between sequence_header and the first picture_start_code (0x00000100).
        ' two extension types matter for PAMF:
        ' ext_id=1 (sequence_extension) for chroma /progressive / FRC extensions
        ' ext_id=2 (sequence_display_extension) for video_format and colour_description fields that Sony forwards into ci(20..23)
        Dim sawSeq As Boolean = False
        Dim sawDisp As Boolean = False
        Dim scanAt As Integer = shAt + 4
        While scanAt < [end] - 4 AndAlso Not (sawSeq AndAlso sawDisp)
            ' stop if we've passed into picture data
            If buf(scanAt) = 0 AndAlso buf(scanAt + 1) = 0 _
            AndAlso buf(scanAt + 2) = 1 AndAlso buf(scanAt + 3) = 0 Then
                Exit While
            End If
            Dim extAt As Integer = FindStartCode(buf, scanAt, [end], &HB5)
            If extAt < 0 Then Exit While
            ' bound the search so we don't wander into picture data looking for an extension that isn't there
            If extAt - shAt > 1024 Then Exit While
            Dim extId As Integer = (buf(extAt + 4) >> 4) And &HF
            Try
                Select Case extId
                    Case 1
                        ApplySequenceExtension(info, buf, extAt + 4, [end])
                        sawSeq = True
                    Case 2
                        ApplySequenceDisplayExtension(info, buf, extAt + 4, [end])
                        sawDisp = True
                End Select
            Catch
                ' partial / malformed, keep what we have
            End Try
            scanAt = extAt + 4
        End While

        Return info
    End Function

    Private Function FindStartCode(buf As Byte(),
                                   off As Integer,
                                   [end] As Integer,
                                   code As Byte) As Integer
        Dim i As Integer = off
        While i < [end] - 4
            If buf(i) = 0 AndAlso buf(i + 1) = 0 _
            AndAlso buf(i + 2) = 1 AndAlso buf(i + 3) = code Then
                Return i
            End If
            i += 1
        End While
        Return -1
    End Function

    ' sequence header fixed fields:
    '   12 bits  horizontal_size_value
    '   12 bits  vertical_size_value
    '    4 bits  aspect_ratio_information
    '    4 bits  frame_rate_code
    '   18 bits  bit_rate_value
    '    1 bit   marker_bit
    '   10 bits  vbv_buffer_size_value
    '    1 bit   constrained_parameters_flag
    '    1 bit   load_intra_quantiser_matrix
    '   (64 × 8 if set above)
    '    1 bit   load_non_intra_quantiser_matrix
    '   (64 × 8 if set above)
    Private Function ParseSequenceHeader(buf As Byte(),
                                         p As Integer,
                                         [end] As Integer) As M2vSequenceInfo
        If p + 8 > [end] Then Return Nothing

        Dim b0 As Integer = buf(p + 0)
        Dim b1 As Integer = buf(p + 1)
        Dim b2 As Integer = buf(p + 2)
        Dim b3 As Integer = buf(p + 3)

        Dim hSize As Integer = (b0 << 4) Or (b1 >> 4)
        Dim vSize As Integer = ((b1 And &HF) << 8) Or b2
        Dim aspect As Integer = b3 >> 4
        Dim frc As Integer = b3 And &HF

        Dim info As New M2vSequenceInfo() With {
            .WidthPixels = hSize,
            .HeightPixels = vSize,
            .AspectRatioInfo = CByte(aspect),
            .FrameRateCode = CByte(frc),
            .ProgressiveSequence = True,    ' default; will be overridden by ext
            .HasExtension = False
        }
        Return info
    End Function

    ' sequence_extension layout bits from MSB of the byte directly following the start code:
    '    4 bits  extension_start_code_identifier  (=1)
    '    8 bits  profile_and_level_indication
    '    1 bit   progressive_sequence
    '    2 bits  chroma_format
    '    2 bits  horizontal_size_extension
    '    2 bits  vertical_size_extension
    '   12 bits  bit_rate_extension
    '    1 bit   marker_bit
    '    8 bits  vbv_buffer_size_extension
    '    1 bit   low_delay
    '    2 bits  frame_rate_extension_n
    '    5 bits  frame_rate_extension_d
    '
    ' total = 48 bits = 6 bytes
    Private Sub ApplySequenceExtension(info As M2vSequenceInfo,
                                       buf As Byte(),
                                       p As Integer,
                                       [end] As Integer)
        If p + 6 > [end] Then Return
        Dim br As New BitReader(buf, p)
        Dim extId As Integer = br.U(4)
        If extId <> 1 Then Return

        info.HasExtension = True
        info.ProfileAndLevel = CByte(br.U(8))
        info.ProgressiveSequence = (br.U(1) <> 0)
        info.ChromaFormat = CByte(br.U(2))
        Dim hExt As Integer = br.U(2)
        Dim vExt As Integer = br.U(2)
        br.U(12)                                   ' bit_rate_extension
        br.U(1)                                    ' marker_bit
        br.U(8)                                    ' vbv_buffer_size_extension
        br.U(1)                                    ' low_delay
        info.FrameRateExtN = CByte(br.U(2))
        info.FrameRateExtD = CByte(br.U(5))

        ' apply the extension bits to the dimensions parsed from the seq header proper (HD streams use these, for 1280x720 theyre zero)
        info.WidthPixels = (hExt << 12) Or info.WidthPixels
        info.HeightPixels = (vExt << 12) Or info.HeightPixels
    End Sub

    ' sequence_display_extension layout:
    '    4 bits  extension_start_code_identifier  (=2)
    '    3 bits  video_format                     (0..5; 5 = unspecified)
    '    1 bit   colour_description_flag
    '    if (colour_description_flag)
    '       {8 bits  colour_primaries
    '        8 bits  transfer_characteristics
    '        8 bits  matrix_coefficients}
    '   14 bits  display_horizontal_size
    '    1 bit   marker_bit
    '   14 bits  display_vertical_size
    '
    ' without colour_description present: 4+3+1+14+1+14 = 37 bits = 5 bytes
    ' with colour_description present:    +24 bits      = 61 bits = 8 bytes
    Private Sub ApplySequenceDisplayExtension(info As M2vSequenceInfo,
                                              buf As Byte(),
                                              p As Integer,
                                              [end] As Integer)
        If p + 5 > [end] Then Return
        Dim br As New BitReader(buf, p)
        Dim extId As Integer = br.U(4)
        If extId <> 2 Then Return

        info.HasDisplayExtension = True
        info.VideoFormat = CByte(br.U(3))
        Dim hasColour As Integer = br.U(1)
        If hasColour <> 0 Then
            If p + 8 > [end] Then Return
            info.HasColourDescription = True
            info.ColourPrimaries = CByte(br.U(8))
            info.TransferCharacteristics = CByte(br.U(8))
            info.MatrixCoefficients = CByte(br.U(8))
        End If
        ' display_horizontal_size / marker_bit / display_vertical_size is display-only info (overscan / pillarbox) that PAMF doesn't carry
    End Sub

    ' minimal MSB-first bit reader against absolute byte offset
    Private Class BitReader
        Private ReadOnly _buf As Byte()
        Private _bytePos As Integer
        Private _bitPos As Integer
        Public Sub New(buf As Byte(), byteOffset As Integer)
            _buf = buf
            _bytePos = byteOffset
            _bitPos = 0
        End Sub
        Public Function U(n As Integer) As Integer
            Dim v As Integer = 0
            For i As Integer = 0 To n - 1
                If _bytePos >= _buf.Length Then
                    Throw New InvalidOperationException("BitReader past end of buffer")
                End If
                Dim bit As Integer = (_buf(_bytePos) >> (7 - _bitPos)) And 1
                v = (v << 1) Or bit
                _bitPos += 1
                If _bitPos = 8 Then
                    _bitPos = 0
                    _bytePos += 1
                End If
            Next
            Return v
        End Function
    End Class

End Module