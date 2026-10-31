' RbspReader.vb - github.com/ravenDS/PAMFtool
'
' shared helpers for H.264 / MPEG-2 header parsing
' Annex-B NAL walking, emulation-prevention removal, MSB-first bit reader with Exp-Golomb codes

Friend Module RbspUtil

    ' strip emulation_prevention_three_byte (00 00 03 -> 00 00) from buf[start, end)
    ' maxOut > 0 stops once that many output bytes exist (slice headers only need first few hundred)
    Public Function Unescape(buf As Byte(), start As Integer, [end] As Integer,
                             Optional maxOut As Integer = 0) As Byte()
        Dim n As Integer = [end] - start
        If n <= 0 Then Return Array.Empty(Of Byte)()
        Dim cap As Integer = If(maxOut > 0, Math.Min(maxOut, n), n)
        Dim outBuf(cap - 1) As Byte
        Dim o As Integer = 0
        Dim zeros As Integer = 0
        Dim i As Integer = start
        While i < [end] AndAlso o < cap
            Dim b As Byte = buf(i)
            If zeros >= 2 AndAlso b = 3 Then
                zeros = 0
                i += 1
                Continue While
            End If
            outBuf(o) = b
            o += 1
            zeros = If(b = 0, zeros + 1, 0)
            i += 1
        End While
        If o < cap Then ReDim Preserve outBuf(o - 1)
        Return outBuf
    End Function

    ' position of next 00 00 01 at or after p (before [end]); -1 if none
    Public Function FindStartCode(buf As Byte(), p As Integer, [end] As Integer) As Integer
        Dim i As Integer = p
        While i + 2 < [end]
            If buf(i + 2) > 1 Then
                i += 3
            ElseIf buf(i + 2) = 1 AndAlso buf(i + 1) = 0 AndAlso buf(i) = 0 Then
                Return i
            Else
                i += 1
            End If
        End While
        Return -1
    End Function

    ' one Annex-B NAL unit: Header = index of nal_unit_header byte
    ' PayloadEnd = first byte of the next start code (trailing zero bytes before it are excluded)
    Public Structure NalSpan
        Public Header As Integer
        Public PayloadEnd As Integer
    End Structure

    ' next NAL unit whose start code begins at or after p; False when there is none
    Public Function NextNal(buf As Byte(), p As Integer, [end] As Integer, ByRef nal As NalSpan) As Boolean
        Dim sc As Integer = FindStartCode(buf, p, [end])
        If sc < 0 OrElse sc + 3 >= [end] Then Return False
        Dim hdr As Integer = sc + 3
        Dim nxt As Integer = FindStartCode(buf, hdr, [end])
        Dim pe As Integer = If(nxt < 0, [end], nxt)
        While pe > hdr + 1 AndAlso buf(pe - 1) = 0
            pe -= 1
        End While
        nal = New NalSpan() With {.Header = hdr, .PayloadEnd = pe}
        Return True
    End Function

    Public Function ListNals(buf As Byte(), start As Integer, [end] As Integer) As List(Of NalSpan)
        Dim result As New List(Of NalSpan)()
        Dim sc As Integer = FindStartCode(buf, start, [end])
        While sc >= 0
            Dim hdr As Integer = sc + 3
            If hdr >= [end] Then Exit While
            Dim nxt As Integer = FindStartCode(buf, hdr, [end])
            Dim pe As Integer = If(nxt < 0, [end], nxt)
            While pe > hdr + 1 AndAlso buf(pe - 1) = 0
                pe -= 1
            End While
            result.Add(New NalSpan() With {.Header = hdr, .PayloadEnd = pe})
            sc = nxt
        End While
        Return result
    End Function

End Module

' MSB-first bit reader over an RBSP (emulation prevention already removed)
Friend Class RbspReader
    Private ReadOnly _data As Byte()
    Private _pos As Integer

    Public Sub New(data As Byte())
        _data = data
    End Sub

    Public ReadOnly Property BitPosition As Integer
        Get
            Return _pos
        End Get
    End Property

    Public ReadOnly Property BitsLeft As Integer
        Get
            Return _data.Length * 8 - _pos
        End Get
    End Property

    Public Function U(n As Integer) As UInteger
        If n < 0 OrElse n > 32 Then Throw New ArgumentOutOfRangeException(NameOf(n))
        If _pos + n > _data.Length * 8 Then Throw New InvalidOperationException("RBSP read past end")
        Dim v As ULong = 0UL
        For k As Integer = 0 To n - 1
            Dim bit As Integer = (_data(_pos >> 3) >> (7 - (_pos And 7))) And 1
            v = (v << 1) Or CULng(bit)
            _pos += 1
        Next
        Return CUInt(v)
    End Function

    Public Function Flag() As Boolean
        Return U(1) <> 0UI
    End Function

    Public Sub Skip(n As Integer)
        If _pos + n > _data.Length * 8 Then Throw New InvalidOperationException("RBSP read past end")
        _pos += n
    End Sub

    Public Function Ue() As UInteger
        Dim zeros As Integer = 0
        While U(1) = 0UI
            zeros += 1
            If zeros > 31 Then Throw New InvalidOperationException("Exp-Golomb code too long")
        End While
        If zeros = 0 Then Return 0UI
        Return CUInt(((1UL << zeros) - 1UL) + CULng(U(zeros)))
    End Function

    Public Function Se() As Integer
        Dim k As UInteger = Ue()
        If (k And 1UI) = 1UI Then Return CInt((CLng(k) + 1L) \ 2L)
        Return -CInt(CLng(k) \ 2L)
    End Function
End Class
