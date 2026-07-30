' Mpeg2PsPrimitives.vb - github.com/ravenDS/PAMFtool
' Byte-level encoders for MPEG-2 Program Stream structures

Imports System.IO

Namespace PamfMux

    Public Module Mpeg2PsPrimitives

        Public Const SectorSize As Integer = 2048

        Public Const SC_PackHeader As Byte = &HBA
        Public Const SC_SystemHeader As Byte = &HBB
        Public Const SC_ProgStreamMap As Byte = &HBC
        Public Const SC_PrivateStream1 As Byte = &HBD
        Public Const SC_PaddingStream As Byte = &HBE
        Public Const SC_PrivateStream2 As Byte = &HBF
        Public Const SC_ProgramEnd As Byte = &HB9

        Public Const PackHeaderLen As Integer = 14
        Public Const SystemHeaderLen As Integer = 18
        Public Const VideoPesHeaderLen As Integer = 22
        Public Const VideoPesHeaderContinuationLen As Integer = 9   ' see WriteVideoPesHeaderContinuation
        Public Const AudioPesHeaderLen As Integer = 21
        Public Const AudioSubHeaderLen As Integer = 4

        ' pack_header (14 bytes)

        Public Sub WritePackHeader(out As Stream,
                                   scrBase33 As Long,
                                   scrExt9 As Integer,
                                   muxRateUnits As Integer)
            If scrBase33 < 0L OrElse scrBase33 >= (1L << 33) Then
                Throw New ArgumentOutOfRangeException(NameOf(scrBase33))
            End If
            If scrExt9 < 0 OrElse scrExt9 >= 512 Then
                Throw New ArgumentOutOfRangeException(NameOf(scrExt9))
            End If
            If muxRateUnits < 0 OrElse muxRateUnits >= (1 << 22) Then
                Throw New ArgumentOutOfRangeException(NameOf(muxRateUnits))
            End If

            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_PackHeader)

            Dim s32_30 As Integer = CInt((scrBase33 >> 30) And &H7L)
            Dim s29_28 As Integer = CInt((scrBase33 >> 28) And &H3L)
            Dim s27_20 As Integer = CInt((scrBase33 >> 20) And &HFFL)
            Dim s19_15 As Integer = CInt((scrBase33 >> 15) And &H1FL)
            Dim s14_13 As Integer = CInt((scrBase33 >> 13) And &H3L)
            Dim s12_5 As Integer = CInt((scrBase33 >> 5) And &HFFL)
            Dim s4_0 As Integer = CInt(scrBase33 And &H1FL)
            Dim e8_7 As Integer = (scrExt9 >> 7) And &H3
            Dim e6_0 As Integer = scrExt9 And &H7F

            out.WriteByte(CByte(&H40 Or (s32_30 << 3) Or &H4 Or s29_28))
            out.WriteByte(CByte(s27_20))
            out.WriteByte(CByte((s19_15 << 3) Or &H4 Or s14_13))
            out.WriteByte(CByte(s12_5))
            out.WriteByte(CByte((s4_0 << 3) Or &H4 Or e8_7))
            out.WriteByte(CByte((e6_0 << 1) Or &H1))

            out.WriteByte(CByte((muxRateUnits >> 14) And &HFF))
            out.WriteByte(CByte((muxRateUnits >> 6) And &HFF))
            out.WriteByte(CByte(((muxRateUnits And &H3F) << 2) Or &H3))

            out.WriteByte(&HF8)
        End Sub

        ' system_header (18 bytes)

        Public Sub WriteSystemHeader(out As Stream,
                                     rateBoundUnits As Integer,
                                     videoPstdBufferSize As Integer)
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_SystemHeader)
            out.WriteByte(0) : out.WriteByte(12)

            Dim r As Integer = rateBoundUnits And &H3FFFFF
            out.WriteByte(CByte(&H80 Or ((r >> 15) And &H7F)))
            out.WriteByte(CByte((r >> 7) And &HFF))
            out.WriteByte(CByte(((r And &H7F) << 1) Or 1))

            out.WriteByte(&H80)
            out.WriteByte(&HF0)
            out.WriteByte(&H7F)

            Dim vsize As Integer = videoPstdBufferSize And &H1FFF
            out.WriteByte(&HB9)
            out.WriteByte(CByte(&HC0 Or &H20 Or ((vsize >> 8) And &H1F)))
            out.WriteByte(CByte(vsize And &HFF))

            out.WriteByte(&HBD)
            out.WriteByte(&HE7)
            out.WriteByte(&H28)
        End Sub

        ' 5-byte timestamp (PTS or DTS)

        Public Sub WriteTimestamp5(out As Stream, ts33 As Long, prefix4Bits As Integer)
            If ts33 < 0L OrElse ts33 >= (1L << 33) Then
                Throw New ArgumentOutOfRangeException(NameOf(ts33))
            End If
            Dim t32_30 As Integer = CInt((ts33 >> 30) And &H7L)
            Dim t29_15 As Integer = CInt((ts33 >> 15) And &H7FFFL)
            Dim t14_0 As Integer = CInt(ts33 And &H7FFFL)
            out.WriteByte(CByte(((prefix4Bits And &HF) << 4) Or (t32_30 << 1) Or 1))
            out.WriteByte(CByte((t29_15 >> 7) And &HFF))
            out.WriteByte(CByte(((t29_15 And &H7F) << 1) Or 1))
            out.WriteByte(CByte((t14_0 >> 7) And &HFF))
            out.WriteByte(CByte(((t14_0 And &H7F) << 1) Or 1))
        End Sub

        ' video PES header (22 bytes: PTS+DTS+P-STD)

        Public Sub WriteVideoPesHeader(out As Stream,
                                       streamId As Byte,
                                       payloadLen As Integer,
                                       pts90 As Long, dts90 As Long,
                                       pstdBufferSize As Integer)
            If streamId < &HE0 OrElse streamId > &HEF Then
                Throw New ArgumentException("video PES requires streamId 0xE0..0xEF")
            End If
            Dim pesLen As Integer = 16 + payloadLen
            If pesLen > &HFFFF Then
                Throw New ArgumentOutOfRangeException(NameOf(payloadLen))
            End If

            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1) : out.WriteByte(streamId)
            out.WriteByte(CByte((pesLen >> 8) And &HFF))
            out.WriteByte(CByte(pesLen And &HFF))
            out.WriteByte(&H81)
            out.WriteByte(&HC1)
            out.WriteByte(13)
            WriteTimestamp5(out, pts90, prefix4Bits:=3)
            WriteTimestamp5(out, dts90, prefix4Bits:=1)
            out.WriteByte(&H1E)
            Dim sz As Integer = pstdBufferSize And &H1FFF
            out.WriteByte(CByte(&H40 Or &H20 Or ((sz >> 8) And &H1F)))
            out.WriteByte(CByte(sz And &HFF))
        End Sub

        ' continuation-only video PES header E
        ' mitted for every video PES that carries the tail of an access unit rather than its start
        ' only PES packets that actually begin a new AU carry PTS + DTS + P-STD,
        ' and every subsequent split of the same AU uses the short 9-byte header below
        ' (byte 6 = 0x81 marker + alignment, byte 7 = 0x00 no flags, byte 8 = 0x00 no header data)
        Public Sub WriteVideoPesHeaderContinuation(out As Stream,
                                                   streamId As Byte,
                                                   payloadLen As Integer)
            If streamId < &HE0 OrElse streamId > &HEF Then
                Throw New ArgumentException("video PES requires streamId 0xE0..0xEF")
            End If
            Dim pesLen As Integer = 3 + payloadLen  ' bytes 6, 7, 8 + payload
            If pesLen > &HFFFF Then
                Throw New ArgumentOutOfRangeException(NameOf(payloadLen))
            End If
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1) : out.WriteByte(streamId)
            out.WriteByte(CByte((pesLen >> 8) And &HFF))
            out.WriteByte(CByte(pesLen And &HFF))
            out.WriteByte(&H81)     ' byte 6: "10" marker + data_alignment_indicator
            out.WriteByte(&H0)      ' byte 7: no PTS/DTS, no extensions
            out.WriteByte(&H0)      ' byte 8: header_data_length = 0
        End Sub

        ' audio PES header (14 bytes including sub-hdr space, PTS only)

        Public Sub WriteAudioPesHeader(out As Stream,
                                       payloadLenIncludingSubHeader As Integer,
                                       pts90 As Long,
                                       pstdBufferSize As Integer)
            Dim pesLen As Integer = 11 + payloadLenIncludingSubHeader
            If pesLen > &HFFFF Then
                Throw New ArgumentOutOfRangeException(NameOf(payloadLenIncludingSubHeader))
            End If

            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_PrivateStream1)
            out.WriteByte(CByte((pesLen >> 8) And &HFF))
            out.WriteByte(CByte(pesLen And &HFF))
            out.WriteByte(&H81)
            out.WriteByte(&H81)
            out.WriteByte(8)
            WriteTimestamp5(out, pts90, prefix4Bits:=2)
            out.WriteByte(&H1E)
            Dim sz As Integer = pstdBufferSize And &H1FFF
            ' audio PES P-STD descriptor: bits 7-6 = "01" marker, bit 5 = P-STD_buffer_scale (1 = 1024-byte units, 0 = 128-byte units)
            ' PAMF use scale=1 for both audio and video PES descriptors
            ' with pstdBufferSize=20 the resulting bytes are 0x60 0x14
            out.WriteByte(CByte(&H40 Or &H20 Or ((sz >> 8) And &H1F)))
            out.WriteByte(CByte(sz And &HFF))
        End Sub

        Public Sub WriteAudioSubHeader(out As Stream,
                                       subStreamId As Byte,
                                       numFrameHeaders As Byte,
                                       firstAuPtr As UShort)
            out.WriteByte(subStreamId)
            out.WriteByte(numFrameHeaders)
            out.WriteByte(CByte((firstAuPtr >> 8) And &HFF))
            out.WriteByte(CByte(firstAuPtr And &HFF))
        End Sub

        ' padding_stream
        Public Sub WritePaddingStream(out As Stream, totalLen As Integer)
            If totalLen < 7 Then
                Throw New ArgumentException("padding_stream needs at least 7 bytes")
            End If
            Dim payloadLen As Integer = totalLen - 6
            If payloadLen > &HFFFF Then
                Throw New ArgumentException("padding > 65541; emit multiple packets")
            End If
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_PaddingStream)
            out.WriteByte(CByte((payloadLen >> 8) And &HFF))
            out.WriteByte(CByte(payloadLen And &HFF))
            For i As Integer = 0 To payloadLen - 1
                out.WriteByte(&HFF)
            Next
        End Sub

        ' program_end
        Public Sub WriteProgramEnd(out As Stream)
            out.WriteByte(0) : out.WriteByte(0) : out.WriteByte(1)
            out.WriteByte(SC_ProgramEnd)
        End Sub

    End Module

End Namespace