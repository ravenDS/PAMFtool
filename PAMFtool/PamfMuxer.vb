' PamfMuxer.vb - github.com/ravenDS/PAMFtool

Imports System.IO

Namespace PamfMux

    Public Class PamfMuxer

        Public ReadOnly Property PsMuxer As New Mpeg2PsMuxer()
        Public ReadOnly Property HeaderWriter As New PamfHeaderWriter()

        Public Property TotalDuration90 As Long = 0
        Public Property StartPts90 As Long = 90000L      ' used by header start_pts

        ' when True, header is written without an EP (entry-point) seek table
        Public Property SkipEpTable As Boolean = False

        Public Function AddAvcStream(profileIdc As Byte, levelIdc As Byte,
                                     frameMbsOnly As Byte, frameRateCode As Byte,
                                     widthPx As Integer, heightPx As Integer,
                                     Optional aspectRatioIdc As Byte = 1,
                                     Optional sarWidth As Integer = 0,
                                     Optional sarHeight As Integer = 0,
                                     Optional cropLeft As Integer = 0,
                                     Optional cropRight As Integer = 0,
                                     Optional cropTop As Integer = 0,
                                     Optional cropBottom As Integer = 0,
                                     Optional hasVideoSignalInfo As Byte = 1,
                                     Optional videoFormat As Byte = 5,
                                     Optional fullRangeFlag As Byte = 0,
                                     Optional colourPrimaries As Byte = 1,
                                     Optional transferChars As Byte = 1,
                                     Optional matrixCoeffs As Byte = 1,
                                     Optional cabacFlag As Byte = 0,
                                     Optional deblockFlag As Byte = 0) As PamfMuxStream
            Dim ps As PamfMuxStream = PsMuxer.AddStream(PamfStreamType.AVC)
            Dim widthMbs As Integer = (widthPx + 15) \ 16
            Dim heightMbs As Integer = (heightPx + 15) \ 16
            HeaderWriter.AddAvcStream(
                channel:=0, pesStreamId:=ps.PesStreamId,
                profileIdc:=profileIdc, levelIdc:=levelIdc,
                frameMbsOnlyFlag:=frameMbsOnly,
                videoSignalInfoFlag:=hasVideoSignalInfo,
                frameRateCode:=frameRateCode,
                aspectRatioIdc:=aspectRatioIdc,
                widthMbs:=widthMbs, heightMbs:=heightMbs,
                sarWidth:=sarWidth, sarHeight:=sarHeight,
                cropLeft:=cropLeft, cropRight:=cropRight,
                cropTop:=cropTop, cropBottom:=cropBottom,
                videoFormat:=videoFormat,
                videoFullRangeFlag:=fullRangeFlag,
                colourPrimaries:=colourPrimaries,
                transferCharacteristics:=transferChars,
                matrixCoefficients:=matrixCoeffs,
                cabacFlag:=cabacFlag,
                deblockingFilterFlag:=deblockFlag)
            Return ps
        End Function

        Public Function AddM2vStream(profileLevel As Byte, progressive As Byte,
                                     frameRateCode As Byte,
                                     widthPx As Integer, heightPx As Integer,
                                     Optional colourPrimaries As Byte = 1,
                                     Optional transferChars As Byte = 1,
                                     Optional matrixCoeffs As Byte = 1) As PamfMuxStream
            Dim ps As PamfMuxStream = PsMuxer.AddStream(PamfStreamType.MPEG2Video)
            Dim widthMbs As Integer = (widthPx + 15) \ 16
            Dim heightMbs As Integer = (heightPx + 15) \ 16
            HeaderWriter.AddM2vStream(
                channel:=0, pesStreamId:=ps.PesStreamId,
                profileAndLevel:=profileLevel, progressiveSeq:=progressive,
                videoSignalInfoFlag:=1, frameRateCode:=frameRateCode,
                aspectRatioIdc:=1,
                widthMbs:=widthMbs, heightMbs:=heightMbs,
                widthPx:=widthPx, heightPx:=heightPx,
                colourPrimaries:=colourPrimaries,
                transferCharacteristics:=transferChars,
                matrixCoefficients:=matrixCoeffs)
            Return ps
        End Function

        Public Function AddAtrac3plusStream(numChannels As Byte) As PamfMuxStream
            Dim ps As PamfMuxStream = PsMuxer.AddStream(PamfStreamType.ATRAC3plus)
            ' for AT3+, PAMF stream entry private_stream_id = channel index (0x00, 0x01, ...)
            HeaderWriter.AddAtrac3plusStream(
                channel:=numChannels, subStreamId:=ps.SubStreamId,
                numChannels:=numChannels, samplingFreqCode:=1)
            Return ps
        End Function

        Public Function AddAc3Stream(numChannels As Byte) As PamfMuxStream
            Dim ps As PamfMuxStream = PsMuxer.AddStream(PamfStreamType.AC3)
            ' AC-3 uses sub_stream_id 0x30 | ch
            HeaderWriter.AddAc3Stream(
                channel:=numChannels, subStreamId:=ps.SubStreamId,
                numChannels:=numChannels, samplingFreqCode:=1)
            Return ps
        End Function

        Public Function AddLpcmStream(sampleRate As Integer,
                                      numChannels As Byte,
                                      bitsPerSample As Integer) As PamfMuxStream
            Dim ps As PamfMuxStream = PsMuxer.AddStream(PamfStreamType.LPCM)
            ps.NumChannels = numChannels
            ps.BitsPerSample = CByte(bitsPerSample And &HFF)
            ' LPCM uses sub_stream_id 0x40 | ch
            HeaderWriter.AddLpcmStream(
                channel:=numChannels, subStreamId:=ps.SubStreamId,
                sampleRate:=sampleRate, numChannels:=numChannels,
                bitsPerSample:=bitsPerSample)
            Return ps
        End Function

        Public Sub QueueAu(stream As PamfMuxStream, au As AccessUnit)
            PsMuxer.QueueAu(stream, au)
            If au.Pts > TotalDuration90 Then TotalDuration90 = au.Pts
        End Sub

        Public Sub WritePamfFile(output As Stream)
            If Not output.CanSeek Then
                Throw New InvalidOperationException(
                    "PAMF output stream must be seekable (Mpeg2PsMuxer patches the header and the private_stream_2 directory tags in place).")
            End If

            Dim headerStart As Long = output.Position

            Dim placeholder(Mpeg2PsPrimitives.SectorSize - 1) As Byte
            output.Write(placeholder, 0, placeholder.Length)

            Dim psStart As Long = output.Position
            PsMuxer.WritePackedStream(output)
            Dim psBytes As Long = output.Position - psStart

            If psBytes Mod Mpeg2PsPrimitives.SectorSize <> 0 Then
                Throw New InvalidOperationException(
                    "PS payload not sector-aligned; muxer bug. Got " & psBytes)
            End If

            Dim numPacks As Long = psBytes \ Mpeg2PsPrimitives.SectorSize
            If numPacks > UInteger.MaxValue Then
                Throw New InvalidOperationException(
                    "PAMF numPacks exceeds 32-bit range (" & numPacks & "); header field would overflow.")
            End If

            ' copy EP entries from each video stream into the header
            If Not SkipEpTable Then
                For Each s In PsMuxer.Streams
                    For Each e In s.EpEntries
                        HeaderWriter.AddEpEntry(e.Pts, e.ByteOffset)
                    Next
                Next
            End If

            ' translate MuxRateBps -> PAMF header mux_rate field. same units as MPEG-2 PS pack header mux_rate (50 bytes/sec per unit)
            Dim muxRateUnits As Integer = CInt(CLng(PsMuxer.MuxRateBps) \ 8L \ 50L)

            Dim hdr As Byte() = HeaderWriter.Build(CInt(numPacks), TotalDuration90, muxRateUnits)
            If hdr.Length > Mpeg2PsPrimitives.SectorSize Then
                Throw New InvalidOperationException(
                    "PAMF header (" & hdr.Length & " bytes) exceeds the reserved sector at file start.")
            End If

            Dim endPos As Long = output.Position
            output.Position = headerStart
            output.Write(hdr, 0, hdr.Length)
            ' If the header we built happens to be smaller than the reserved sector,
            ' the tail is already zero from the placeholder write.
            output.Position = endPos
        End Sub

    End Class

End Namespace