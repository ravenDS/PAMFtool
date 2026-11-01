' PamfMuxer.vb - github.com/ravenDS/PAMFtool

Imports System.IO

Namespace PamfMux

    Public Class PamfMuxer

        Public ReadOnly Property PsMuxer As New Mpeg2PsMuxer()
        Public ReadOnly Property HeaderWriter As New PamfHeaderWriter()

        Public Property StartPts90 As Long = 90000L      ' header start_pts
        ' header end_pts: first video PTS + total video duration (set by caller from video timing)
        ' 0 = derive from queued AUs
        Public Property EndPts90 As Long = 0

        ' std_delay_bound written into PAMF header at offset 0x66 (u32, 1/90000 sec units)
        ' used as scheduler lead limit
        ' 1.00 s (90000) unless -std-delay says otherwise
        Public Property StdDelayBoundTicks As Integer = 90000

        ' P-STD values that scheduler may raise when declared video buffer would make AUs arrive late
        Public Property AllowPstdGrowth As Boolean = True

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
            ' level-table default, caller replaces it with HRD-derived value when SPS has one
            ps.PstdBufferSize = PamfHeaderWriter.AvcPstdKbForLevel(levelIdc)
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
        End Sub

        Public Sub WritePamfFile(output As Stream)
            PsMuxer.StdDelayTicks = StdDelayBoundTicks

            ' video buffer must at least hold largest AU plus two packs, or room rule can never let it in
            For Each s In PsMuxer.Streams
                Dim maxAu As Long = 0
                For Each au In s.Aus
                    If au.Data.LongLength > maxAu Then maxAu = au.Data.LongLength
                Next
                Dim minKb As Integer = CInt((maxAu + 2L * Mpeg2PsPrimitives.SectorSize + 1023L) \ 1024L)
                If s.PstdBufferSize < minKb Then
                    If minKb > Mpeg2PsMuxer.MaxPstdKb Then
                        Throw New InvalidDataException($"stream {s.Index}: an access unit of {maxAu:N0} bytes does not fit the largest P-STD buffer ({Mpeg2PsMuxer.MaxPstdKb} KB)")
                    End If
                    Console.WriteLine($"  [warn] s{s.Index:D2}: P-STD {s.PstdBufferSize} KB is smaller than the largest AU ({maxAu:N0} B), raised to {minKb} KB")
                    s.PstdBufferSize = minKb
                End If
            Next

            ' if video AUs would be late because the declared buffer is too small, grow it and replan
            ' keep smallest buffers that give fewest late AUs
            Dim n0 As Integer = PsMuxer.Streams.Count
            Dim bestKb(n0 - 1) As Integer
            Dim bestLate As Long = Long.MaxValue
            Do
                PsMuxer.Plan()
                Dim lateNow As Long = 0
                For Each s In PsMuxer.Streams
                    lateNow += s.LateAus
                Next
                If lateNow < bestLate Then
                    bestLate = lateNow
                    For i As Integer = 0 To n0 - 1
                        bestKb(i) = PsMuxer.Streams(i).PstdBufferSize
                    Next
                End If
                Dim grew As Boolean = False
                If AllowPstdGrowth AndAlso lateNow > 0 Then
                    For Each s In PsMuxer.Streams
                        If s.IsVideo AndAlso s.LateAus > 0 AndAlso s.PstdBufferSize < Mpeg2PsMuxer.MaxPstdKb Then
                            Dim nk As Integer = Math.Min(Mpeg2PsMuxer.MaxPstdKb, CInt(Math.Ceiling(s.PstdBufferSize * 1.25)))
                            Console.WriteLine($"  [warn] s{s.Index:D2}: {s.LateAus} AUs would arrive after their DTS with P-STD {s.PstdBufferSize} KB, retrying with {nk} KB")
                            s.PstdBufferSize = nk
                            grew = True
                        End If
                    Next
                End If
                If Not grew Then Exit Do
            Loop
            Dim changed As Boolean = False
            For i As Integer = 0 To n0 - 1
                If PsMuxer.Streams(i).PstdBufferSize <> bestKb(i) Then changed = True
            Next
            If changed Then
                For i As Integer = 0 To n0 - 1
                    If PsMuxer.Streams(i).PstdBufferSize <> bestKb(i) Then
                        Console.WriteLine($"  s{i:D2}: a larger P-STD did not reduce late AUs, keeping {bestKb(i)} KB")
                    End If
                    PsMuxer.Streams(i).PstdBufferSize = bestKb(i)
                Next
                PsMuxer.Plan()
            End If
            For Each s In PsMuxer.Streams
                If s.LateAus > 0 Then
                    Console.WriteLine($"  [warn] s{s.Index:D2}: {s.LateAus} AUs arrive late (up to {s.MaxLate27 / 27000.0:F1} ms) - the mux rate cannot carry this bit rate within std_delay_bound")
                End If
            Next

            Dim n As Integer = PsMuxer.Streams.Count
            For i As Integer = 0 To n - 1
                Dim s As PamfMuxStream = PsMuxer.Streams(i)
                HeaderWriter.SetPstdKb(i, s.PstdBufferSize)
                If s.IsVideo AndAlso Not SkipEpTable Then
                    HeaderWriter.SetEpEntries(i, s.EpEntries)
                Else
                    HeaderWriter.SetEpEntries(i, Nothing)
                End If
            Next

            Dim endPts As Long = EndPts90
            If endPts <= 0 Then
                For Each s In PsMuxer.Streams
                    For Each au In s.Aus
                        If au.Pts > endPts Then endPts = au.Pts
                    Next
                Next
            End If

            ' data sectors = packs (+ the program_end sector when it is written)
            Dim dataSectors As Long = PsMuxer.PackCount + If(PsMuxer.WriteEndSector, 1L, 0L)
            If dataSectors > UInteger.MaxValue Then
                Throw New InvalidOperationException("PAMF data sector count exceeds 32 bits")
            End If
            ' header mux_rate_bound, same units as pack header mux_rate (50 bytes/sec per unit)
            Dim muxRateUnits As Integer = CInt(CLng(PsMuxer.MuxRateBps) \ 8L \ 50L)
            Dim hdr As Byte() = HeaderWriter.Build(dataSectors, StartPts90, endPts, muxRateUnits, StdDelayBoundTicks)
            output.Write(hdr, 0, hdr.Length)
            PsMuxer.WritePacks(output)
        End Sub

    End Class

End Namespace