' PAMFtool 1.4 - by ravenDS
' github.com/ravenDS

Imports System.IO

Module Program

    Sub Main(args As String())
        Dim positional As New List(Of String)
        Dim wantWav As Boolean = False
        Dim infoOnly As Boolean = False
        Dim muxMode As Boolean = False
        Dim demuxMode As Boolean = False
        Dim noEp As Boolean = False
        Dim forceDeblock As Boolean = False
        Dim forceNoDeblock As Boolean = False
        Dim noAtsc As Boolean = False
        ' -pace <v> : accepted for compatibility and ignored; packs follow Sony's buffer-model scheduler
        Dim paceMbps As Double = Double.NaN
        ' -pstd <KB> : override AVC P-STD buffer size (KB)
        '              0 = use per-level default
        Dim overridePstdKb As Integer = 0
        ' -mmb <n> : override max_mean_bitrate in AVC codec_info byte 25
        Dim overrideMmb As Integer = -1
        ' -ps2-block <N> : minimum pictures between private_stream_2 / EP blocks
        '                  (blocks always start at random access points)
        '                  default -1 (0.4 s for AVC, every GOP for MPEG-2)
        Dim ps2FramesPerBlock As Integer = -1
        ' -muxrate <kbps> : override mux_rate (48000/24000/12000)
        '                   0 = leave at PsMuxer default (48000 kbps)
        Dim overrideMuxRateBps As Integer = 0
        ' -std-delay <ticks> : override std_delay_bound
        '                      0 = default 90000 (1.00 s)
        Dim overrideStdDelayTicks As Integer = 0
        ' -initial-scr <ticks> : SCR value at pack 0, Sony uses either 30 or 30030 ("game" style, 0.67 s pre-roll)
        '                        use 30030 for game-content files where Sony schedules audio bursts at packs 1/4/10
        Dim overrideInitialScr As Long = -1L
        ' -endsector : append a program_end_code sector after the last pack
        Dim endSector As Boolean = False
        Dim i As Integer = 0
        While i < args.Length
            Dim a As String = args(i)
            Select Case a.ToLowerInvariant()
                Case "-info", "--info", "/info" : infoOnly = True
                Case "-mux", "--mux", "/mux" : muxMode = True
                Case "-demux", "--demux", "/demux" : demuxMode = True
                Case "-noep", "--noep", "/noep" : noEp = True
                Case "-deblock", "--deblock", "/deblock" : forceDeblock = True
                Case "-nodeblock", "--nodeblock", "/nodeblock" : forceNoDeblock = True
                Case "-noatsc", "--noatsc", "/noatsc" : noAtsc = True
                Case "-endsector", "--endsector", "/endsector" : endSector = True
                Case "-pace", "--pace", "/pace"
                    If i + 1 >= args.Length Then
                        Console.Error.WriteLine("Error: -pace requires a value in Mbps (e.g. -pace 7, or -pace auto, or -pace off).")
                        Environment.Exit(1)
                    End If
                    paceMbps = 0.0     ' value ignored (see PamfMuxRunner)
                    i += 1
                Case "-pstd", "--pstd", "/pstd"
                    If i + 1 >= args.Length OrElse Not Integer.TryParse(args(i + 1), overridePstdKb) _
                       OrElse overridePstdKb <= 0 OrElse overridePstdKb > 8191 Then
                        Console.Error.WriteLine("Error: -pstd requires a positive integer in KB (1..8191).")
                        Environment.Exit(1)
                    End If
                    i += 1
                Case "-mmb", "--mmb", "/mmb"
                    If i + 1 >= args.Length OrElse Not Integer.TryParse(args(i + 1), overrideMmb) _
                       OrElse overrideMmb < 0 OrElse overrideMmb > 255 Then
                        Console.Error.WriteLine("Error: -mmb requires a value 0..255.")
                        Environment.Exit(1)
                    End If
                    i += 1
                Case "-ps2-block", "--ps2-block", "/ps2-block"
                    If i + 1 >= args.Length OrElse Not Integer.TryParse(args(i + 1), ps2FramesPerBlock) _
                       OrElse ps2FramesPerBlock < 0 OrElse ps2FramesPerBlock > PamfMux.Mpeg2PsMuxer.MaxPs2Pictures Then
                        Console.Error.WriteLine($"Error: -ps2-block requires a value 1..{PamfMux.Mpeg2PsMuxer.MaxPs2Pictures} (minimum pictures per block).")
                        Environment.Exit(1)
                    End If
                    i += 1
                Case "-muxrate", "--muxrate", "/muxrate"
                    Dim v As Integer
                    If i + 1 >= args.Length OrElse Not Integer.TryParse(args(i + 1), v) Then
                        Console.Error.WriteLine("Error: -muxrate requires a value in kbps (48000, 24000, or 12000).")
                        Environment.Exit(1)
                    End If
                    Dim bps As Integer = v * 1000
                    If Not PamfMux.Mpeg2PsMuxer.IsAllowedMuxRate(bps) Then
                        Console.Error.WriteLine("Error: -muxrate must be 48000, 24000, or 12000 kbps per Sony's PAMF Tools spec.")
                        Environment.Exit(1)
                    End If
                    overrideMuxRateBps = bps
                    i += 1
                Case "-std-delay", "--std-delay", "/std-delay"
                    Dim v As Integer
                    If i + 1 >= args.Length OrElse Not Integer.TryParse(args(i + 1), v) Then
                        Console.Error.WriteLine("Error: -std-delay requires a value in 1/90000 sec ticks (typical 90000 or 67500).")
                        Environment.Exit(1)
                    End If
                    If v <= 0 OrElse v > 3 * 90000 Then
                        Console.Error.WriteLine("Error: -std-delay out of range (must be 1..270000).")
                        Environment.Exit(1)
                    End If
                    overrideStdDelayTicks = v
                    i += 1
                Case "-initial-scr", "--initial-scr", "/initial-scr"
                    Dim v As Long
                    If i + 1 >= args.Length OrElse Not Long.TryParse(args(i + 1), v) Then
                        Console.Error.WriteLine("Error: -initial-scr requires a value in 90 kHz ticks (0 / 30000, or the old 30 / 30030).")
                        Environment.Exit(1)
                    End If
                    If v < 0L OrElse v > 90000L Then
                        Console.Error.WriteLine("Error: -initial-scr out of range (must be 0..90000).")
                        Environment.Exit(1)
                    End If
                    overrideInitialScr = v
                    i += 1
                Case "-h", "--help", "/?", "/h", "-?"
                    PrintUsage() : Return
                Case Else : positional.Add(a)
            End Select
            i += 1
        End While

        If muxMode AndAlso demuxMode Then
            Console.Error.WriteLine("Error: -mux and -demux are mutually exclusive.")
            Environment.Exit(1)
        End If
        If forceDeblock AndAlso forceNoDeblock Then
            Console.Error.WriteLine("Error: -deblock and -nodeblock are mutually exclusive.")
            Environment.Exit(1)
        End If

        ' Auto-detect mode
        If Not muxMode AndAlso Not demuxMode AndAlso positional.Count >= 1 Then
            Dim first As String = positional(0)
            If Directory.Exists(first) Then
                muxMode = True
            ElseIf File.Exists(first) Then
                demuxMode = True
            End If
        End If

        If muxMode Then
            If positional.Count = 1 AndAlso Directory.Exists(positional(0)) Then
                positional.Add(BuildAutoRemuxPath(positional(0)))
            End If
            PamfMuxRunner.Run(positional, noEp, forceDeblock, forceNoDeblock, noAtsc, paceMbps,
                              overridePstdKb, overrideMmb, ps2FramesPerBlock, overrideMuxRateBps,
                              overrideStdDelayTicks, overrideInitialScr, endSector)
            Return
        End If

        ' Demux / extract path
        If positional.Count < 1 Then
            PrintUsage() : Return
        End If
        PamfExtractRunner.Run(positional, wantWav, infoOnly)
    End Sub

    Private Function BuildAutoRemuxPath(inDir As String) As String
        Dim fullDir As String = Path.GetFullPath(
            inDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
        Dim parent As String = Path.GetDirectoryName(fullDir)
        If String.IsNullOrEmpty(parent) Then parent = "."
        Dim leaf As String = Path.GetFileName(fullDir)
        Return Path.Combine(parent, leaf & "_remux.pamf")
    End Function

    Private Sub PrintUsage()
        Console.WriteLine()
        Console.WriteLine("PAMFtool v1.4 - github.com/ravenDS/PAMFtool")
        Console.WriteLine("PlayStation Advanced Movie Format (PAMF) Muxer/Demuxer")
        Console.WriteLine()
        Console.WriteLine("Demux:  PAMFtool [-demux] <input.pamf> [outDir] [-info]")
        Console.WriteLine("Mux:    PAMFtool [-mux]   <inputDir>   [output.pamf] [-noep] [-noatsc]")
        Console.WriteLine("        [-deblock | -nodeblock] [-pstd <KB>] [-mmb <n>] [-ps2-block <N>]")
        Console.WriteLine("        [-muxrate <kbps>] [-std-delay <ticks>] [-initial-scr <ticks>] [-endsector]")
        Console.WriteLine()
        Console.WriteLine("Mux parameters:")
        Console.WriteLine("  -noep            Skip writing the entry-point (EP) seek tables in the header.")
        Console.WriteLine("  -noatsc          Ignore any 'atsc' RIFF chunk in .at3 inputs.")
        Console.WriteLine("  -deblock         Force codec-info deblock flag to 1 (default: 1 if any slice uses the filter).")
        Console.WriteLine("  -nodeblock       Force codec-info deblock flag to 0.")
        Console.WriteLine("  -muxrate <kbps>  Pack mux_rate: 48000 (default), 24000 or 12000.")
        Console.WriteLine()
        Console.WriteLine("  -pstd <KB>       Force the video P-STD buffer size in KB (1..8191).")
        Console.WriteLine("                      Default: from the AVC HRD cpb_size / MPEG-2 vbv_buffer_size.")
        Console.WriteLine("                      (level table when the SPS has no HRD).")
        Console.WriteLine()
        Console.WriteLine("  -mmb <n>         Override max_mean_bitrate byte in AVC codec_info (0..255; default: from the AU sizes).")
        Console.WriteLine("                      Usually 11 for 1080p L4.1 CABAC, 5 for 720p L3.1 CAVLC.")
        Console.WriteLine()
        Console.WriteLine("  -ps2-block <N>   Minimum pictures between ps2/EP blocks (1..492). ")
        Console.WriteLine("                      Default: 0.4 s (AVC), every GOP (MPEG-2).")
        Console.WriteLine()
        Console.WriteLine("  -std-delay <tk>  std_delay_bound (u32 at header 0x66) and scheduler lead limit.")
        Console.WriteLine("                      Default: 90000 (1.00 s). Sony's muxer also offers 67500 (0.75 s) and 45000 (0.50 s).")
        Console.WriteLine()
        Console.WriteLine("  -initial-scr <tk> SCR origin: pack 0 SCR = origin + whole packs. ")
        Console.WriteLine("                      Default: 0 for AVC High profile, 30000 otherwise (30 / 30030 accepted as 0 / 30000).")
        Console.WriteLine()
        Console.WriteLine("  -endsector       Append a 2048-byte sector with MPEG_program_end_code after the last pack")
        Console.WriteLine("                      Default: off.")
        Console.WriteLine()
        Console.WriteLine("Additional parameters:")
        Console.WriteLine("  -info       Print info on PAMF file & streams.")
        Console.WriteLine()
    End Sub

End Module