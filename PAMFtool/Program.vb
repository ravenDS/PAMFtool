' PAMFtool 1.2 - by ravenDS
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
        ' -pace <Mbps> : SCR advances at this rate instead of mux_rate (48 Mbps)
        ' 0  = auto (derive from measured content bitrate)
        ' -1 = disabled (legacy behavior, SCR advances at MuxRateBps)
        Dim paceMbps As Double = 0.0
        ' -pstd <KB> : override AVC P-STD buffer size (KB)
        '              0 = use per-level default
        Dim overridePstdKb As Integer = 0
        ' -mmb <n> : override max_mean_bitrate in AVC codec_info byte 25
        ' Sony encodes 11 for 1080p L4.1 CABAC and 5 for 720p L3.1 CAVLC, some games may inspect it
        Dim overrideMmb As Integer = -1
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
                Case "-pace", "--pace", "/pace"
                    If i + 1 >= args.Length Then
                        Console.Error.WriteLine("Error: -pace requires a value in Mbps (e.g. -pace 7, or -pace auto, or -pace off).")
                        Environment.Exit(1)
                    End If
                    Dim v As String = args(i + 1).ToLowerInvariant()
                    Select Case v
                        Case "auto" : paceMbps = 0.0
                        Case "off", "none", "disable" : paceMbps = -1.0
                        Case Else
                            If Not Double.TryParse(v, System.Globalization.NumberStyles.Float,
                                                   System.Globalization.CultureInfo.InvariantCulture, paceMbps) _
                               OrElse paceMbps <= 0 Then
                                Console.Error.WriteLine("Error: -pace value must be a positive number of Mbps, 'auto', or 'off'.")
                                Environment.Exit(1)
                            End If
                    End Select
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
                              overridePstdKb, overrideMmb)
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
        Console.WriteLine("PAMFtool v1.2 - github.com/ravenDS/PAMFtool")
        Console.WriteLine("PlayStation Advanced Movie Format (PAMF) Muxer/Demuxer")
        Console.WriteLine()
        Console.WriteLine("Demux:  PAMFtool [-demux] <input.pamf> [outDir] [-info]")
        Console.WriteLine("Mux:    PAMFtool [-mux]   <inputDir>   [output.pamf] [-noep] [-noatsc] [-deblock | -nodeblock] [-pace <Mbps>] [-pstd <KB>] [-mmb <n>]")
        Console.WriteLine()
        Console.WriteLine("Mux parameters:")
        Console.WriteLine("  -noep       Skip writing an entry-point seek table in the header.")
        Console.WriteLine("  -noatsc     Ignore any 'atsc' RIFF chunk in .at3 inputs.")
        Console.WriteLine("  -deblock    Force codec-info deblock byte to 1 (overrides PPS-derived value).")
        Console.WriteLine("  -nodeblock  Force codec-info deblock byte to 0 (overrides PPS-derived value).")
        Console.WriteLine("  -pace <v>   Set SCR pacing rate.")
        Console.WriteLine("                <v> = <Mbps>  Fixed rate")
        Console.WriteLine("                <v> = auto    Compute from measured content bitrate (default)")
        Console.WriteLine("                <v> = off     Advance SCR at mux_rate (legacy front-loaded delivery)")
        Console.WriteLine("  -pstd <KB>  Override AVC P-STD buffer size in KB (1..8191).")
        Console.WriteLine("              Default is per-level (1505 for L3.1, 3703 for L4.1, etc).")
        Console.WriteLine("  -mmb <n>    Override max_mean_bitrate byte in AVC codec_info (0..255).")
        Console.WriteLine("              Sony encodes 11 for 1080p L4.1 CABAC, 5 for 720p L3.1 CAVLC.")
        Console.WriteLine()
        Console.WriteLine("Additional parameters:")
        Console.WriteLine("  -info       Print info on PAMF file & streams.")
        Console.WriteLine()
    End Sub

End Module