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
        For Each a In args
            Select Case a.ToLowerInvariant()
                Case "-info", "--info", "/info" : infoOnly = True
                Case "-mux", "--mux", "/mux" : muxMode = True
                Case "-demux", "--demux", "/demux" : demuxMode = True
                Case "-noep", "--noep", "/noep" : noEp = True
                Case "-deblock", "--deblock", "/deblock" : forceDeblock = True
                Case "-nodeblock", "--nodeblock", "/nodeblock" : forceNoDeblock = True
                Case "-noatsc", "--noatsc", "/noatsc" : noAtsc = True
                Case "-h", "--help", "/?", "/h", "-?"
                    PrintUsage() : Return
                Case Else : positional.Add(a)
            End Select
        Next

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
            PamfMuxRunner.Run(positional, noEp, forceDeblock, forceNoDeblock, noAtsc)
            Return
        End If

        ' Demux / extract path
        If positional.Count < 1 Then
            PrintUsage() : Return
        End If
        PamfExtractRunner.Run(positional, wantWav, infoOnly)
    End Sub

    ' Compute "<dirname>_remux.pamf" next to the given directory.
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
        Console.WriteLine("Mux:    PAMFtool [-mux]   <inputDir>   [output.pamf] [-noep] [-noatsc] [-deblock | -nodeblock]")
        Console.WriteLine()
        Console.WriteLine("Mux parameters:")
        Console.WriteLine("  -noep       Skip writing an entry-point seek table in the header.")
        Console.WriteLine("  -noatsc     Ignore any 'atsc' RIFF chunk in .at3 inputs.")
        Console.WriteLine("  -deblock    Force codec-info deblock byte to 1 (overrides PPS-derived value).")
        Console.WriteLine("  -nodeblock  Force codec-info deblock byte to 0 (overrides PPS-derived value).")
        Console.WriteLine()
        Console.WriteLine("Additional parameters:")
        Console.WriteLine("  -info       Print info on PAMF file & streams.")
        Console.WriteLine()
    End Sub

End Module