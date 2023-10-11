Imports System.IO

Module Module1
    Sub Main()
        Dim ConsoleApplicationBase As New ApplicationServices.ConsoleApplicationBase
        Dim strZIPFile As String = Nothing
        Dim strEXEFile As String = Nothing

        If ConsoleApplicationBase.CommandLineArgs.Count = 2 Then
            For Each strCommandLineArg As String In ConsoleApplicationBase.CommandLineArgs
                If strCommandLineArg.StartsWith("--zip=", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = strCommandLineArg.Replace("--zip=", "", StringComparison.OrdinalIgnoreCase)
                ElseIf strCommandLineArg.StartsWith("--exe=", StringComparison.OrdinalIgnoreCase) Then
                    strEXEFile = strCommandLineArg.Replace("--exe=", "", StringComparison.OrdinalIgnoreCase)
                End If
            Next


            Dim commandLineArgument As String = ConsoleApplicationBase.CommandLineArgs(0).Trim
            Dim currentLocation As String = New FileInfo(Windows.Forms.Application.ExecutablePath).DirectoryName

            If commandLineArgument.StartsWith("--zip=", StringComparison.OrdinalIgnoreCase) Then
                Using zipFileStream As New FileStream(strZIPFile, FileMode.Open, FileAccess.ReadWrite)
                    Using zipFileObject As New Compression.ZipArchive(zipFileStream, Compression.ZipArchiveMode.Read)
                        For Each fileInZIP As Compression.ZipArchiveEntry In zipFileObject.Entries
                            File.Delete(Path.Combine(currentLocation, fileInZIP.Name))

                            Using fileStream As New FileStream(Path.Combine(currentLocation, fileInZIP.Name), FileMode.OpenOrCreate)
                                fileInZIP.Open.CopyTo(fileStream)
                            End Using
                        Next
                    End Using
                End Using
            End If

            Process.Start(Path.Combine(currentLocation, strEXEFile))
            Process.GetCurrentProcess.Kill()
        End If
    End Sub
End Module