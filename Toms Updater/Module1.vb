Imports System.IO
Imports System.Text.RegularExpressions

Module Module1
    Private strVersionString As String = "1.0"
    Private strMessageBoxTitleText As String = "Tom's Updater"

    Private Sub RunNGEN(strFileName As String)
        Console.Write("Removing old .NET Cached Compiled Assembly...")

        Dim psi As New ProcessStartInfo With {
            .UseShellExecute = True,
            .FileName = Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe"),
            .Arguments = $"uninstall {Chr(34)}{strFileName}{Chr(34)}",
            .WindowStyle = ProcessWindowStyle.Hidden
        }

        Dim proc As Process = Process.Start(psi)
        proc.WaitForExit()

        Console.WriteLine(" Done.")

        Console.Write("Installing new .NET Cached Compiled Assembly...")

        psi = New ProcessStartInfo With {
            .UseShellExecute = True,
            .FileName = IO.Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe"),
            .Arguments = $"install {Chr(34)}{strFileName}{Chr(34)}",
            .WindowStyle = ProcessWindowStyle.Hidden
        }

        proc = Process.Start(psi)
        proc.WaitForExit()

        Console.WriteLine(" Done.")
        ' ==============================
        ' == Ends code that runs NGEN ==
        ' ==============================
    End Sub

    Sub Main()
        Console.WriteLine("-----------------------")
        Console.WriteLine("Starting Tom's Updater.")
        Console.WriteLine("-----------------------")

        Dim ConsoleApplicationBase As New ApplicationServices.ConsoleApplicationBase
        Dim strProgramCode As String = Nothing
        Dim strProgramEXE As String = Nothing
        Dim strZIPFile As String = Nothing
        Dim strBaseURL As String = "www.toms-world.org/download/"
        Dim currentLocation As String = New FileInfo(Windows.Forms.Application.ExecutablePath).DirectoryName

        If ConsoleApplicationBase.CommandLineArgs.Count = 1 Then
            For Each strCommandLineArg As String In ConsoleApplicationBase.CommandLineArgs
                If strCommandLineArg.StartsWith("--programcode=", StringComparison.OrdinalIgnoreCase) Then
                    strProgramCode = strCommandLineArg.Replace("--programcode=", "", StringComparison.OrdinalIgnoreCase)
                End If
            Next

            If Not String.IsNullOrWhiteSpace(strProgramCode) Then
                strProgramCode = strProgramCode.Trim

                If String.Equals(strProgramCode, "hasher", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Hasher.zip"
                    strProgramEXE = "Hasher.exe"
                    Console.WriteLine("INFO: Updating Hasher.")
                ElseIf String.Equals(strProgramCode, "simpleqr", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "SimpleQR.zip"
                    strProgramEXE = "SimpleQR.exe"
                    Console.WriteLine("INFO: Updating SimpleQR.")
                ElseIf String.Equals(strProgramCode, "startprogramewithnouac", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Start Program at Startup without UAC Prompt.zip"
                    strProgramEXE = "Start Program at Startup without UAC Prompt.exe"
                    Console.WriteLine("INFO: Updating Start Program at Startup without UAC Prompt.")
                ElseIf String.Equals(strProgramCode, "dnsoverhttps", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "DNS Over HTTPS Well Known Servers.zip"
                    strProgramEXE = "DNS Over HTTPS Well Known Servers.exe"
                    Console.WriteLine("INFO: Updating DNS Over HTTPS Well Known Servers.")
                ElseIf String.Equals(strProgramCode, "freesyslog", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Free SysLog.zip"
                    strProgramEXE = "Free SysLog.exe"
                    Console.WriteLine("INFO: Updating Free SysLog.")
                Else
                    MsgBox("Invalid Program Code!", MsgBoxStyle.Critical, strMessageBoxTitleText)
                    Process.GetCurrentProcess.Kill()
                    Console.WriteLine("ERROR: Invalid program code.")
                    Exit Sub
                End If
            End If

            Dim strCombinedZIPFileURL As String = $"{strBaseURL}{strZIPFile}"
            Dim programZipFileSHA256URL = $"{strCombinedZIPFileURL}.sha2"

            Dim httpHelper As HttpHelper = CreateNewHTTPHelperObject()

            Using memoryStream As New MemoryStream()
                Console.Write($"INFO: Downloading ZIP package file {Chr(34)}{strZIPFile}{Chr(34)}...")
                If Not httpHelper.DownloadFile(strCombinedZIPFileURL, memoryStream, False) Then
                    MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, strMessageBoxTitleText)
                    Console.WriteLine(" Something went wrong, update process aborted.")
                    Process.GetCurrentProcess.Kill()
                    Exit Sub
                End If
                Console.WriteLine(" Done.")

                Console.Write($"INFO: Verifying ZIP package file {Chr(34)}{strZIPFile}{Chr(34)}...")
                If Not VerifyChecksum(programZipFileSHA256URL, memoryStream, httpHelper, True) Then
                    MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, strMessageBoxTitleText)
                    Console.WriteLine(" Something went wrong, verification failed; update process aborted.")
                    Process.GetCurrentProcess.Kill()
                    Exit Sub
                End If
                Console.WriteLine(" Done.")

                memoryStream.Position = 0

                Using zipFileObject As New Compression.ZipArchive(memoryStream, Compression.ZipArchiveMode.Read)
                    For Each fileInZIP As Compression.ZipArchiveEntry In zipFileObject.Entries
                        Console.Write($"INFO: Deleting file {Chr(34)}{fileInZIP.Name}{Chr(34)}...")
                        File.Delete(Path.Combine(currentLocation, fileInZIP.Name))
                        Console.WriteLine(" Done.")

                        Console.Write($"INFO: Extracting and writing file {Chr(34)}{fileInZIP.Name}{Chr(34)}...")
                        Using fileStream As New FileStream(Path.Combine(currentLocation, fileInZIP.Name), FileMode.OpenOrCreate)
                            fileInZIP.Open.CopyTo(fileStream)
                        End Using
                        Console.WriteLine(" Done.")
                    Next
                End Using
            End Using

            RunNGEN(strProgramEXE)

            Console.WriteLine("Update process complete.")

            Process.Start(Path.Combine(currentLocation, strProgramEXE))

            Console.WriteLine("Starting new instance updated program.")
            Console.WriteLine("You may now close this console window.")

            Process.GetCurrentProcess.Kill()
        End If
    End Sub

    Public Function CreateNewHTTPHelperObject() As HttpHelper
        Dim httpHelper As New HttpHelper With {
            .SetUserAgent = CreateHTTPUserAgentHeaderString(),
            .UseHTTPCompression = True,
            .SetProxyMode = True
        }
        httpHelper.AddHTTPHeader("PROGRAM_NAME", "Tom's Updater")
        httpHelper.AddHTTPHeader("PROGRAM_VERSION", strVersionString)
        httpHelper.AddHTTPHeader("OPERATING_SYSTEM", GetFullOSVersionString())
        If File.Exists("dontcount") Then httpHelper.AddHTTPCookie("dontcount", "True", "www.toms-world.org", False)

        httpHelper.SetURLPreProcessor = Function(strURLInput As String) As String
                                            Try
                                                If Not strURLInput.Trim.StartsWith("http", StringComparison.OrdinalIgnoreCase) Then
                                                    Return $"https://{strURLInput}"
                                                Else
                                                    Return strURLInput
                                                End If
                                            Catch ex As Exception
                                                Return strURLInput
                                            End Try
                                        End Function

        Return httpHelper
    End Function

    Private Function CreateHTTPUserAgentHeaderString() As String
        Dim versionInfo As String() = Windows.Forms.Application.ProductVersion.Split(".")
        Dim versionString As String = $"{versionInfo(0)}.{versionInfo(1)} Build {versionInfo(2)}"
        Return $"Tom's Updater version {versionString} on {GetFullOSVersionString()}"
    End Function

    Private Function GetFullOSVersionString() As String
        Try
            Dim intOSMajorVersion As Integer = Environment.OSVersion.Version.Major
            Dim intOSMinorVersion As Integer = Environment.OSVersion.Version.Minor
            Dim dblDOTNETVersion As Double = Double.Parse($"{Environment.Version.Major}.{Environment.Version.Minor}")
            Dim strOSName As String

            If intOSMajorVersion = 5 And intOSMinorVersion = 0 Then
                strOSName = "Windows 2000"
            ElseIf intOSMajorVersion = 5 And intOSMinorVersion = 1 Then
                strOSName = "Windows XP"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 0 Then
                strOSName = "Windows Vista"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 1 Then
                strOSName = "Windows 7"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 2 Then
                strOSName = "Windows 8"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 3 Then
                strOSName = "Windows 8.1"
            ElseIf intOSMajorVersion = 10 Then
                strOSName = "Windows 10"
            ElseIf intOSMajorVersion = 11 Then
                strOSName = "Windows 11"
            Else
                strOSName = $"Windows NT {intOSMajorVersion}.{intOSMinorVersion}"
            End If

            Return $"{strOSName} {If(Environment.Is64BitOperatingSystem, "64", "32")}-bit (Microsoft .NET {dblDOTNETVersion })"
        Catch ex As Exception
            Try
                Return $"Unknown Windows Operating System ({Environment.OSVersion.VersionString})"
            Catch ex2 As Exception
                Return "Unknown Windows Operating System"
            End Try
        End Try
    End Function

    Private Function VerifyChecksum(urlOfChecksumFile As String, ByRef memStream As MemoryStream, ByRef httpHelper As HttpHelper, boolGiveUserAnErrorMessage As Boolean) As Boolean
        Dim checksumFromWeb As String = Nothing
        memStream.Position = 0

        Try
            If httpHelper.GetWebData(urlOfChecksumFile, checksumFromWeb) Then
                Dim regexObject As New Regex("([a-zA-Z0-9]{64})")

                ' Checks to see if we have a valid SHA256 file.
                If regexObject.IsMatch(checksumFromWeb) Then
                    ' Now that we have a valid SHA256 file we need to parse out what we want.
                    checksumFromWeb = regexObject.Match(checksumFromWeb).Groups(1).Value.Trim()

                    ' Now we do the actual checksum verification by passing the name of the file to the SHA256() function
                    ' which calculates the checksum of the file on disk. We then compare it to the checksum from the web.
                    If SHA256ChecksumStream(memStream).Equals(checksumFromWeb, StringComparison.OrdinalIgnoreCase) Then
                        Return True ' OK, things are good; the file passed checksum verification so we return True.
                    Else
                        ' The checksums don't match. Oops.
                        If boolGiveUserAnErrorMessage Then
                            MsgBox("There was an error in the download, checksums don't match. Update process aborted.", MsgBoxStyle.Critical, strMessageBoxTitleText)
                        End If

                        Return False
                    End If
                Else
                    If boolGiveUserAnErrorMessage Then
                        MsgBox("Invalid SHA2 file detected. Update process aborted.", MsgBoxStyle.Critical, strMessageBoxTitleText)
                    End If

                    Return False
                End If
            Else
                If boolGiveUserAnErrorMessage Then
                    MsgBox("There was an error downloading the checksum verification file. Update process aborted.", MsgBoxStyle.Critical, strMessageBoxTitleText)
                End If

                Return False
            End If
        Catch ex As Exception
            If boolGiveUserAnErrorMessage Then
                MsgBox("There was an error downloading the checksum verification file. Update process aborted.", MsgBoxStyle.Critical, strMessageBoxTitleText)
            End If

            Return False
        End Try
    End Function

    Private Function SHA256ChecksumStream(ByRef stream As Stream) As String
        Using SHA256Engine As New Security.Cryptography.SHA256CryptoServiceProvider
            Return BitConverter.ToString(SHA256Engine.ComputeHash(stream)).ToLower().Replace("-", "").Trim
        End Using
    End Function
End Module