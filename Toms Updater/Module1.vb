Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text.RegularExpressions

Module Module1
    Private Const strVersionString As String = "1.62"
    Private Const strMessageBoxTitleText As String = "Tom's Updater"
    Private Const strBaseURL As String = "https://www.toms-world.org/download/"
    Private Const byteRoundFileSizes As Short = 2
    Public strEXEPath As String = Process.GetCurrentProcess.MainModule.FileName

    Private Sub RunNGEN(extractedFiles As Specialized.StringCollection)
        For Each strFileName As String In extractedFiles
            ColoredConsoleLineWriter("INFO:")
            Console.Write($" Removing old .NET Cached Compiled Assembly for {strFileName}...")

            Dim psi As New ProcessStartInfo With {
                .UseShellExecute = True,
                .FileName = Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe"),
                .Arguments = $"uninstall ""{strFileName}""",
                .WindowStyle = ProcessWindowStyle.Hidden
            }

            Dim proc As Process = Process.Start(psi)
            proc.WaitForExit()

            Console.WriteLine(" Done.")

            ColoredConsoleLineWriter("INFO:")
            Console.Write($" Installing new .NET Cached Compiled Assembly for {strFileName}...")

            psi = New ProcessStartInfo With {
                .UseShellExecute = True,
                .FileName = Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe"),
                .Arguments = $"install ""{strFileName}""",
                .WindowStyle = ProcessWindowStyle.Hidden
            }

            proc = Process.Start(psi)
            proc.WaitForExit()

            Console.WriteLine(" Done.")
        Next
    End Sub

    Private Sub ColoredConsoleLineWriter(strStringToWriteToTheConsole As String, Optional color As ConsoleColor = ConsoleColor.Green)
        Console.ForegroundColor = color
        Console.Write(strStringToWriteToTheConsole)
        Console.ResetColor()
    End Sub

    Private Function GetFileHash(strFilePath As String) As String
        If File.Exists(strFilePath) Then
            Using fileStream As FileStream = File.OpenRead(strFilePath)
                Return GetFileHash(fileStream)
            End Using
        Else
            Return ""
        End If
    End Function

    Private Function GetFileHash(fileStream As Stream) As String
        Using sha512Engine As New Security.Cryptography.SHA512CryptoServiceProvider()
            Dim hashBytes As Byte() = sha512Engine.ComputeHash(fileStream)
            Return BitConverter.ToString(hashBytes).ToLower().Replace("-", "")
        End Using
    End Function

    Sub Main()
        Dim strProgramTitleString As String = $"== {strMessageBoxTitleText} version {strVersionString} =="

        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine(New String("=", strProgramTitleString.Length))
        Console.WriteLine(strProgramTitleString)
        Console.WriteLine(New String("=", strProgramTitleString.Length))
        Console.ResetColor()

        Dim ConsoleApplicationBase As New ApplicationServices.ConsoleApplicationBase
        Dim strProgramCode As String = Nothing
        Dim strProgramEXE, strZIPFile As String
        Dim strCurrentLocation As String = New FileInfo(strEXEPath).DirectoryName
        Dim extractedFiles As New Specialized.StringCollection

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

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating Hasher.")
                ElseIf String.Equals(strProgramCode, "simpleqr", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "SimpleQR.zip"
                    strProgramEXE = "SimpleQR.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating SimpleQR.")
                ElseIf String.Equals(strProgramCode, "startprogramewithnouac", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Start Program at Startup without UAC Prompt.zip"
                    strProgramEXE = "Start Program at Startup without UAC Prompt.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating Start Program at Startup without UAC Prompt.")
                ElseIf String.Equals(strProgramCode, "dnsoverhttps", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "DNS Over HTTPS Well Known Servers.zip"
                    strProgramEXE = "DNS Over HTTPS Well Known Servers.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating DNS Over HTTPS Well Known Servers.")
                ElseIf String.Equals(strProgramCode, "freesyslog", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Free SysLog.zip"
                    strProgramEXE = "Free SysLog.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating Free SysLog.")
                Else
                    ColoredConsoleLineWriter("ERROR:", ConsoleColor.Red)
                    Console.WriteLine(" Invalid program code.")
                    Exit Sub
                End If
            Else
                ColoredConsoleLineWriter("ERROR:", ConsoleColor.Red)
                Console.WriteLine(" Invalid program code.")
                Exit Sub
            End If

            ColoredConsoleLineWriter("INFO:")
            Console.Write($" Checking to see if we can write to the current location ({strCurrentLocation})...")

            If Not CheckFolderPermissionsByACLs(strCurrentLocation) Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine(" No. Restarting with admin privileges.")
                Console.ResetColor()

                Dim startInfo As New ProcessStartInfo With {
                    .FileName = "updater.exe",
                    .Arguments = $"--programcode={strProgramCode}",
                    .Verb = "runas"
                }

                Process.Start(startInfo)
                Process.GetCurrentProcess.Kill()
            Else
                Console.WriteLine(" Yes. Continuing update process.")
            End If

            Dim strCombinedZIPFileURL As String = $"{strBaseURL}{strZIPFile}"
            Dim programZipFileSHA256URL = $"{strCombinedZIPFileURL}.sha2"

            Dim httpHelper As HttpHelper = CreateNewHTTPHelperObject()
            httpHelper.SetHTTPTimeout = 5

            ColoredConsoleLineWriter("INFO:")
            Console.Write($" Killing Process for {strProgramEXE}...")

            SearchForProcessAndKillIt(strProgramEXE, False)

            Console.WriteLine(" Done.")

            Dim RemoteFileStats As HttpHelper.RemoteFileStats = Nothing

            Try
                httpHelper.GetRemoteFileStats(strCombinedZIPFileURL, RemoteFileStats, True)
            Catch ex As Exception
                ' Do nothing
            End Try

            Using memoryStream As New MemoryStream()
                ColoredConsoleLineWriter("INFO:")

                If RemoteFileStats.contentLength = 0 Then
                    Console.Write($" Downloading ZIP package file ""{strZIPFile}"" from ""{strCombinedZIPFileURL}""...")
                Else
                    Console.Write($" Downloading ZIP package file ""{strZIPFile}"" from ""{strCombinedZIPFileURL}"" (File Size: {FileSizeToHumanSize(RemoteFileStats.contentLength)}, Last Modified: {Date.Parse(RemoteFileStats.headers("Last-Modified")).ToLocalTime})...")
                End If

                If Not httpHelper.DownloadFile(strCombinedZIPFileURL, memoryStream, False) Then
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine(" Something went wrong, update process aborted.")
                    Console.ResetColor()
                    Exit Sub
                End If

                Console.WriteLine(" Done.")

                ColoredConsoleLineWriter("INFO:")
                Console.Write($" Verifying ZIP package file ""{strZIPFile}""...")

                If Not VerifyChecksum(programZipFileSHA256URL, memoryStream, httpHelper) Then
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine(" Something went wrong, verification failed; update process aborted.")
                    Console.ResetColor()
                    Exit Sub
                End If

                Console.WriteLine(" Done.")

                memoryStream.Position = 0

                ColoredConsoleLineWriter("INFO:")
                Console.WriteLine(" Opening ZIP file for file extraction.")

                Dim strLocalFileHash, strHashOfFileInZIP As String

                Try
                    Using zipFileObject As New Compression.ZipArchive(memoryStream, Compression.ZipArchiveMode.Read)
                        If zipFileObject.Entries.Count = 0 Then
                            ColoredConsoleLineWriter("ERROR:", ConsoleColor.Red)
                            Console.WriteLine(" No files found in ZIP file, possible corrupt ZIP file.")
                        Else
                            For Each fileInZIP As Compression.ZipArchiveEntry In zipFileObject.Entries
                                If fileInZIP IsNot Nothing Then
                                    Using zipFileMemoryStream As New MemoryStream
                                        fileInZIP.Open.CopyTo(zipFileMemoryStream)

                                        zipFileMemoryStream.Position = 0

                                        strLocalFileHash = GetFileHash(Path.Combine(strCurrentLocation, fileInZIP.Name))
                                        strHashOfFileInZIP = GetFileHash(zipFileMemoryStream)

                                        If strLocalFileHash.Equals(strHashOfFileInZIP, StringComparison.OrdinalIgnoreCase) Then
                                            ColoredConsoleLineWriter("INFO:")
                                            Console.WriteLine($" Local file ""{fileInZIP.Name}"" is the same, there's no need to update it.")
                                        Else
                                            zipFileMemoryStream.Position = 0

                                            ColoredConsoleLineWriter("INFO:")
                                            Console.Write($" Extracting and writing file ""{fileInZIP.Name}"" to ""{Path.Combine(strCurrentLocation, fileInZIP.Name)}""...")

                                            Try
                                                extractedFiles.Add(fileInZIP.Name)

                                                Using fileStream As New FileStream(Path.Combine(strCurrentLocation, fileInZIP.Name), FileMode.OpenOrCreate)
                                                    fileStream.SetLength(0)
                                                    zipFileMemoryStream.CopyTo(fileStream)
                                                End Using

                                                Console.WriteLine(" Done.")
                                            Catch ex As IOException
                                                Console.ForegroundColor = ConsoleColor.Red
                                                Console.WriteLine(" Failed. An IOException occurred.")
                                                Console.ResetColor()

                                                memoryStream.Close()
                                                memoryStream.Dispose()

                                                ColoredConsoleLineWriter("ERROR:", ConsoleColor.Red)
                                                Console.Write(" An IOException occurred while extracting files from ZIP file. Update process aborted.")

                                                Threading.Thread.Sleep(TimeSpan.FromSeconds(5).TotalMilliseconds)
                                                Exit Sub
                                            End Try
                                        End If
                                    End Using
                                End If
                            Next
                        End If
                    End Using

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Closing ZIP file.")
                Catch ex As InvalidDataException
                    memoryStream.Close()
                    memoryStream.Dispose()

                    ColoredConsoleLineWriter("ERROR:", ConsoleColor.Red)
                    Console.Write(" An InvalidDataException occurred while extracting files from ZIP file. Update process aborted.")

                    Threading.Thread.Sleep(TimeSpan.FromSeconds(5).TotalMilliseconds)
                    Exit Sub
                End Try
            End Using

            If AreWeAnAdministrator() Then
                RunNGEN(extractedFiles)
            Else
                ColoredConsoleLineWriter("INFO:")
                Console.WriteLine(" Skipping NGEN process since we're not running as an administrator.")
            End If

            Console.ForegroundColor = ConsoleColor.Green
            Console.WriteLine("Update process complete.")

            Process.Start(Path.Combine(strCurrentLocation, strProgramEXE))

            Console.WriteLine("Starting new instance of updated program.")
            Console.WriteLine("You may now close this console window.")
            Console.ResetColor()

            Threading.Thread.Sleep(TimeSpan.FromSeconds(5).TotalMilliseconds)
        Else
            ColoredConsoleLineWriter("ERROR:", ConsoleColor.Red)
            Console.WriteLine(" Invalid command line arguments.")
            Console.WriteLine("Program must be ran with a command line argument of --programcode=.")
        End If
    End Sub

    Private Function DoesProcessIDExist(PID As Integer, ByRef processObject As Process) As Boolean
        Try
            processObject = Process.GetProcessById(PID)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub KillProcess(processID As Integer)
        KillProcessSubRoutine(processID)
        Threading.Thread.Sleep(250) ' We're going to sleep to give the system some time to kill the process.
        KillProcessSubRoutine(processID)
        Threading.Thread.Sleep(250) ' We're going to sleep (again) to give the system some time to kill the process.
    End Sub

    Private Sub KillProcessSubRoutine(processID As Integer)
        Dim processObject As Process = Nothing
        If DoesProcessIDExist(processID, processObject) Then
            Try
                processObject.Kill()
            Catch ex As Exception
                ' Wow, it seems that even with double-checking if a process exists by it's PID number things can still go wrong.
                ' So this Try-Catch block is here to trap any possible errors when trying to kill a process by it's PID number.
            End Try
        End If
    End Sub

    Private Sub SearchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
        Dim processExecutablePath As String

        For Each process As Process In Process.GetProcesses()
            processExecutablePath = GetProcessExecutablePath(process.Id)

            If Not String.IsNullOrWhiteSpace(processExecutablePath) Then
                Try
                    processExecutablePath = If(boolFullFilePathPassed, New IO.FileInfo(processExecutablePath).FullName, New IO.FileInfo(processExecutablePath).Name)
                    If strFileName.Equals(processExecutablePath, StringComparison.OrdinalIgnoreCase) Then KillProcess(process.Id)
                Catch ex As ArgumentException
                End Try
            End If
        Next
    End Sub

    Private Function GetProcessExecutablePath(processID As Integer) As String
        Try
            Dim memoryBuffer As New Text.StringBuilder(1024)
            Dim processHandle As IntPtr = NativeMethod.NativeMethods.OpenProcess(NativeMethod.ProcessAccessFlags.PROCESS_QUERY_LIMITED_INFORMATION, False, processID)

            If Not processHandle.Equals(IntPtr.Zero) Then
                Try
                    Dim memoryBufferSize As Integer = memoryBuffer.Capacity
                    If NativeMethod.NativeMethods.QueryFullProcessImageName(processHandle, 0, memoryBuffer, memoryBufferSize) Then Return memoryBuffer.ToString()
                Finally
                    NativeMethod.NativeMethods.CloseHandle(processHandle)
                End Try
            End If

            NativeMethod.NativeMethods.CloseHandle(processHandle)
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function AreWeAnAdministrator() As Boolean
        Try
            Return New WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)
        Catch ex As Exception
            Return False
        End Try
    End Function

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

    Private Function VerifyChecksum(urlOfChecksumFile As String, ByRef memStream As MemoryStream, ByRef httpHelper As HttpHelper) As Boolean
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
                        Return False
                    End If
                Else
                    ' Handles regex parsing errors.
                    Return False
                End If
            Else
                ' Handles any HTTP errors.
                Return False
            End If
        Catch ex As Exception
            ' Handles any exceptions.
            Return False
        End Try
    End Function

    Private Function SHA256ChecksumStream(ByRef stream As Stream) As String
        Using SHA256Engine As New Security.Cryptography.SHA256CryptoServiceProvider
            Return BitConverter.ToString(SHA256Engine.ComputeHash(stream)).ToLower().Replace("-", "").Trim
        End Using
    End Function

    Private Function CheckFolderPermissionsByACLs(folderPath As String) As Boolean
        Try
            Dim directoryACLs As DirectorySecurity = Directory.GetAccessControl(folderPath)
            Dim directoryAccessRights As FileSystemAccessRule

            For Each rule As AuthorizationRule In directoryACLs.GetAccessRules(True, True, GetType(SecurityIdentifier))
                If rule.IdentityReference.Value.Equals(WindowsIdentity.GetCurrent.User.Value, StringComparison.OrdinalIgnoreCase) Then
                    directoryAccessRights = DirectCast(rule, FileSystemAccessRule)

                    If directoryAccessRights.AccessControlType = AccessControlType.Allow AndAlso directoryAccessRights.FileSystemRights = (FileSystemRights.Read Or FileSystemRights.Modify Or FileSystemRights.Write Or FileSystemRights.FullControl) Then
                        Return True
                    End If
                End If
            Next

            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function FileSizeToHumanSize(size As Long, Optional roundToNearestWholeNumber As Boolean = False) As String
        Dim result As String
        Dim shortRoundNumber As Short = If(roundToNearestWholeNumber, 0, byteRoundFileSizes)

        If size <= (2 ^ 10) Then
            result = $"{size} Bytes"
        ElseIf size > (2 ^ 10) And size <= (2 ^ 20) Then
            result = $"{MyRoundingFunction(size / (2 ^ 10), shortRoundNumber)} KBs"
        ElseIf size > (2 ^ 20) And size <= (2 ^ 30) Then
            result = $"{MyRoundingFunction(size / (2 ^ 20), shortRoundNumber)} MBs"
        ElseIf size > (2 ^ 30) And size <= (2 ^ 40) Then
            result = $"{MyRoundingFunction(size / (2 ^ 30), shortRoundNumber)} GBs"
        ElseIf size > (2 ^ 40) And size <= (2 ^ 50) Then
            result = $"{MyRoundingFunction(size / (2 ^ 40), shortRoundNumber)} TBs"
        ElseIf size > (2 ^ 50) And size <= (2 ^ 60) Then
            result = $"{MyRoundingFunction(size / (2 ^ 50), shortRoundNumber)} PBs"
        ElseIf size > (2 ^ 60) And size <= (2 ^ 70) Then
            result = $"{MyRoundingFunction(size / (2 ^ 50), shortRoundNumber)} EBs"
        Else
            result = "(None)"
        End If

        Return result
    End Function

    Private Function MyRoundingFunction(value As Double, digits As Integer) As String
        If digits = 0 Then
            Return Math.Round(value, digits).ToString
        Else
            Dim strFormatString As String = "{0:0." & New String("0", digits) & "}"
            Return String.Format(strFormatString, Math.Round(value, digits))
        End If
    End Function
End Module