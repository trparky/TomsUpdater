Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text.RegularExpressions

Module Module1
    Private Const strVersionString As String = "1.2"
    Private Const strMessageBoxTitleText As String = "Tom's Updater"
    Private Const strBaseURL As String = "https://www.toms-world.org/download/"

    Private Sub RunNGEN(strFileName As String)
        Console.ForegroundColor = ConsoleColor.Green
        Console.Write("INFO:")
        Console.ResetColor()
        Console.Write(" Removing old .NET Cached Compiled Assembly...")

        Dim psi As New ProcessStartInfo With {
            .UseShellExecute = True,
            .FileName = Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe"),
            .Arguments = $"uninstall ""{strFileName}""",
            .WindowStyle = ProcessWindowStyle.Hidden
        }

        Dim proc As Process = Process.Start(psi)
        proc.WaitForExit()

        Console.WriteLine(" Done.")

        Console.ForegroundColor = ConsoleColor.Green
        Console.Write("INFO:")
        Console.ResetColor()
        Console.Write(" Installing new .NET Cached Compiled Assembly...")

        psi = New ProcessStartInfo With {
            .UseShellExecute = True,
            .FileName = Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe"),
            .Arguments = $"install ""{strFileName}""",
            .WindowStyle = ProcessWindowStyle.Hidden
        }

        proc = Process.Start(psi)
        proc.WaitForExit()

        Console.WriteLine(" Done.")
        ' ==============================
        ' == Ends code that runs NGEN ==
        ' ==============================
    End Sub

    Private Sub ColoredConsoleLineWriter(strStringToWriteToTheConsole As String, Optional color As ConsoleColor = ConsoleColor.Green)
        Console.ForegroundColor = color
        Console.Write(strStringToWriteToTheConsole)
        Console.ResetColor()
    End Sub

    Sub Main()
        Dim strProgramTitleString As String = $"== Starting {strMessageBoxTitleText} version {strVersionString} =="

        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine(New String("=", strProgramTitleString.Length))
        Console.WriteLine(strProgramTitleString)
        Console.WriteLine(New String("=", strProgramTitleString.Length))
        Console.ResetColor()

        Dim ConsoleApplicationBase As New ApplicationServices.ConsoleApplicationBase
        Dim strProgramCode As String = Nothing
        Dim strProgramEXE As String = Nothing
        Dim strZIPFile As String = Nothing
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

                    MsgBox("Invalid Program Code!", MsgBoxStyle.Critical, strMessageBoxTitleText)
                    Exit Sub
                End If
            End If

            ColoredConsoleLineWriter("INFO:")
            Console.Write($" Checking to see if we can write to the current location...")

            If Not CheckFolderPermissionsByACLs(New FileInfo(Windows.Forms.Application.ExecutablePath).DirectoryName) Then
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

            ColoredConsoleLineWriter("INFO:")
            Console.Write($" Killing Process for {strProgramEXE}...")

            SearchForProcessAndKillIt(strProgramEXE, False)

            Console.WriteLine(" Done.")

            Using memoryStream As New MemoryStream()
                ColoredConsoleLineWriter("INFO:")
                Console.Write($" Downloading ZIP package file ""{strZIPFile}"" from ""{strCombinedZIPFileURL}""...")

                If Not httpHelper.DownloadFile(strCombinedZIPFileURL, memoryStream, False) Then
                    MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, strMessageBoxTitleText)
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine(" Something went wrong, update process aborted.")
                    Console.ResetColor()
                    Exit Sub
                End If

                Console.WriteLine(" Done.")

                ColoredConsoleLineWriter("INFO:")
                Console.Write($" Verifying ZIP package file ""{strZIPFile}""...")

                If Not VerifyChecksum(programZipFileSHA256URL, memoryStream, httpHelper, True) Then
                    MsgBox("There was an error while downloading required files.", MsgBoxStyle.Critical, strMessageBoxTitleText)
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine(" Something went wrong, verification failed; update process aborted.")
                    Console.ResetColor()
                    Exit Sub
                End If

                Console.WriteLine(" Done.")

                memoryStream.Position = 0

                Using zipFileObject As New Compression.ZipArchive(memoryStream, Compression.ZipArchiveMode.Read)
                    For Each fileInZIP As Compression.ZipArchiveEntry In zipFileObject.Entries
                        ColoredConsoleLineWriter("INFO:")
                        Console.Write($" Extracting and writing file ""{fileInZIP.Name}""...")

                        Using fileStream As New FileStream(Path.Combine(currentLocation, fileInZIP.Name), FileMode.OpenOrCreate)
                            fileStream.SetLength(0)
                            fileInZIP.Open.CopyTo(fileStream)
                        End Using

                        Console.WriteLine(" Done.")
                    Next
                End Using
            End Using

            If AreWeAnAdministrator() Then RunNGEN(strProgramEXE)

            Console.ForegroundColor = ConsoleColor.Green
            Console.WriteLine("Update process complete.")

            Process.Start(Path.Combine(currentLocation, strProgramEXE))

            Console.WriteLine("Starting new instance of updated program.")
            Console.WriteLine("You may now close this console window.")
            Console.ResetColor()
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
End Module