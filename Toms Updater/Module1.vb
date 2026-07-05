Imports System.IO
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports System.Text.RegularExpressions

Module Module1
    Private Const strVersionString As String = "1.76"
    Private Const strMessageBoxTitleText As String = "Tom's Updater"
    Private Const strBaseURL As String = "https://www.toms-world.org/download/"
    Private Const byteRoundFileSizes As Short = 2
    Public strEXEPath As String = Process.GetCurrentProcess.MainModule.FileName

    Private Sub RunNGEN(extractedFiles As Specialized.StringCollection)
        For Each strFileName As String In extractedFiles
            ColoredConsoleLineWriter("INFO:")
            Console.Write($" Installing new .NET Cached Compiled Assembly for {strFileName}...")

            Dim psi As New ProcessStartInfo With {
                .UseShellExecute = True,
                .FileName = Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe"),
                .Arguments = $"install ""{strFileName}""",
                .WindowStyle = ProcessWindowStyle.Hidden
            }

            Dim proc As Process = Process.Start(psi)
            proc.WaitForExit()

            Console.WriteLine(" Done.")
        Next
    End Sub

    ' This function writes a string to the console in a specified color. The default color is green.
    Private Sub ColoredConsoleLineWriter(strStringToWriteToTheConsole As String, Optional color As ConsoleColor = ConsoleColor.Green)
        Console.ForegroundColor = color
        Console.Write(strStringToWriteToTheConsole)
        Console.ResetColor()
    End Sub

    ' This function gets the SHA512 hash of a file at the specified path. If the file doesn't exist then we return an empty string.
    Private Function GetFileHash(strFilePath As String) As String
        If File.Exists(strFilePath) Then
            ' This opens the file at the specified path and gets the SHA512 hash of the file. We use a
            ' Using block to ensure that the file is closed after we're done with it.
            Using fileStream As FileStream = File.OpenRead(strFilePath)
                Return GetFileHash(fileStream) ' This calls the GetFileHash function that takes a Stream as a parameter to get the SHA512 hash of the file.
            End Using
        Else
            Return String.Empty ' Return an empty string if the file doesn't exist.
        End If
    End Function

    ' This function gets the SHA512 hash of a file from a stream. We use this function to get
    ' the hash of the file in the ZIP file and compare it to the hash of the local file.
    Private Function GetFileHash(fileStream As Stream) As String
        Using sha512Engine As New Security.Cryptography.SHA512CryptoServiceProvider()
            Dim hashBytes As Byte() = sha512Engine.ComputeHash(fileStream)
            Return BitConverter.ToString(hashBytes).ToLower().Replace("-", "")
        End Using
    End Function

    ' This function parses the command line arguments passed to the program and returns a dictionary of key-value pairs.
    ' The keys are the argument names and the values are the argument values. If an argument is a boolean flag
    ' (i.e., it doesn't have a value), then the value is set to True.
    Private Function ParseArguments(args As ObjectModel.ReadOnlyCollection(Of String)) As Dictionary(Of String, Object)
        ' This dictionary will hold the parsed arguments. We use StringComparer.OrdinalIgnoreCase to make the keys case-insensitive.
        Dim parsedArguments As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Dim strValue As String

        ' This loops through each argument passed to the program and parses it into a key-value pair.
        ' If the argument is a boolean flag, then the value is set to True.
        For Each strArgument As String In args
            ' This checks to see if the argument starts with "--". If it does then we parse it into a key-value pair.
            If strArgument.StartsWith("--") Then
                ' This splits the argument into a key and a value. If the argument doesn't have a value, then the value is set to True.
                Dim splitArg As String() = strArgument.Substring(2).Split(New Char() {"="c}, 2)
                Dim key As String = splitArg(0)

                If splitArg.Length = 2 Then
                    ' Argument with a value
                    strValue = splitArg(1)
                    parsedArguments(key) = strValue
                Else
                    ' Boolean flag
                    parsedArguments(key) = True
                End If
            Else
                ' This tells the user that the argument format is unrecognized and we skip it.
                Console.WriteLine($"Unrecognized argument format: {strArgument}")
            End If
        Next

        Return parsedArguments
    End Function

    Sub Main()
        Dim strProgramTitleString As String = $"== {strMessageBoxTitleText} version {strVersionString} =="

        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine(New String("=", strProgramTitleString.Length))
        Console.WriteLine(strProgramTitleString)
        Console.WriteLine(New String("=", strProgramTitleString.Length))
        Console.ResetColor()

        Dim ConsoleApplicationBase As New ApplicationServices.ConsoleApplicationBase
        Dim strProgramCode As String
        Dim strProgramEXE, strZIPFile As String
        Dim strCurrentLocation As String = AppDomain.CurrentDomain.BaseDirectory
        Dim extractedFiles As New Specialized.StringCollection
        Dim parsedArguments As Dictionary(Of String, Object) = ParseArguments(ConsoleApplicationBase.CommandLineArgs)

        If parsedArguments.Count > 0 AndAlso parsedArguments.ContainsKey("programcode") Then
            strProgramCode = parsedArguments("programcode")

            If Not String.IsNullOrWhiteSpace(strProgramCode) Then
                strProgramCode = strProgramCode.Trim

                If strProgramCode.Equals("hasher", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Hasher.zip"
                    strProgramEXE = "Hasher.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating Hasher.")
                ElseIf strProgramCode.Equals("simpleqr", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "SimpleQR.zip"
                    strProgramEXE = "SimpleQR.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating SimpleQR.")
                ElseIf strProgramCode.Equals("startprogramewithnouac", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Start Program at Startup without UAC Prompt.zip"
                    strProgramEXE = "Start Program at Startup without UAC Prompt.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating Start Program at Startup without UAC Prompt.")
                ElseIf strProgramCode.Equals("dnsoverhttps", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "DNS Over HTTPS Well Known Servers.zip"
                    strProgramEXE = "DNS Over HTTPS Well Known Servers.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating DNS Over HTTPS Well Known Servers.")
                ElseIf strProgramCode.Equals("freesyslog", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Free SysLog.zip"
                    strProgramEXE = "Free SysLog.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Updating Free SysLog.")
                ElseIf strProgramCode.Equals("scheduledtaskscanner", StringComparison.OrdinalIgnoreCase) Then
                    strZIPFile = "Scheduled Task Scanner.zip"
                    strProgramEXE = "Scheduled Task Scanner.exe"

                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Scheduled Task Scanner.")
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
            Console.Write($" Checking to see if we can write to the current location ({strCurrentLocation.TrimEnd("\")})...")

            If Not CanIWriteThere(strCurrentLocation) Then
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

            Dim RemoteFileStats As HttpHelper.RemoteFileStats = Nothing
            Dim boolDidWeGetRemoteFileStats As Boolean = False

            Try
                httpHelper.GetRemoteFileStats(strCombinedZIPFileURL, RemoteFileStats, True)
                boolDidWeGetRemoteFileStats = True
            Catch ex As Exception
                ' Do nothing
            End Try

            ' This is the MemoryStream that'll hold the contents of the ZIP file we download from the web.
            Using memoryStreamWithContentsFromTheWeb As New MemoryStream()
                ColoredConsoleLineWriter("INFO:")

                ' This checks to see if we were able to get the remote file stats for the ZIP file we are downloading from the web.
                ' If we were then we display the file size and last modified date of the ZIP file. If we weren't then we just
                ' display that we're downloading the ZIP file.
                If boolDidWeGetRemoteFileStats AndAlso RemoteFileStats.contentLength = 0 Then
                    Console.Write($" Downloading ZIP package file ""{strZIPFile}"" from ""{strCombinedZIPFileURL}"" (File Size: {FileSizeToHumanSize(RemoteFileStats.contentLength)}, Last Modified: {Date.Parse(RemoteFileStats.headers("Last-Modified")).ToLocalTime})...")
                Else
                    Console.Write($" Downloading ZIP package file ""{strZIPFile}"" from ""{strCombinedZIPFileURL}""...")
                End If

                ' This calls the DownloadFile function to download the ZIP file from the web and store it in the
                ' memoryStreamWithContentsFromTheWeb object. If it fails then we abort the update process.
                If Not httpHelper.DownloadFile(strCombinedZIPFileURL, memoryStreamWithContentsFromTheWeb, False) Then
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine(" Something went wrong, update process aborted.")
                    Console.ResetColor()
                    Exit Sub
                End If

                Console.WriteLine(" Done.")

                ' Tells the user that we're verifying the checksum of the ZIP file we downloaded from the web.
                ColoredConsoleLineWriter("INFO:")
                Console.Write($" Verifying ZIP package file ""{strZIPFile}""...")

                ' This calls the VerifyChecksum function to verify the checksum of the ZIP file we downloaded from the web.
                ' We pass the programZipFileSHA256URL, memoryStreamWithContentsFromTheWeb, and httpHelper objects to this
                ' function so that it can verify the checksum of the ZIP file. If it fails then we abort the update process.
                If Not VerifyChecksum(programZipFileSHA256URL, memoryStreamWithContentsFromTheWeb, httpHelper) Then
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine(" Something went wrong, verification failed; update process aborted.")
                    Console.ResetColor()
                    Exit Sub
                End If

                Console.WriteLine(" Done.")

                memoryStreamWithContentsFromTheWeb.Position = 0 ' Reset the position of the MemoryStream to 0 so that we can read from it again.

                ' This is where we kill the process of the program we're updating. We do this so that we can overwrite the files of the program we're updating.
                ColoredConsoleLineWriter("INFO:")
                Console.Write($" Killing Process for {strProgramEXE}...")

                ' This calls the SearchForProcessAndKillIt function to search for the process of the program we're updating and kill it.
                ' We pass the strProgramEXE variable to this function so that it knows which process to search for and kill.
                SearchForProcessAndKillIt(strProgramEXE, False)

                Console.WriteLine(" Done.") ' This tells the user that we're done killing the process of the program we're updating.

                ColoredConsoleLineWriter("INFO:")
                Console.WriteLine(" Opening ZIP file for file extraction.")

                ' This sets up the variables that we use later in this code.
                Dim strLocalFileHash, strHashOfFileInZIP, strDestinationPath As String

                Try
                    ' This is the zipFileObject that's used to extract the contents of the ZIP file we downloaded from the web.
                    Using zipFileObject As New Compression.ZipArchive(memoryStreamWithContentsFromTheWeb, Compression.ZipArchiveMode.Read)
                        ' This checks to see if the ZIP file we downloaded from the web has any files in it. If it doesn't then we assume the ZIP file is corrupt and we abort the update process.
                        If zipFileObject.Entries.Count = 0 Then
                            ' Give the user an error message that the ZIP file is corrupt and we abort the update process.
                            ColoredConsoleLineWriter("ERROR:", ConsoleColor.Red)
                            Console.WriteLine(" No files found in ZIP file, possible corrupt ZIP file.")
                        Else
                            ' This is the loop that goes through each file in the ZIP file and extracts it to the current location.
                            For Each fileInZIP As Compression.ZipArchiveEntry In zipFileObject.Entries
                                ' This checks to see if the fileInZIP object is Nothing or if the fileInZIP.Name is Nothing or empty. If it is then we skip this file and continue to the next file in the ZIP file.
                                If fileInZIP Is Nothing Then Continue For
                                If String.IsNullOrEmpty(fileInZIP.Name) Then Continue For

                                ' This is the MemoryStream that'll hold the contents of the file we are extracting from the ZIP file.
                                Using workingMemoryStream As New MemoryStream
                                    ' This is the stream that holds the contents of the file we are extracting from the ZIP file. We use
                                    ' this stream to copy the contents of the fileInZIP object to the workingMemoryStream object.
                                    Using contentsOfFileFromZIPFileStream As Stream = fileInZIP.Open()
                                        ' This copies the contents of the fileInZIP object to the workingMemoryStream object.
                                        contentsOfFileFromZIPFileStream.CopyTo(workingMemoryStream)
                                    End Using

                                    ' This sets the position of the workingMemoryStream object to 0 so that we can read from it.
                                    workingMemoryStream.Position = 0

                                    ' This is the destination path of the file we're extracting from the ZIP file. We use Path.GetFullPath to get the
                                    ' full path of the file we are extracting and combine it with the current location of the updater.exe file.
                                    strDestinationPath = Path.GetFullPath(Path.Combine(strCurrentLocation, fileInZIP.FullName))

                                    ' This checks to see if the strDestinationPath starts with the strCurrentLocation. If it doesn't then we assume that the
                                    ' fileInZIP.FullName is trying to extract a file outside of the current location and we abort the update process.
                                    If strDestinationPath.StartsWith(strCurrentLocation, StringComparison.OrdinalIgnoreCase) Then
                                        strLocalFileHash = String.Empty ' We reset the strLocalFileHash variable to an empty string so that we can use it to store the hash of the local file if it exists.

                                        ' This checks to see if the file we're extracting from the ZIP file already exists on disk. If it does then we get the hash of the local file and compare
                                        ' it to the hash of the file in the ZIP file. If they are the same then we skip extracting the file and continue to the next file in the ZIP file.
                                        If File.Exists(strDestinationPath) Then
                                            ' This gets the hash of the local file and stores it in the strLocalFileHash variable.
                                            strLocalFileHash = GetFileHash(strDestinationPath)
                                        End If

                                        ' This gets the hash of the file in the ZIP file and stores it in the strHashOfFileInZIP variable.
                                        strHashOfFileInZIP = GetFileHash(workingMemoryStream)

                                        ' This compares the hash of the local file to the hash of the file in the ZIP file. If they are
                                        ' the same then we skip extracting the file and continue to the next file in the ZIP file.
                                        If strLocalFileHash.Equals(strHashOfFileInZIP, StringComparison.OrdinalIgnoreCase) Then
                                            ' Tell the user that the local file is the same as the file in the ZIP file and there's no need to update it.
                                            ColoredConsoleLineWriter("INFO:")
                                            Console.WriteLine($" Local file ""{fileInZIP.FullName}"" is the same, there's no need to update it.")
                                        Else
                                            ' This sets the position of the workingMemoryStream object to 0 so that we can read from it again.
                                            workingMemoryStream.Position = 0

                                            ' This tells the user that we're extracting the file from the ZIP file and writing it to disk.
                                            ColoredConsoleLineWriter("INFO:")
                                            Console.Write($" Extracting and writing file ""{fileInZIP.FullName}"" to ""{strDestinationPath}""...")

                                            ' This checks to see if the file we're extracting from the ZIP file is an .exe or .dll file. If
                                            ' it is then we add it to the extractedFiles collection so that we can run NGEN on it later.
                                            If fileInZIP.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) OrElse fileInZIP.Name.EndsWith(".dll", StringComparison.OrdinalIgnoreCase) Then
                                                extractedFiles.Add(strDestinationPath)
                                            End If

                                            ' This calls the WriteExtractedFileToDisk function to write the contents of the workingMemoryStream
                                            ' object to disk. If it fails then we abort the update process. The workingMemoryStream object
                                            ' contains the contents of the file we're extracting from the ZIP file.
                                            If Not WriteExtractedFileToDisk(strCurrentLocation, fileInZIP, workingMemoryStream) Then Exit Sub
                                        End If
                                    Else
                                        ' This tells the user that the fileInZIP.FullName is trying to extract a
                                        ' file outside of the current location and we abort the update process.
                                        ColoredConsoleLineWriter("ERROR:")
                                        Console.WriteLine($" ZIP entry ""{fileInZIP.FullName}"" points outside the extraction directory.")
                                        Exit Sub
                                    End If
                                End Using
                            Next
                        End If
                    End Using

                    ' This tells the user that we're closing the ZIP file and we're done extracting files from it.
                    ColoredConsoleLineWriter("INFO:")
                    Console.WriteLine(" Closing ZIP file.")
                Catch ex As InvalidDataException
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

    ' This function is the meat and potatoes of the update process. It takes the contents of the workingMemoryStream object, which
    ' contains the contents of the file we're extracting from the ZIP file, and writes it to disk at the specified path. If it
    ' fails then we abort the update process. We pass the strCurrentLocation, fileInZIP, and workingMemoryStream objects to this
    ' function so that we can use them to write the file to disk.
    Private Function WriteExtractedFileToDisk(strCurrentLocation As String, ByRef fileInZIP As Compression.ZipArchiveEntry, ByRef zipFileMemoryStream As MemoryStream) As Boolean
        Try
            ' This is the destination path of the file we're extracting from the ZIP file. We use Path.GetDirectoryName to get the
            ' directory of the file we are extracting and combine it with the current location of the updater.exe file.
            Dim strDestinationPath As String = Path.GetDirectoryName(Path.Combine(strCurrentLocation, fileInZIP.FullName))

            ' This checks to see if the strDestinationPath exists. If it doesn't then we create it.
            If Not Directory.Exists(strDestinationPath) Then
                Directory.CreateDirectory(strDestinationPath) ' Creates the directory and any subdirectories in the specified path.
            End If

            ' This writes the contents of the zipFileMemoryStream to disk at the specified path. We use
            ' FileMode.OpenOrCreate to open the file if it exists or create it if it doesn't.
            Using fileStream As New FileStream(Path.Combine(strCurrentLocation, fileInZIP.FullName), FileMode.OpenOrCreate)
                fileStream.SetLength(0) ' Truncate the file to zero length before writing new contents.
                zipFileMemoryStream.CopyTo(fileStream) ' Copies the contents of the zipFileMemoryStream to the fileStream, which writes it to disk.
            End Using

            Console.WriteLine(" Done.") ' This tells the user that we're done writing the file to disk.

            Return True ' Return True to indicate that the file was successfully written to disk.
        Catch ex As IOException
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine(" Failed. An IOException occurred.")
            Console.ResetColor()

            ColoredConsoleLineWriter("ERROR:", ConsoleColor.Red)
            Console.Write(" An IOException occurred while extracting files from ZIP file. Update process aborted.")

            Threading.Thread.Sleep(TimeSpan.FromSeconds(5).TotalMilliseconds)

            Return False
        End Try
    End Function

    ' This function checks to see if a process with the specified PID exists. If it does then we return True and set the processObject
    ' to the process with the specified PID. If it doesn't then we return False and set the processObject to Nothing.
    Private Function DoesProcessIDExist(PID As Integer, ByRef processObject As Process) As Boolean
        Try
            processObject = Process.GetProcessById(PID)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' This function kills a process with the specified PID. If the process doesn't exist then we do nothing. If the process does exist then we kill it.
    Private Sub KillProcess(processID As Integer)
        Dim processObject As Process = Nothing ' This is the process object that we'll use to kill the process with the specified PID.

        ' This checks to see if a process with the specified PID exists.
        If DoesProcessIDExist(processID, processObject) Then
            Try
                processObject.Kill() ' Kills the process with the specified PID.
                processObject.WaitForExit(5000) ' Blocks until actually exited, or 5s safety timeout
            Catch ex As Exception
                ' Wow, it seems that even with double-checking if a process exists by it's PID number things can still go wrong.
                ' So this Try-Catch block is here to trap any possible errors when trying to kill a process by it's PID number.
            End Try
        End If
    End Sub

    ' This function searches for a process with the specified file name and kills it. If the process doesn't exist then we do nothing. If the process does exist then we kill it.
    Private Sub SearchForProcessAndKillIt(strFileName As String, boolFullFilePathPassed As Boolean)
        Dim processExecutablePath As String ' This is the full path of the executable file for the process we're checking. We use this to compare to the strFileName passed to this function.

        ' This loops through all the processes running on the system and checks to see if any of them have the same executable file name as the strFileName passed to this function.
        For Each process As Process In Process.GetProcesses()
            ' This gets the full path of the executable file for the process we're checking. We use this to compare to the strFileName passed to this function.
            processExecutablePath = GetProcessExecutablePath(process.Id)

            ' This checks to see if the processExecutablePath is Nothing or empty. If it is then we skip this process and continue to the next process.
            If Not String.IsNullOrWhiteSpace(processExecutablePath) Then
                ' This checks to see if the strFileName passed to this function is the same as the processExecutablePath. If it is then we kill the process.
                Try
                    ' This checks to see if the boolFullFilePathPassed is True. If it is then we use the full path of the
                    ' processExecutablePath to compare to the strFileName passed to this function. If it's False then we
                    ' use just the file name of the processExecutablePath to compare to the strFileName passed to this function.
                    processExecutablePath = If(boolFullFilePathPassed, New IO.FileInfo(processExecutablePath).FullName, New IO.FileInfo(processExecutablePath).Name)
                    If strFileName.Equals(processExecutablePath, StringComparison.OrdinalIgnoreCase) Then KillProcess(process.Id)
                Catch ex As ArgumentException
                End Try
            End If
        Next
    End Sub

    ' This function gets the full path of the executable file for the process with the specified PID. If the process doesn't exist then we return Nothing.
    Private Function GetProcessExecutablePath(processID As Integer) As String
        Dim processHandle As IntPtr = IntPtr.Zero ' This is the handle to the process with the specified PID. We use this to get the full path of the executable file for the process.

        Try
            ' This opens a handle to the process with the specified PID. We use this handle to get the full path of the executable file for the process.
            processHandle = NativeMethod.NativeMethods.OpenProcess(NativeMethod.ProcessAccessFlags.PROCESS_QUERY_LIMITED_INFORMATION, False, processID)

            ' This checks to see if the processHandle is IntPtr.Zero. If it is then we return Nothing.
            If processHandle.Equals(IntPtr.Zero) Then Return Nothing

            ' This is the memory buffer that we'll use to store the full path of the executable file for the process.
            ' We use a StringBuilder because it's more efficient than using a string.
            Dim memoryBuffer As New Text.StringBuilder(1024)

            ' This calls the QueryFullProcessImageName function to get the full path of the executable file for the process. If it fails then we return Nothing.
            If NativeMethod.NativeMethods.QueryFullProcessImageName(processHandle, 0, memoryBuffer, memoryBuffer.Capacity) Then
                Return memoryBuffer.ToString() ' This returns the full path of the executable file for the process.
            End If

            ' This returns Nothing if we couldn't get the full path of the executable file for the process.
            Return Nothing
        Catch ex As Exception
            Return Nothing ' This returns Nothing if an exception occurred while trying to get the full path of the executable file for the process.
        Finally
            ' This closes the handle to the process with the specified PID. We do this in a Finally block to ensure that we always close the handle, even if an exception occurs.
            If Not processHandle.Equals(IntPtr.Zero) Then NativeMethod.NativeMethods.CloseHandle(processHandle)
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
            Dim osVersion As Version = Environment.OSVersion.Version
            Dim intOSMajorVersion As Integer = osVersion.Major
            Dim intOSMinorVersion As Integer = osVersion.Minor
            Dim strDotNetVersion As String = $"{Environment.Version.Major}.{Environment.Version.Minor}"
            Dim strOSName As String

            If intOSMajorVersion = 6 And intOSMinorVersion = 1 Then
                strOSName = "Windows 7"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 2 Then
                strOSName = "Windows 8"
            ElseIf intOSMajorVersion = 6 And intOSMinorVersion = 3 Then
                strOSName = "Windows 8.1"
            ElseIf intOSMajorVersion = 10 Then
                strOSName = $"Windows {If(osVersion.Build >= 22000, "11", "10")}"
            Else
                strOSName = $"Windows NT {intOSMajorVersion}.{intOSMinorVersion}"
            End If

            Return $"{strOSName} {If(Environment.Is64BitOperatingSystem, "64", "32")}-bit (Microsoft .NET {strDotNetVersion})"
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

    Private Function CanIWriteThere(folderPath As String) As Boolean
        ' We make sure we get valid folder path by taking off the leading slash.
        If folderPath.EndsWith("\") Then folderPath = folderPath.Substring(0, folderPath.Length - 1)
        If String.IsNullOrEmpty(folderPath) OrElse Not Directory.Exists(folderPath) Then Return False

        Try
            Dim strOurTestFileName As String = Path.Combine(folderPath, "test" & RandomStringGenerator(10) & ".txt")

            File.Create(strOurTestFileName, 1, FileOptions.DeleteOnClose).Close()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function RandomStringGenerator(length As Integer) As String
        Dim random As New Random()
        Dim builder As New Text.StringBuilder()
        Dim ch As Char
        Dim legalCharacters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890"

        For cntr As Integer = 0 To length - 1
            ch = legalCharacters.Substring(random.Next(0, legalCharacters.Length), 1)
            builder.Append(ch)
        Next

        Return builder.ToString()
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
            result = $"{MyRoundingFunction(size / (2 ^ 60), shortRoundNumber)} EBs"
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