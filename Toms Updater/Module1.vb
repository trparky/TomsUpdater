Imports System.Net
Imports System.IO
Imports System.Management
Imports System.IO.Compression

Module Module1
    Private numberOfTimesToLoopWhileWaiting As Short = 0
    Private Const strQuote As String = ControlChars.Quote
    'Private boolWinXP As Boolean = False ' Default value is False.
    Private boolDidWeInformTheUserThatTheProcessIsStillRunning As Boolean = False

'    Private Const Microsoft_Win32_TaskScheduler_dll_URL As String = "http://www.toms-world.org/download/Microsoft.Win32.TaskScheduler.dll"
'    Private Const Microsoft_Win32_TaskScheduler_dll_checksum As String = "515D0BE690D2010BF76E9E798332B6B2C7325765"

'    Private Const SmoothProgressBar_dll_URL As String = "http://www.toms-world.org/download/SmoothProgressBar.dll"
'    Private Const SmoothProgressBar_dll_checksum As String = "490E436F2E9C98AD4475EDACA068370E3646773F"

'    Private Const ThoughtWorks_QRCode_dll_URL As String = "http://www.toms-world.org/download/ThoughtWorks.QRCode.dll"
'    Private Const ThoughtWorks_QRCode_dll_checksum As String = "16FEDFC35F846CFE73BDCC34EFDD5660B4AB461B"

'    Private Const ICSharpCode_SharpZipLib_dll_URL As String = "http://www.toms-world.org/download/ICSharpCode.SharpZipLib.dll"
'    Private Const ICSharpCode_SharpZipLib_dll_checksum As String = "7A9DF9C25D49690B6A3C451607D311A866B131F4"

    Private Function doesPIDExist(PID As Integer) As Boolean
        Dim searcher As New ManagementObjectSearcher("root\CIMV2", String.Format("SELECT * FROM Win32_Process WHERE ProcessId={0}", PID))

        If searcher.Get.Count = 0 Then
            searcher.Dispose()
            Return False
        Else
            searcher.Dispose()
            Return True
        End If
    End Function

    'Function WMIKillProcess(PID As Integer) As Boolean
    '    Try
    '        Dim classInstance As New ManagementObject("root\CIMV2", String.Format("Win32_Process.Handle='{0}'", PID), Nothing)

    '        ' Obtain [in] parameters for the method
    '        Dim inParams As ManagementBaseObject = classInstance.GetMethodParameters("Terminate")

    '        ' Add the input parameters.

    '        ' Execute the method and obtain the return values.
    '        Dim outParams As ManagementBaseObject = classInstance.InvokeMethod("Terminate", inParams, Nothing)

    '        If outParams("ReturnValue") = 0 Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Catch err As ManagementException
    '        Return False
    '    End Try
    'End Function

    Private Sub killProcess(PID As Integer)
        Try
            Dim processDetail As Process

            Console.Write("Killing PID {0}...", PID)

            processDetail = Process.GetProcessById(PID)
            processDetail.Kill()

            Threading.Thread.Sleep(500)

            If doesPIDExist(PID) Then
                Console.WriteLine(" Process still running.  Attempting to kill process again.")
                killProcess(PID)
            Else
                Console.WriteLine(" Process Killed.")
            End If
        Catch ex As Exception
            ' Does nothing
        End Try
    End Sub

    Private Sub searchForProcess(fileName As String)
        Dim fullFileName As String = New FileInfo(fileName).FullName
        'Dim PID As Integer

        Console.WriteLine("Killing all processes that belong to parent executable file. Please Wait.")
        'Console.WriteLine(String.Format("SELECT * FROM Win32_Process WHERE ExecutablePath = '{0}'", fullFileName.Replace("\", "\\")))

        Dim searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_Process")

        Try
            For Each queryObj As ManagementObject In searcher.Get()
                If queryObj("ExecutablePath") IsNot Nothing Then
                    If queryObj("ExecutablePath") = fullFileName Then
                        killProcess(Integer.Parse(queryObj("ProcessId").ToString))
                    End If
                End If
            Next

            Console.WriteLine("All processes killed... Update process can continue.")
        Catch err As ManagementException

        End Try


        'Dim searcher As New ManagementObjectSearcher("root\CIMV2", String.Format("SELECT * FROM Win32_Process WHERE name = '{0}'", fileName))
        'Dim exePath As String = ""
        'Dim processDetail As Process

        'Try
        '    If searcher.Get().Count = 0 = False Then
        '        For Each queryObj As ManagementObject In searcher.Get()
        '            Try
        '                If queryObj("ExecutablePath") Is Nothing Then
        '                    exePath = ""
        '                Else
        '                    exePath = queryObj("ExecutablePath")
        '                End If

        '                Debug.WriteLine("exePath = " & exePath)

        '                If exePath.ToLower.Contains(fileName.ToLower) = True Then
        '                    If numberOfTimesToLoopWhileWaiting = 10 Then
        '                        Console.Write("Attempting to kill process ID {0}, please wait...", queryObj("ProcessId").ToString)

        '                        Try
        '                            processDetail = Process.GetProcessById(Integer.Parse(queryObj("ProcessId").ToString))
        '                            processDetail.Kill()
        '                        Catch ex As Exception
        '                            Exit Sub
        '                        End Try

        '                        numberOfTimesToLoopWhileWaiting = 0
        '                        boolDidWeInformTheUserThatTheProcessIsStillRunning = False

        '                        Console.WriteLine(" Process kill command sent.")

        '                        boolDidWeInformTheUserThatTheProcessIsStillRunning = False
        '                        Threading.Thread.Sleep(1000)
        '                        searchForProcess(fileName)
        '                        Exit Sub
        '                    End If

        '                    numberOfTimesToLoopWhileWaiting += 1

        '                    If boolDidWeInformTheUserThatTheProcessIsStillRunning = False Then
        '                        Console.WriteLine("Process still running (PID: {0}), waiting for process to quit.", queryObj("ProcessId").ToString)
        '                        boolDidWeInformTheUserThatTheProcessIsStillRunning = True
        '                    End If

        '                    Threading.Thread.Sleep(1000)
        '                    searchForProcess(fileName)
        '                End If
        '            Catch err As ManagementException
        '                ' Does nothing
        '            End Try
        '        Next
        '    End If
        'Catch ex As Exception
        '    Console.WriteLine("A bad thing happened, aborting update process. -- " & ex.Message & " -- " & ex.StackTrace)
        '    Threading.Thread.Sleep(5000)
        '    End
        'End Try


        'For Each processDetail In Process.GetProcesses
        '    Try
        '        If processDetail.MainModule.FileName.Contains(fileName) Then
        '            If numberOfTimesToLoopWhileWaiting = 10 Then
        '                Console.Write("Attempting to kill process, please wait...")
        '                processDetail.Kill()
        '                Console.WriteLine(" Process kill command sent.")
        '                boolDidWeInformTheUserThatTheProcessIsStillRunning = False
        '                Threading.Thread.Sleep(1000)
        '                searchForProcess(fileName)
        '            End If
        '            numberOfTimesToLoopWhileWaiting += 1
        '            If boolDidWeInformTheUserThatTheProcessIsStillRunning = False Then
        '                Console.WriteLine("Process still running, waiting for process to quit.")
        '                boolDidWeInformTheUserThatTheProcessIsStillRunning = True
        '            End If
        '            Threading.Thread.Sleep(1000)
        '            searchForProcess(fileName)
        '        End If
        '    Catch ex As Exception
        '        Console.WriteLine("A bad thing happened, aborting update process.")
        '        Threading.Thread.Sleep(5000)
        '        End
        '    End Try
        'Next
    End Sub

    'Function calculateSHA1Value(file As String) As String
    '    Dim sha1Obj As New System.Security.Cryptography.SHA1CryptoServiceProvider
    '    Dim stream As New IO.FileStream(file, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
    '    Dim bytesToHash() As Byte = sha1Obj.ComputeHash(stream)
    '    stream.Close()
    '    stream.Dispose()
    '    Return BitConverter.ToString(bytesToHash).ToUpper().Replace("-", "")
    'End Function

    'Private attemptedDownloads As Short

    'Sub downloadRequiredDLLFile(fileName As String, downloadURL As String, properCheckSum As String, Optional resetDownloadCount As Boolean = True)
    '    Dim webClient As New System.Net.WebClient

    '    If resetDownloadCount = True Then
    '        attemptedDownloads = 0
    '    End If

    '    Try
    '        If IO.File.Exists(fileName) = True Then
    '            Console.WriteLine(fileName & " found.")
    '            Console.Write("Checking file integrity...")

    '            If calculateSHA1Value(fileName) = properCheckSum Then
    '                Console.WriteLine(" File integrity check complete.")
    '            Else
    '                Console.WriteLine(" File integrity check failed.")
    '                Console.WriteLine("Deleting file and redownloading file.")

    '                System.IO.File.Delete(fileName)
    '                downloadRequiredDLLFile(fileName, downloadURL, properCheckSum, True)
    '            End If
    '        Else
    '            Console.WriteLine()
    '            Console.Write("Downloading " & fileName & "... ")
    '            attemptedDownloads += 1

    '            webClient.DownloadFile(downloadURL, fileName)

    '            Console.WriteLine("Download complete.")

    '            Console.Write("Checking downloaded file integrity... ")
    '            If calculateSHA1Value(fileName) = properCheckSum Then
    '                Console.WriteLine("Integrity validated.")
    '                Console.WriteLine("Integrity check complete.")
    '                Console.WriteLine()
    '            Else
    '                Console.WriteLine("Corrupted file detected.")

    '                If attemptedDownloads <> 5 Then
    '                    Console.WriteLine("Corrupted file detected.  Attempting to re-download.")
    '                    If IO.File.Exists(fileName) = True Then System.IO.File.Delete(fileName)
    '                    downloadRequiredDLLFile(fileName, downloadURL, properCheckSum, False)
    '                Else
    '                    Console.WriteLine("Multiple downloads have been attempted, all failed.")
    '                    Throw New Exception()
    '                End If
    '            End If

    '            webClient.Dispose()
    '            webClient = Nothing
    '        End If
    '    Catch ex As System.Net.WebException
    '        If attemptedDownloads <> 5 Then
    '            If IO.File.Exists(fileName) = True Then System.IO.File.Delete(fileName)
    '            downloadRequiredDLLFile(fileName, downloadURL, properCheckSum, False)
    '        Else
    '            Console.WriteLine("There was an error while attempting to download required assemblies.")
    '            Console.WriteLine("Update process aborted.")

    '            If IO.File.Exists(fileName) = True Then IO.File.Delete(fileName)

    '            Threading.Thread.Sleep(5000)
    '            Process.GetCurrentProcess.Kill()
    '        End If
    '    Catch ex As Exception
    '        If attemptedDownloads <> 5 Then
    '            If System.IO.File.Exists(fileName) Then System.IO.File.Delete(fileName)
    '            downloadRequiredDLLFile(fileName, downloadURL, properCheckSum, False)
    '        Else
    '            Console.WriteLine("There was an error while attempting to download required assemblies.")
    '            Console.WriteLine("Update process aborted.")

    '            If System.IO.File.Exists(fileName) Then IO.File.Delete(fileName)

    '            Threading.Thread.Sleep(5000)
    '            Process.GetCurrentProcess.Kill()
    '        End If
    '    End Try
    'End Sub

    'Sub requiredDLLDownloadProcessSubFunction(fi As FileInfo)
    '    Console.WriteLine()
    '    Console.WriteLine("Checking for required DLLs.")

    '    'If fi.Name.ToLower.Contains("restore point creator") Then
    '    '    downloadRequiredDLLFile("Microsoft.Win32.TaskScheduler.dll", Microsoft_Win32_TaskScheduler_dll_URL, Microsoft_Win32_TaskScheduler_dll_checksum, True)
    '    '    downloadRequiredDLLFile("SmoothProgressBar.dll", SmoothProgressBar_dll_URL, SmoothProgressBar_dll_checksum, True)
    '    If fi.Name.ToLower.Contains("simpleqr") Then
    '        downloadRequiredDLLFile("ThoughtWorks.QRCode.dll", ThoughtWorks_QRCode_dll_URL, ThoughtWorks_QRCode_dll_checksum, True)
    '    ElseIf fi.Name.ToLower.Contains("yawa") Then
    '        downloadRequiredDLLFile("Microsoft.Win32.TaskScheduler.dll", Microsoft_Win32_TaskScheduler_dll_URL, Microsoft_Win32_TaskScheduler_dll_checksum, True)
    '    ElseIf fi.Name.ToLower.Contains("start program at startup without uac prompt") Then
    '        downloadRequiredDLLFile("Microsoft.Win32.TaskScheduler.dll", Microsoft_Win32_TaskScheduler_dll_URL, Microsoft_Win32_TaskScheduler_dll_checksum, True)
    '    ElseIf fi.Name.ToLower.Contains("mbackup") Then
    '        downloadRequiredDLLFile("SmoothProgressBar.dll", SmoothProgressBar_dll_URL, SmoothProgressBar_dll_checksum, True)
    '        downloadRequiredDLLFile("ICSharpCode.SharpZipLib.dll", ICSharpCode_SharpZipLib_dll_URL, ICSharpCode_SharpZipLib_dll_checksum, True)
    '    End If

    '    Console.WriteLine()
    'End Sub

    Sub Main()
        'Dim psi As ProcessStartInfo
        'Dim proc As Process

        Console.WriteLine("===============================================")
        Console.WriteLine("==             Tom's App Updater             ==")
        Console.WriteLine("===============================================")
        Console.WriteLine("")

        If System.Environment.GetCommandLineArgs.Count = 2 Then
            'System.Environment.CommandLine
            Dim fileName As String = System.Environment.GetCommandLineArgs(1).ToString
            Dim fileInfo As System.IO.FileInfo

            'If Text.RegularExpressions.Regex.IsMatch(fileName, "^--dlldownload=", Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            '    fileName = Text.RegularExpressions.Regex.Replace(fileName, "--dlldownload=", "", Text.RegularExpressions.RegexOptions.IgnoreCase)

            '    fileInfo = New System.IO.FileInfo(fileName)
            '    Console.WriteLine("Beginning Required DLL Download Process.")
            '    requiredDLLDownloadProcessSubFunction(fileInfo)

            '    Console.WriteLine("Required DLL Download Complete.")
            '    Console.WriteLine("You may now relaunch the program.")
            'ElseIf Text.RegularExpressions.Regex.IsMatch(fileName, "^--file=", Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            If Text.RegularExpressions.Regex.IsMatch(fileName, "^--file=", Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                fileName = Text.RegularExpressions.Regex.Replace(fileName, "--file=", "", Text.RegularExpressions.RegexOptions.IgnoreCase)

                fileInfo = New System.IO.FileInfo(fileName)
                searchForProcess(fileName)

                'Console.WriteLine("Process not found.")
                Console.WriteLine("Beginning update process.")

                'requiredDLLDownloadProcessSubFunction(fileInfo)

                Console.Write("Checking for new file...")

                Debug.WriteLine(New System.IO.FileInfo(fileName).FullName & ".new")
                If System.IO.File.Exists(fileName & ".new") Then
                    Console.WriteLine(" Found new file.")

                    ' ===============================
                    ' == Begin code that runs NGEN ==
                    ' ===============================
                    'Console.Write("Removing old .NET Cached Compiled Assembly...")

                    'psi = New ProcessStartInfo
                    'If boolWinXP = False Then
                    '    psi.Verb = "runas"
                    'End If
                    'psi.UseShellExecute = True
                    'psi.FileName = IO.Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe")
                    ''psi.Arguments = String.Format("uninstall {0}{1}{0}", Chr(34), fileInfo.FullName)
                    'psi.Arguments = String.Format("uninstall {0}{1}{0}", strQuote, fileInfo.FullName)
                    'psi.WindowStyle = ProcessWindowStyle.Hidden

                    'proc = Process.Start(psi)

                    'proc.WaitForExit()

                    'Console.WriteLine(" Done.")
                    'proc = Nothing
                    'psi = Nothing
                    ' ==============================
                    ' == Ends code that runs NGEN ==
                    ' ==============================

                    Console.Write("Deleting old file...")
                    Try
                        System.IO.File.Delete(fileName)
                        Console.WriteLine(" Done.")
                    Catch ex As System.UnauthorizedAccessException
                        Console.WriteLine(" ERROR!")
                        Console.WriteLine("Parent process is still in use!  Unable to update at this time.")

                        deleteFileAtReboot(fileName, False)
                        deleteFileAtReboot(fileName & ".new", fileName, False)

                        Console.WriteLine("Update Process Aborted! Update will occur the next time this system reboots.")
                        Threading.Thread.Sleep(5000)
                        Exit Sub
                    End Try

                    Console.Write("Renaming File...")
                    System.IO.File.Move(fileName & ".new", fileName)
                    Console.WriteLine(" Done.")

                    ' ===============================
                    ' == Begin code that runs NGEN ==
                    ' ===============================
                    'Console.Write("Installing .NET Cached Compiled Assembly...")

                    'psi = New ProcessStartInfo
                    'If boolWinXP = False Then
                    '    psi.Verb = "runas"
                    'End If
                    'psi.UseShellExecute = True
                    'psi.FileName = IO.Path.Combine(Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory(), "ngen.exe")
                    'psi.Arguments = String.Format("install {0}{1}{0}", strQuote, fileInfo.FullName)
                    'psi.WindowStyle = ProcessWindowStyle.Hidden
                    'proc = Process.Start(psi)
                    'proc.WaitForExit()
                    'proc = Nothing
                    'psi = Nothing

                    'Console.WriteLine(" Done.")
                    ' ==============================
                    ' == Ends code that runs NGEN ==
                    ' ==============================

                    'checkForRequiredDLLs(fileName)

                    Console.WriteLine("Update Complete.")
                    Console.WriteLine("You may now relaunch the program.")
                Else
                    Console.WriteLine(" New file not found.")
                    Console.WriteLine("Something went horribly wrong!")
                    Console.WriteLine("Update process aborted.")
                End If
            End If
        Else
            Console.WriteLine("Invalid Command Line Arguments have been applied.")
            Console.WriteLine("Process aborted.")
        End If

        Threading.Thread.Sleep(5000)
    End Sub

    'Private Sub checkForRequiredDLLs(fileName As String)
    '    Console.WriteLine("Checking for required DLL files...")
    '    Dim webRequest As WebRequest = webRequest.Create("http://www.toms-world.org/requireddlls.php?program=" & fileName)
    '    Dim webresponse As WebResponse = webRequest.GetResponse()
    '    Dim inStream As StreamReader = New System.IO.StreamReader(webresponse.GetResponseStream())
    '    Dim requiredDLLsFromSite As String = inStream.ReadToEnd.Trim()
    '    webresponse.Close()
    '    webresponse = Nothing
    '    webRequest = Nothing
    '    inStream.Close()
    '    inStream = Nothing

    '    If requiredDLLsFromSite.Contains(",") Then
    '        Dim requiredDLLsFromSiteSplit As String() = requiredDLLsFromSite.Split(",")
    '        For i = 0 To requiredDLLsFromSiteSplit.Count - 1
    '            requiredDLLsFromSiteSplit(i) = requiredDLLsFromSiteSplit(i).Trim
    '            If System.IO.File.Exists(requiredDLLsFromSiteSplit(i)) Then
    '                Console.WriteLine("Required """ & requiredDLLsFromSiteSplit(i) & """ found.")
    '            Else
    '                Console.Write("Downloading required """ & requiredDLLsFromSiteSplit(i) & """ DLL... ")
    '                Dim webClient As New System.Net.WebClient
    '                webClient.DownloadFile("http://www.toms-world.org/download/" & requiredDLLsFromSiteSplit(i), requiredDLLsFromSiteSplit(i))
    '                webClient = Nothing
    '                Console.WriteLine("DLL Download Complete.")
    '            End If
    '        Next
    '    Else
    '        requiredDLLsFromSite = requiredDLLsFromSite.Trim
    '        If System.IO.File.Exists(requiredDLLsFromSite) Then
    '            Console.WriteLine("Required """ & requiredDLLsFromSite & """ found.")
    '        Else
    '            Console.Write("Downloading required """ & requiredDLLsFromSite & """ DLL... ")
    '            Dim webClient As New System.Net.WebClient
    '            webClient.DownloadFile("http://www.toms-world.org/download/" & requiredDLLsFromSite, requiredDLLsFromSite)
    '            webClient = Nothing
    '            Console.WriteLine("DLL Download Complete.")
    '        End If
    '    End If

    '    Console.WriteLine("DLL Check Complete.")
    'End Sub
End Module