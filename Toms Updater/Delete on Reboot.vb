Module Delete_on_Reboot
    Private Declare Auto Function MoveFileEx Lib "kernel32.dll" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Int32) As Boolean

    Public Sub deleteFileAtReboot(fileName As String, Optional askForReboot As Boolean = True)
        MoveFileEx(fileName, vbNullString, 4)

        If askForReboot = True Then
            Dim rebootRequestResponse As MsgBoxResult = MsgBox("The file has been scheduled to be deleted at reboot." & vbCrLf & vbCrLf & "Do you want to reboot your computer now?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Reboot now?")

            If rebootRequestResponse = MsgBoxResult.Yes Then
                Shell("shutdown.exe -r -t 0", AppWinStyle.Hide)
            End If
        End If
    End Sub

    Public Sub deleteFileAtReboot(fileName As String, newFileName As String, Optional askForReboot As Boolean = True)
        MoveFileEx(fileName, newFileName, 4)

        If askForReboot = True Then
            Dim rebootRequestResponse As MsgBoxResult = MsgBox("The file has been scheduled to be deleted at reboot." & vbCrLf & vbCrLf & "Do you want to reboot your computer now?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Reboot now?")

            If rebootRequestResponse = MsgBoxResult.Yes Then
                Shell("shutdown.exe -r -t 0", AppWinStyle.Hide)
            End If
        End If
    End Sub
End Module