Attribute VB_Name = "LogHandler"
Public Sub SaveLogInFile(strLog As String)
    Dim logPath As String

    logPath = App.Path & "\" & "Logs\"
    If Right(logPath, 1) <> "\" Then logPath = logPath & "\"
    
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
    End If
    
    Open (logPath & Format(Date, "YYYY-MM-DD") & ".log") For Append As #1
    Print #1, CStr(Time) & "> " & strLog
    Close #1
End Sub
