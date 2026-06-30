Option Explicit

Private mLogs() As String
Private mLogCount As Long
Private mLogOffset As Long
Private Const LOG_MAX_COUNT As Long = 500
Public Const LOG_PERSIST_SHEET As Long = 1
Public Const LOG_PERSIST_FILE As Long = 2
Public Const LOG_PERSIST_BOTH As Long = 3

Private mFileLogPath As String
Private mFileLogPathReady As Boolean

Public Sub Log(ByVal msg As String, Optional ByVal persist As Long = 0)
    Dim line As String
    line = Format(Now, "hh:mm:ss") & "  " & msg

    Debug.Print line

    If mLogCount = 0 And mLogOffset = 0 Then
        ReDim mLogs(0 To LOG_MAX_COUNT - 1)
    End If

    Dim slot As Long
    slot = (mLogOffset + mLogCount) Mod LOG_MAX_COUNT

    mLogs(slot) = line

    If mLogCount < LOG_MAX_COUNT Then
        mLogCount = mLogCount + 1
    Else
        mLogOffset = (mLogOffset + 1) Mod LOG_MAX_COUNT
    End If

    If persist And LOG_PERSIST_FILE Then AppendLogLineToFile line
End Sub

Public Sub LogPersist(ByVal msg As String)
    Log msg, LOG_PERSIST_FILE
End Sub

Public Function GetPersistedLogFilePath() As String
    EnsureFileLogPath
    GetPersistedLogFilePath = mFileLogPath
End Function

Public Sub FlushToSheet()
    If mLogCount = 0 Then
        MsgBox UiMsgDebugLogEmptyText(), vbInformation
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To mLogCount - 1
        Dim idx As Long
        idx = (mLogOffset + i) Mod LOG_MAX_COUNT
        AppendLogLineToFile mLogs(idx)
    Next i

    MsgBox mLogCount & UiMsgDebugLogFlushSuffixText() & vbCrLf & GetPersistedLogFilePath(), vbInformation

    On Error Resume Next
    Shell "notepad.exe " & Chr$(34) & GetPersistedLogFilePath() & Chr$(34), vbNormalFocus
    On Error GoTo 0
End Sub

Public Sub ClearLog()
    mLogCount = 0
    mLogOffset = 0
    On Error Resume Next
    ReDim mLogs(0 To LOG_MAX_COUNT - 1)
    On Error GoTo 0
    Debug.Print "--- log clear ---"
End Sub

Private Sub EnsureFileLogPath()
    If mFileLogPathReady And Len(mFileLogPath) > 0 Then Exit Sub

    Dim baseName As String
    baseName = Trim$(ThisWorkbook.Name)
    If Len(baseName) = 0 Then baseName = "Workbook"
    baseName = Replace(baseName, ":", "_")
    baseName = Replace(baseName, "\", "_")
    baseName = Replace(baseName, "/", "_")
    baseName = Replace(baseName, "?", "_")
    baseName = Replace(baseName, "*", "_")
    baseName = Replace(baseName, """", "_")
    baseName = Replace(baseName, "<", "_")
    baseName = Replace(baseName, ">", "_")
    baseName = Replace(baseName, "|", "_")

    mFileLogPath = Environ$("TEMP") & "\" & baseName & "_DebugLog.txt"
    mFileLogPathReady = True
End Sub

Private Sub AppendLogLineToFile(ByVal line As String)
    On Error Resume Next

    EnsureFileLogPath
    If Len(mFileLogPath) = 0 Then Exit Sub

    Dim fileNo As Integer
    fileNo = FreeFile
    Open mFileLogPath For Append Access Write As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:mm:ss") & "  " & line
    Close #fileNo

    On Error GoTo 0
End Sub
