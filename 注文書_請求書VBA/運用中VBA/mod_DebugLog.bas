Option Explicit

Private mLogs() As String
Private mLogCount As Long
Private mLogOffset As Long
Private Const LOG_SHEET_NAME As String = "DebugLog"
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

    If persist And LOG_PERSIST_SHEET Then AppendLogLineToSheet line
    If persist And LOG_PERSIST_FILE Then AppendLogLineToFile line
End Sub

Public Sub LogPersist(ByVal msg As String)
    Log msg, LOG_PERSIST_BOTH
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

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = GetOrCreateLogSheet()
    End If

    ws.Cells.ClearContents
    ws.Range("A1").value = ChrW$(&H6642) & ChrW$(&H523B)
    ws.Range("B1").value = ChrW$(&H30ED) & ChrW$(&H30B0)

    Dim i As Long
    For i = 0 To mLogCount - 1
        Dim idx As Long
        idx = (mLogOffset + i) Mod LOG_MAX_COUNT
        Dim parts() As String
        parts = Split(mLogs(idx), "  ", 2)
        If UBound(parts) >= 1 Then
            ws.Cells(i + 2, 1).value = parts(0)
            ws.Cells(i + 2, 2).value = parts(1)
        Else
            ws.Cells(i + 2, 2).value = mLogs(idx)
        End If
    Next i

    ws.Columns("A:B").AutoFit
    ws.Activate
    MsgBox mLogCount & UiMsgDebugLogFlushSuffixText(), vbInformation
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

Private Function GetOrCreateLogSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
        ws.Range("A1").Value = ChrW$(&H6642) & ChrW$(&H523B)
        ws.Range("B1").Value = ChrW$(&H30ED) & ChrW$(&H30B0)
    End If

    Set GetOrCreateLogSheet = ws
End Function

Private Sub AppendLogLineToSheet(ByVal line As String)
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = GetOrCreateLogSheet()
    If ws Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If nextRow < 1 Then nextRow = 1
    If Len(Trim$(CStr(ws.Cells(1, 1).Value))) > 0 Then
        nextRow = nextRow + 1
    Else
        nextRow = 2
    End If

    Dim parts() As String
    parts = Split(line, "  ", 2)
    If UBound(parts) >= 1 Then
        ws.Cells(nextRow, 1).Value = parts(0)
        ws.Cells(nextRow, 2).Value = parts(1)
    Else
        ws.Cells(nextRow, 2).Value = line
    End If

    On Error GoTo 0
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
