Option Explicit

Private mLogs() As String
Private mLogCount As Long
Private mLogOffset As Long
Private Const LOG_SHEET_NAME As String = "DebugLog"
Private Const LOG_MAX_COUNT As Long = 500

Public Sub Log(ByVal msg As String)
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
End Sub

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
        Set ws = ThisWorkbook.worksheets.Add(After:=ThisWorkbook.worksheets(ThisWorkbook.worksheets.Count))
        ws.Name = LOG_SHEET_NAME
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
