Option Explicit

Private mLogs() As String
Private mLogCount As Long
Private Const LOG_SHEET_NAME As String = "DebugLog"

Public Sub Log(ByVal msg As String)
    Dim line As String
    line = Format(Now, "hh:mm:ss") & "  " & msg

    Debug.Print line

    If mLogCount = 0 Then
        ReDim mLogs(0 To 99)
    ElseIf mLogCount > UBound(mLogs) Then
        ReDim Preserve mLogs(0 To UBound(mLogs) + 100)
    End If
    mLogs(mLogCount) = line
    mLogCount = mLogCount + 1
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
    ws.Range("A1").value = "時刻"
    ws.Range("B1").value = "ログ"

    Dim i As Long
    For i = 0 To mLogCount - 1
        Dim parts() As String
        parts = Split(mLogs(i), "  ", 2)
        If UBound(parts) >= 1 Then
            ws.Cells(i + 2, 1).value = parts(0)
            ws.Cells(i + 2, 2).value = parts(1)
        Else
            ws.Cells(i + 2, 2).value = mLogs(i)
        End If
    Next i

    ws.Columns("A:B").AutoFit
    ws.Activate
    MsgBox mLogCount & UiMsgDebugLogFlushSuffixText(), vbInformation
End Sub

Public Sub ClearLog()
    mLogCount = 0
    ReDim mLogs(0 To 0)
    Debug.Print "--- ログクリア ---"
End Sub
