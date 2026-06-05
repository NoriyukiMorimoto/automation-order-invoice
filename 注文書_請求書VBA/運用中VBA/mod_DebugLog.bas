Option Explicit

'==========================================================================
'  デバッグログモジュール  mod_DebugLog
'  B6変更時のコンボ表示フロー追跡用
'  使い方：
'    1. このモジュールをインポート
'    2. B6を変更する
'    3. イミディエイトウィンドウ(Ctrl+G)またはログシートで結果確認
'    4. 確認後は mod_DebugLog.ClearLog で消去
'==========================================================================

Private mLogs() As String
Private mLogCount As Long
Private Const LOG_SHEET_NAME As String = "DebugLog"

'--------------------------------------------------------------------------
' ログ追記（イミディエイトウィンドウ＋内部バッファ両方に出力）
'--------------------------------------------------------------------------
Public Sub Log(ByVal msg As String)
    Dim line As String
    line = Format(Now, "hh:mm:ss") & "  " & msg

    ' イミディエイトウィンドウへ
    Debug.Print line

    ' 内部バッファへ
    If mLogCount = 0 Then
        ReDim mLogs(0 To 99)
    ElseIf mLogCount > UBound(mLogs) Then
        ReDim Preserve mLogs(0 To UBound(mLogs) + 100)
    End If
    mLogs(mLogCount) = line
    mLogCount = mLogCount + 1
End Sub

'--------------------------------------------------------------------------
' ログをシートに書き出す（B6変更後にこれを手動実行して確認）
'--------------------------------------------------------------------------
Public Sub FlushToSheet()
    If mLogCount = 0 Then
        MsgBox "ログはありません。", vbInformation
        Exit Sub
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
    End If

    ws.Cells.ClearContents
    ws.Range("A1").Value = "時刻"
    ws.Range("B1").Value = "ログ"

    Dim i As Long
    For i = 0 To mLogCount - 1
        Dim parts() As String
        parts = Split(mLogs(i), "  ", 2)
        If UBound(parts) >= 1 Then
            ws.Cells(i + 2, 1).Value = parts(0)
            ws.Cells(i + 2, 2).Value = parts(1)
        Else
            ws.Cells(i + 2, 2).Value = mLogs(i)
        End If
    Next i

    ws.Columns("A:B").AutoFit
    ws.Activate
    MsgBox mLogCount & " 件のログをシートに出力しました。", vbInformation
End Sub

'--------------------------------------------------------------------------
' ログクリア
'--------------------------------------------------------------------------
Public Sub ClearLog()
    mLogCount = 0
    ReDim mLogs(0 To 0)
    Debug.Print "--- ログクリア ---"
End Sub
