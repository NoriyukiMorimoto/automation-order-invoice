Option Explicit

' 基本情報シートの施工会社ブロック27行目(注文番号、mod_OrderTpl_Shared.ORDER_TPL_BLOCK_ORDER_NO_ROW)
' セルをダブルクリックした際に、テンキー入力補助フォーム(frmNumericKeypad)を表示する。
' 対象列は F9(施工会社数)に応じて F/I/L/O/R/U/X/AA/AD/AG(3列おき、最大10社)まで自動的に広がる。
' モーダルフォームをダブルクリックイベント中に同期表示するとハングする恐れがあるため、
' mod_PrefectureSelector と同様に Application.OnTime でイベント処理完了後に遅延起動する。
' 改修履歴: CHANGELOG.md 参照

Private mScheduled As Boolean
Private mPendingSheetName As String
Private mPendingCellAddress As String

' 施工会社ブロック27行目(注文番号)のいずれかの会社列セルかどうかを判定する
Public Function IsOrderNumberKeypadTarget(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    If wsInfo Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.MergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    Dim topLeft As Range
    Set topLeft = targetArea.Cells(1, 1)

    If topLeft.Row <> mod_OrderTpl_Shared.ORDER_TPL_BLOCK_ORDER_NO_ROW Then Exit Function

    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 1 To vendorCount
        If topLeft.Column = mod_Construction_BasicTotals.BasicInfoVendorColumn(i) Then
            IsOrderNumberKeypadTarget = True
            Exit Function
        End If
    Next i
End Function

' ダブルクリックからのエントリーポイント。
' イベント中にモーダルフォームを同期表示するとハングするため、OnTime で遅延起動する。
Public Sub RequestOrderNumberKeypad(ByVal wsInfo As Worksheet, ByVal target As Range)
    If wsInfo Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    If mScheduled Then Exit Sub

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.MergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    mPendingSheetName = wsInfo.Name
    mPendingCellAddress = targetArea.Cells(1, 1).Address(False, False)
    mScheduled = True

    mod_DebugLog.Log "[OrderNoKeypad] scheduled cell=" & mPendingCellAddress

    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledOrderNumberKeypad", _
                       Schedule:=True
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                           Procedure:="mod_BasicInfoOrderNumberKeypad.RunScheduledOrderNumberKeypad", _
                           Schedule:=True
        If Err.Number <> 0 Then
            Err.Clear
            mScheduled = False
        End If
    End If
    On Error GoTo 0
End Sub

' Application.OnTime から呼び出されるエントリーポイント(Public 必須)
Public Sub RunScheduledOrderNumberKeypad()
    mScheduled = False
    DoEvents
    ShowOrderNumberKeypadForPendingCell
End Sub

Private Sub ShowOrderNumberKeypadForPendingCell()
    If Len(mPendingSheetName) = 0 Then Exit Sub
    If Len(mPendingCellAddress) = 0 Then Exit Sub

    Dim wsInfo As Worksheet
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Worksheets(mPendingSheetName)
    On Error GoTo 0
    If wsInfo Is Nothing Then Exit Sub

    Dim targetCell As Range
    On Error Resume Next
    Set targetCell = wsInfo.Range(mPendingCellAddress)
    On Error GoTo 0

    mPendingSheetName = ""
    mPendingCellAddress = ""

    If targetCell Is Nothing Then Exit Sub

    ShowOrderNumberKeypad wsInfo, targetCell
End Sub

' テンキーフォームを表示し、確定値をセルへ書き戻す
Public Sub ShowOrderNumberKeypad(ByVal wsInfo As Worksheet, ByVal targetCell As Range)
    If wsInfo Is Nothing Then Exit Sub
    If targetCell Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = targetCell
    On Error Resume Next
    If anchor.MergeCells Then Set anchor = anchor.MergeArea.Cells(1, 1)
    On Error GoTo 0

    Dim f As frmNumericKeypad
    Set f = New frmNumericKeypad
    f.SetValue CommonNzText(anchor.value)
    f.Show vbModal

    Dim isConfirmed As Boolean
    isConfirmed = f.confirmed

    Dim inputText As String
    inputText = f.GetValue()

    Unload f
    Set f = Nothing

    If Not isConfirmed Then Exit Sub

    WriteOrderNumberValue wsInfo, anchor, inputText
End Sub

Private Sub WriteOrderNumberValue(ByVal wsInfo As Worksheet, ByVal anchor As Range, ByVal inputText As String)
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents

    On Error GoTo CleanExit
    Application.EnableEvents = False

    Dim s As String
    s = Trim$(inputText)

    If s = "" Then
        anchor.value = ""
    ElseIf IsNumeric(s) Then
        If InStr(s, ".") > 0 Then
            anchor.value = CDbl(s)
        Else
            On Error Resume Next
            anchor.value = CLng(s)
            If Err.Number <> 0 Then
                Err.Clear
                anchor.value = CDbl(s)
            End If
            On Error GoTo CleanExit
        End If
    Else
        anchor.value = s
    End If

    mod_BasicInfoGuide.OnCellChanged wsInfo, anchor
    mod_OrderTpl_Header.HandleBasicInfoHeaderSourceChange wsInfo, anchor

    mod_DebugLog.Log "[OrderNoKeypad] applied cell=" & anchor.Address(False, False) & " value=" & s

CleanExit:
    Application.EnableEvents = prevEvents
End Sub
