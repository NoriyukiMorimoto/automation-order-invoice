Option Explicit

' 基本情報シートの施工会社ブロック39行目(労災保険 加入負担)のセルをダブルクリックした際に、
' frmNumericKeypad を労災モード(甲/乙チェックボックス)で表示する。
' 対象列は F9(施工会社数)に応じて F/I/L/O/R/U/X/AA/AD/AG(3列おき・最大10社)まで自動で広がる。
' モーダルフォームをダブルクリックイベント内で同期表示するとハングする恐れがあるため、
' mod_BasicInfoOrderNumberKeypad と同様に Application.OnTime でイベント完了後に遅延起動する。
' 甲/乙の選択結果をセルへ横並びで書き込み、対応する受注者用(略称)シートの
' H30(甲)/J30(乙)へチェック字形を上下左右中央揃えで表示する。
' 改修履歴: CHANGELOG.md 参照

Private Const BASIC_INFO_LABOR_INSURANCE_ROW As Long = 39
Private Const ACCEPTANCE_KO_CELL As String = "H30"
Private Const ACCEPTANCE_OTSU_CELL As String = "J30"

Private mScheduled As Boolean
Private mPendingSheetName As String
Private mPendingCellAddress As String

' 施工会社ブロック39行目(労災保険 加入負担)のいずれかの会社列セルかどうかを判定する
Public Function IsLaborInsuranceTarget(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    If wsInfo Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.MergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    Dim topLeft As Range
    Set topLeft = targetArea.Cells(1, 1)

    If topLeft.Row <> BASIC_INFO_LABOR_INSURANCE_ROW Then Exit Function

    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 1 To vendorCount
        If topLeft.Column = mod_Construction_BasicTotals.BasicInfoVendorColumn(i) Then
            IsLaborInsuranceTarget = True
            Exit Function
        End If
    Next i
End Function

' ダブルクリックからのエントリーポイント。
' イベント内でモーダル表示するとハングするため、OnTime で遅延起動する。
Public Sub RequestLaborInsuranceSelection(ByVal wsInfo As Worksheet, ByVal target As Range)
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

    mod_DebugLog.Log "[RosaiSelect] scheduled cell=" & mPendingCellAddress

    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledLaborInsuranceSelection", _
                       Schedule:=True
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                           Procedure:="mod_BasicInfoLaborInsurance.RunScheduledLaborInsuranceSelection", _
                           Schedule:=True
        If Err.Number <> 0 Then
            Err.Clear
            mScheduled = False
        End If
    End If
    On Error GoTo 0
End Sub

' Application.OnTime から呼び出されるエントリーポイント(Public 必須)
Public Sub RunScheduledLaborInsuranceSelection()
    mScheduled = False
    DoEvents
    ShowLaborInsuranceForPendingCell
End Sub

Private Sub ShowLaborInsuranceForPendingCell()
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

    ShowLaborInsuranceSelector wsInfo, targetCell
End Sub

' 労災モードのフォームを表示し、確定した甲/乙の選択をセルと受注者用シートへ書き戻す
Public Sub ShowLaborInsuranceSelector(ByVal wsInfo As Worksheet, ByVal targetCell As Range)
    If wsInfo Is Nothing Then Exit Sub
    If targetCell Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = targetCell
    On Error Resume Next
    If anchor.MergeCells Then Set anchor = anchor.MergeArea.Cells(1, 1)
    On Error GoTo 0

    Dim koInit As Boolean
    Dim otsuInit As Boolean
    ParseSelection CommonNzText(anchor.Value), koInit, otsuInit

    Dim f As frmNumericKeypad
    Set f = New frmNumericKeypad
    f.ConfigureLaborInsuranceMode koInit, otsuInit
    f.Show vbModal

    Dim isConfirmed As Boolean
    Dim koResult As Boolean
    Dim otsuResult As Boolean
    isConfirmed = f.confirmed
    koResult = f.rosaiKo
    otsuResult = f.rosaiOtsu

    Unload f
    Set f = Nothing

    If Not isConfirmed Then Exit Sub

    WriteLaborInsuranceValue wsInfo, anchor, koResult, otsuResult
End Sub

Private Sub WriteLaborInsuranceValue(ByVal wsInfo As Worksheet, ByVal anchor As Range, _
                                     ByVal koChecked As Boolean, ByVal otsuChecked As Boolean)
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents

    On Error GoTo CleanExit
    Application.EnableEvents = False

    anchor.Value = BuildCellText(koChecked, otsuChecked)
    anchor.HorizontalAlignment = xlCenter

    WriteAcceptanceSheetCheck wsInfo, anchor.Column, koChecked, otsuChecked

    mod_DebugLog.Log "[RosaiSelect] applied cell=" & anchor.Address(False, False) & _
                     " ko=" & koChecked & " otsu=" & otsuChecked

CleanExit:
    Application.EnableEvents = prevEvents
End Sub

' 対応する受注者用(略称)シートの H30(甲)/J30(乙)へチェック字形を中央揃えで表示する
Private Sub WriteAcceptanceSheetCheck(ByVal wsInfo As Worksheet, ByVal valueColumn As Long, _
                                      ByVal koChecked As Boolean, ByVal otsuChecked As Boolean)
    Dim ws As Worksheet
    Set ws = ResolveAcceptanceSheetForColumn(wsInfo, valueColumn)
    If ws Is Nothing Then Exit Sub

    ApplyCheckGlyph ws.Range(ACCEPTANCE_KO_CELL), koChecked
    ApplyCheckGlyph ws.Range(ACCEPTANCE_OTSU_CELL), otsuChecked
End Sub

Private Sub ApplyCheckGlyph(ByVal cell As Range, ByVal checked As Boolean)
    If cell Is Nothing Then Exit Sub
    cell.Value = CheckGlyph(checked)
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
End Sub

' 会社列(値列)から業者インデックスを求め、対応する受注者用シートを解決する
Private Function ResolveAcceptanceSheetForColumn(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Worksheet
    Dim vendorIndex As Long
    vendorIndex = mod_VendorMaster.GetVendorIndexFromValueColumnPublic(valueColumn)
    If vendorIndex < 1 Then Exit Function

    Dim companyName As String
    companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, vendorIndex)
    If companyName = "" Then Exit Function

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).Value))

    Dim vendorName As String
    Dim aliasText As String
    Dim workText As String
    If Not mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, vendorName, aliasText, workText) Then Exit Function
    If aliasText = "" Then Exit Function

    Dim sheetName As String
    sheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
                    mod_OrderTpl_Shared.OrderTplBaseNameContractorText(), aliasText)
    If Not mod_OrderTpl_Shared.OrderTplSheetExists(sheetName) Then Exit Function

    Set ResolveAcceptanceSheetForColumn = ThisWorkbook.Worksheets(sheetName)
End Function

' チェック字形: ON=U+2611(レ点入りボックス) / OFF=U+2610(空ボックス)
Private Function CheckGlyph(ByVal checked As Boolean) As String
    If checked Then
        CheckGlyph = ChrW$(&H2611)
    Else
        CheckGlyph = ChrW$(&H2610)
    End If
End Function

' セル内表示: 「(甲字形) 甲 (全角空白) (乙字形) 乙」の横並び
Private Function BuildCellText(ByVal koChecked As Boolean, ByVal otsuChecked As Boolean) As String
    BuildCellText = CheckGlyph(koChecked) & " " & KoText() & ChrW$(&H3000) & _
                    CheckGlyph(otsuChecked) & " " & OtsuText()
End Function

' 既存セル値から甲/乙の現在状態を復元する(フォーム初期表示用)
Private Sub ParseSelection(ByVal current As String, ByRef koChecked As Boolean, ByRef otsuChecked As Boolean)
    Dim onGlyph As String
    onGlyph = ChrW$(&H2611)
    koChecked = (InStr(current, onGlyph & " " & KoText()) > 0)
    otsuChecked = (InStr(current, onGlyph & " " & OtsuText()) > 0)
End Sub

' "甲"
Private Function KoText() As String
    KoText = ChrW$(&H7532)
End Function

' "乙"
Private Function OtsuText() As String
    OtsuText = ChrW$(&H4E59)
End Function
