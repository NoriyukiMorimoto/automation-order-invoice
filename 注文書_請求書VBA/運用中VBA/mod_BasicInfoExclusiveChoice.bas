Option Explicit

' 基本情報シート施工会社ブロックの排他2択セル(38/39/42行目)ダブルクリックで
' frmSubconSelector を単一選択モードで表示する共通モジュール(適用/キャンセル付き)。
'   行38: 部分払い費用負担    → 甲/乙             → 受注者用 C30(結合セルC30:E33に「甲・乙 負担」の字形をまとめて表示)
'   行39: 労災保険 加入負担    → 甲/乙             → 受注者用 H30(甲)/J30(乙)
'   行42: 建設リサイクル法該当  → 該当する/該当しない → 受注者用 M34(該当する)/R34(該当しない)
' 対象列は F9(施工会社数)に応じ F/I/L/O/R/U/X/AA/AD/AG(3列おき・最大10社)。
' モーダルをイベント内で表示するとハングするため Application.OnTime で遅延起動する。
' 選択結果を基本情報セルへ横並びで表示し、対応する受注者用シートへ
' 排他(選択=ONの字形 / 非選択=OFFの字形)で表示する(行38は結合1セルへまとめて表示、行39/42は2セルへ個別表示)。
' 改修履歴: CHANGELOG.md 参照

Private Const ROSAI_ROW As Long = 39
Private Const RECYCLE_ROW As Long = 42
Private Const PARTIAL_PAYMENT_ROW As Long = 38   ' 追加: 部分払い費用負担
Private Const EXCL_VENDOR_COL_WIDTH As Double = 42.5
Private Const EXCL_BASE_ROW_HEIGHT As Double = 24#

Private mScheduled As Boolean
Private mPendingSheetName As String
Private mPendingCellAddress As String
Private mPendingCaption As String
Private mPendingTopLabel As String
Private mPendingBottomLabel As String
Private mPendingTopCell As String
Private mPendingBottomCell As String
Private mPendingCombinedSuffix As String   ' 追加: 空なら従来どおり2セル表示、値ありなら結合1セル表示

' --- 行38: 部分払いに要する費用の負担 ---------------------------------
Public Function IsPartialPaymentTarget(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    IsPartialPaymentTarget = IsRowVendorTarget(wsInfo, target, PARTIAL_PAYMENT_ROW)
End Function

Public Sub RequestPartialPaymentSelection(ByVal wsInfo As Worksheet, ByVal target As Range)
    ScheduleChoice wsInfo, target, PartialPaymentCaptionText(), KoText(), OtsuText(), _
                   "C30", "", PartialPaymentBurdenSuffixText()
End Sub

' --- 行39: 労災保険 加入負担 -----------------------------------------
Public Function IsLaborInsuranceTarget(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    IsLaborInsuranceTarget = IsRowVendorTarget(wsInfo, target, ROSAI_ROW)
End Function

Public Sub RequestLaborInsuranceSelection(ByVal wsInfo As Worksheet, ByVal target As Range)
    ScheduleChoice wsInfo, target, LaborInsuranceCaptionText(), KoText(), OtsuText(), "H30", "J30"
End Sub

' --- 行42: 建設リサイクル法 該当有無 --------------------------------
Public Function IsRecycleLawTarget(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    IsRecycleLawTarget = IsRowVendorTarget(wsInfo, target, RECYCLE_ROW)
End Function

Public Sub RequestRecycleLawSelection(ByVal wsInfo As Worksheet, ByVal target As Range)
    ScheduleChoice wsInfo, target, RecycleLawCaptionText(), ApplicableText(), NotApplicableText(), "M34", "R34"
End Sub

' --- 対象判定(指定行・施工会社列) -----------------------------------
Private Function IsRowVendorTarget(ByVal wsInfo As Worksheet, ByVal target As Range, ByVal rowNo As Long) As Boolean
    If wsInfo Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.mergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    Dim topLeft As Range
    Set topLeft = targetArea.Cells(1, 1)
    If topLeft.row <> rowNo Then Exit Function

    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 1 To vendorCount
        If topLeft.Column = mod_Construction_BasicTotals.BasicInfoVendorColumn(i) Then
            IsRowVendorTarget = True
            Exit Function
        End If
    Next i
End Function

' --- OnTime 遅延起動 -------------------------------------------------
Private Sub ScheduleChoice(ByVal wsInfo As Worksheet, ByVal target As Range, _
                           ByVal formCaption As String, ByVal topLabel As String, ByVal bottomLabel As String, _
                           ByVal topCell As String, ByVal bottomCell As String, _
                           Optional ByVal combinedSuffix As String = "")
    If wsInfo Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    If mScheduled Then Exit Sub

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.mergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    mPendingSheetName = wsInfo.Name
    mPendingCellAddress = targetArea.Cells(1, 1).Address(False, False)
    mPendingCaption = formCaption
    mPendingTopLabel = topLabel
    mPendingBottomLabel = bottomLabel
    mPendingTopCell = topCell
    mPendingBottomCell = bottomCell
    mPendingCombinedSuffix = combinedSuffix
    mScheduled = True

    mod_DebugLog.Log "[ExclChoice] scheduled cell=" & mPendingCellAddress

    On Error Resume Next
    Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                       Procedure:="'" & ThisWorkbook.Name & "'!RunScheduledExclusiveChoice", _
                       Schedule:=True
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnTime EarliestTime:=Now + TimeValue("00:00:01"), _
                           Procedure:="mod_BasicInfoExclusiveChoice.RunScheduledExclusiveChoice", _
                           Schedule:=True
        If Err.Number <> 0 Then
            Err.Clear
            mScheduled = False
        End If
    End If
    On Error GoTo 0
End Sub

' Application.OnTime から呼び出されるエントリーポイント(Public 必須)
Public Sub RunScheduledExclusiveChoice()
    mScheduled = False
    DoEvents
    ShowForPendingCell
End Sub

Private Sub ShowForPendingCell()
    If Len(mPendingSheetName) = 0 Then Exit Sub
    If Len(mPendingCellAddress) = 0 Then Exit Sub

    Dim wsInfo As Worksheet
    On Error Resume Next
    Set wsInfo = ThisWorkbook.worksheets(mPendingSheetName)
    On Error GoTo 0
    If wsInfo Is Nothing Then Exit Sub

    Dim targetCell As Range
    On Error Resume Next
    Set targetCell = wsInfo.Range(mPendingCellAddress)
    On Error GoTo 0

    Dim formCaption As String, topLabel As String, bottomLabel As String
    Dim topCell As String, bottomCell As String, combinedSuffix As String
    formCaption = mPendingCaption
    topLabel = mPendingTopLabel
    bottomLabel = mPendingBottomLabel
    topCell = mPendingTopCell
    bottomCell = mPendingBottomCell
    combinedSuffix = mPendingCombinedSuffix

    mPendingSheetName = ""
    mPendingCellAddress = ""

    If targetCell Is Nothing Then Exit Sub

    ShowExclusiveChoice wsInfo, targetCell, formCaption, topLabel, bottomLabel, topCell, bottomCell, combinedSuffix
End Sub

Public Sub ShowExclusiveChoice(ByVal wsInfo As Worksheet, ByVal targetCell As Range, _
                               ByVal formCaption As String, ByVal topLabel As String, ByVal bottomLabel As String, _
                               ByVal topCell As String, ByVal bottomCell As String, _
                               Optional ByVal combinedSuffix As String = "")
    If wsInfo Is Nothing Then Exit Sub
    If targetCell Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = targetCell
    On Error Resume Next
    If anchor.MergeCells Then Set anchor = anchor.mergeArea.Cells(1, 1)
    On Error GoTo 0

    Dim topInit As Boolean, bottomInit As Boolean
    ParseSelection CommonNzText(anchor.value), topLabel, bottomLabel, topInit, bottomInit

    Dim items(0 To 1) As String
    items(0) = topLabel
    items(1) = bottomLabel

    Dim initSel As Variant
    If topInit Then
        initSel = Array(topLabel)
    ElseIf bottomInit Then
        initSel = Array(bottomLabel)
    Else
        initSel = Empty
    End If

    Dim f As frmSubconSelector
    Set f = New frmSubconSelector
    f.ConfigureSingleSectionMode formCaption, items, False, "", initSel
    f.Show vbModal

    Dim isConfirmed As Boolean, selected As String
    isConfirmed = f.confirmed
    selected = Trim$(CStr(f.resultBottom))
    If Len(selected) = 0 Then selected = Trim$(CStr(f.resultTop))

    Unload f
    Set f = Nothing

    If Not isConfirmed Then Exit Sub
    If Len(selected) = 0 Then
        mod_DebugLog.Log "[ExclChoice] confirmed but selection empty cell=" & anchor.Address(False, False)
        Exit Sub
    End If

    Dim topResult As Boolean, bottomResult As Boolean
    topResult = (StrComp(selected, topLabel, vbBinaryCompare) = 0)
    bottomResult = (StrComp(selected, bottomLabel, vbBinaryCompare) = 0)
    If (Not topResult) And (Not bottomResult) Then
        mod_DebugLog.Log "[ExclChoice] selection unmatched selected=[" & selected & "]"
        Exit Sub
    End If

    WriteExclusiveChoice wsInfo, anchor, topResult, bottomResult, topLabel, bottomLabel, topCell, bottomCell, combinedSuffix
End Sub

Private Sub WriteExclusiveChoice(ByVal wsInfo As Worksheet, ByVal anchor As Range, _
                                 ByVal topChecked As Boolean, ByVal bottomChecked As Boolean, _
                                 ByVal topLabel As String, ByVal bottomLabel As String, _
                                 ByVal topCell As String, ByVal bottomCell As String, _
                                 Optional ByVal combinedSuffix As String = "")
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents

    On Error GoTo WriteFailed
    Application.EnableEvents = False

    Dim cellText As String
    cellText = BuildCellText(topChecked, bottomChecked, topLabel, bottomLabel)
    anchor.Value = cellText
    anchor.HorizontalAlignment = xlCenter
    ApplyExclusiveChoiceCellLayout anchor
    mod_DebugLog.Log "[ExclChoice] wrote basic cell=" & anchor.Address(False, False) & " value=[" & cellText & "]"

    Dim mirroredSheets As Collection
    Set mirroredSheets = ResolveMirroredSheetsForColumn(wsInfo, anchor.Column)
    Dim ws As Worksheet
    For Each ws In mirroredSheets
        If Len(combinedSuffix) > 0 Then
            ApplyCombinedCheckText ws.Range(topCell), topChecked, bottomChecked, topLabel, bottomLabel, combinedSuffix
        Else
            ApplyCheckGlyph ws.Range(topCell), topChecked
            ApplyCheckGlyph ws.Range(bottomCell), bottomChecked
        End If
    Next ws

    mod_DebugLog.Log "[ExclChoice] applied cell=" & anchor.Address(False, False) & _
                     " top=" & topChecked & " bottom=" & bottomChecked
    GoTo CleanExit

WriteFailed:
    mod_DebugLog.Log "[ExclChoice] write error " & Err.Number & ": " & Err.Description

CleanExit:
    Application.EnableEvents = prevEvents
End Sub

' 基本情報セルの文字色を白にし、行42は折り返し行高を優先(不要時は24)
Private Sub ApplyExclusiveChoiceCellLayout(ByVal anchor As Range)
    If anchor Is Nothing Then Exit Sub

    On Error Resume Next
    anchor.Font.Color = &HFFFFFF
    anchor.VerticalAlignment = xlCenter
    anchor.ShrinkToFit = False

    ' 行42(該当する/該当しない)は文字数が多く折り返し行高を優先する
    If anchor.Row = RECYCLE_ROW Then
        anchor.WrapText = True
        anchor.Worksheet.Columns(anchor.Column).ColumnWidth = EXCL_VENDOR_COL_WIDTH
        anchor.EntireRow.AutoFit
        If anchor.RowHeight < EXCL_BASE_ROW_HEIGHT Then
            anchor.RowHeight = EXCL_BASE_ROW_HEIGHT
        End If
        anchor.WrapText = True
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyCheckGlyph(ByVal cell As Range, ByVal checked As Boolean)
    If cell Is Nothing Then Exit Sub
    Dim c As Range
    Set c = cell
    On Error Resume Next
    If c.MergeCells Then Set c = c.mergeArea.Cells(1, 1)
    On Error GoTo 0
    c.value = CheckGlyph(checked)
    c.HorizontalAlignment = xlCenter
    c.VerticalAlignment = xlCenter
End Sub

' 結合セルへ「(上字形)甲・(下字形)乙 (接尾語)」の形式でまとめて表示する(部分払い費用負担など)
' 例: 甲選択時は甲がON字形・乙がOFF字形 / 乙選択時は甲がOFF字形・乙がON字形(字形はCheckGlyphのChrW参照)
Private Sub ApplyCombinedCheckText(ByVal cell As Range, ByVal topChecked As Boolean, ByVal bottomChecked As Boolean, _
                                   ByVal topLabel As String, ByVal bottomLabel As String, ByVal suffixText As String)
    If cell Is Nothing Then Exit Sub
    Dim c As Range
    Set c = cell
    On Error Resume Next
    If c.MergeCells Then Set c = c.mergeArea.Cells(1, 1)
    On Error GoTo 0
    c.value = CheckGlyph(topChecked) & topLabel & ChrW$(&H30FB) & _
              CheckGlyph(bottomChecked) & bottomLabel & " " & suffixText
    c.HorizontalAlignment = xlCenter
    c.VerticalAlignment = xlCenter
End Sub

' 会社列(値列)から業者インデックスを求め、対応する 受注者用/注文請書/支店控(略称) シートを
' まとめて返す(存在するものだけ)。行20-34はこの3シートともセル構成が完全一致しているため、
' 甲乙/該当有無の表示は3シートすべてへ同じロジックで適用する。
Private Function ResolveMirroredSheetsForColumn(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Collection
    Dim result As Collection
    Set result = New Collection
    Set ResolveMirroredSheetsForColumn = result

    Dim vendorIndex As Long
    vendorIndex = mod_VendorMaster.GetVendorIndexFromValueColumnPublic(valueColumn)
    If vendorIndex < 1 Then Exit Function

    Dim companyName As String
    companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, vendorIndex)
    If companyName = "" Then Exit Function

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))

    Dim vendorName As String, aliasText As String, workText As String
    If Not mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, vendorName, aliasText, workText) Then Exit Function
    If aliasText = "" Then Exit Function

    Dim baseNameList As Variant
    baseNameList = Array(mod_OrderTpl_Shared.OrderTplBaseNameContractorText(), _
                         mod_OrderTpl_Shared.OrderTplBaseNameAcceptanceText(), _
                         mod_OrderTpl_Shared.OrderTplBaseNameBranchCopyText())

    Dim i As Long
    Dim sheetName As String
    For i = LBound(baseNameList) To UBound(baseNameList)
        sheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName(CStr(baseNameList(i)), aliasText)
        If mod_OrderTpl_Shared.OrderTplSheetExists(sheetName) Then
            result.Add ThisWorkbook.worksheets(sheetName)
        End If
    Next i
End Function

' チェック字形: ON=U+2611 / OFF=U+2610
Private Function CheckGlyph(ByVal checked As Boolean) As String
    If checked Then
        CheckGlyph = ChrW$(&H2611)
    Else
        CheckGlyph = ChrW$(&H2610)
    End If
End Function

' 基本情報セル内表示: 「(上字形) 上ラベル (全角空白) (下字形) 下ラベル」
Private Function BuildCellText(ByVal topChecked As Boolean, ByVal bottomChecked As Boolean, _
                               ByVal topLabel As String, ByVal bottomLabel As String) As String
    BuildCellText = CheckGlyph(topChecked) & " " & topLabel & _
                    ChrW$(&H3000) & ChrW$(&H3000) & ChrW$(&H3000) & _
                    CheckGlyph(bottomChecked) & " " & bottomLabel
End Function

' 既存セル値から上/下の現在状態を復元する(フォーム初期表示用)
' 行42の旧表示「該当あり/該当なし」も読み取れるようにする。
Private Sub ParseSelection(ByVal current As String, ByVal topLabel As String, ByVal bottomLabel As String, _
                           ByRef topChecked As Boolean, ByRef bottomChecked As Boolean)
    Dim onGlyph As String
    onGlyph = ChrW$(&H2611)
    topChecked = (InStr(current, onGlyph & " " & topLabel) > 0)
    bottomChecked = (InStr(current, onGlyph & " " & bottomLabel) > 0)

    If (Not topChecked) And (Not bottomChecked) Then
        If StrComp(topLabel, ApplicableText(), vbBinaryCompare) = 0 Then
            topChecked = (InStr(current, onGlyph & " " & ApplicableLegacyText()) > 0)
            bottomChecked = (InStr(current, onGlyph & " " & NotApplicableLegacyText()) > 0)
        End If
    End If
End Sub

' "該当あり"(旧表記)
Private Function ApplicableLegacyText() As String
    ApplicableLegacyText = ChrW$(&H8A72) & ChrW$(&H5F53) & ChrW$(&H3042) & ChrW$(&H308A)
End Function

' "該当なし"(旧表記)
Private Function NotApplicableLegacyText() As String
    NotApplicableLegacyText = ChrW$(&H8A72) & ChrW$(&H5F53) & ChrW$(&H306A) & ChrW$(&H3057)
End Function

' "労災保険 加入負担選択"
Private Function LaborInsuranceCaptionText() As String
    LaborInsuranceCaptionText = ChrW$(&H52B4) & ChrW$(&H707D) & ChrW$(&H4FDD) & ChrW$(&H967A) & _
                                " " & ChrW$(&H52A0) & ChrW$(&H5165) & ChrW$(&H8CA0) & ChrW$(&H62C5) & _
                                ChrW$(&H9078) & ChrW$(&H629E)
End Function

' "建設リサイクル法の対象建設工事に該当の有無選択"
Private Function RecycleLawCaptionText() As String
    RecycleLawCaptionText = ChrW$(&H5EFA) & ChrW$(&H8A2D) & ChrW$(&H30EA) & ChrW$(&H30B5) & ChrW$(&H30A4) & _
                            ChrW$(&H30AF) & ChrW$(&H30EB) & ChrW$(&H6CD5) & ChrW$(&H306E) & ChrW$(&H5BFE) & _
                            ChrW$(&H8C61) & ChrW$(&H5EFA) & ChrW$(&H8A2D) & ChrW$(&H5DE5) & ChrW$(&H4E8B) & _
                            ChrW$(&H306B) & ChrW$(&H8A72) & ChrW$(&H5F53) & ChrW$(&H306E) & ChrW$(&H6709) & _
                            ChrW$(&H7121) & ChrW$(&H9078) & ChrW$(&H629E)
End Function

' "部分払い費用負担選択"
Private Function PartialPaymentCaptionText() As String
    PartialPaymentCaptionText = ChrW$(&H90E8) & ChrW$(&H5206) & ChrW$(&H6255) & ChrW$(&H3044) & _
                                ChrW$(&H8CBB) & ChrW$(&H7528) & ChrW$(&H8CA0) & ChrW$(&H62C5) & _
                                ChrW$(&H9078) & ChrW$(&H629E)
End Function

' "負担"
Private Function PartialPaymentBurdenSuffixText() As String
    PartialPaymentBurdenSuffixText = ChrW$(&H8CA0) & ChrW$(&H62C5)
End Function

' "甲"
Private Function KoText() As String
    KoText = ChrW$(&H7532)
End Function

' "乙"
Private Function OtsuText() As String
    OtsuText = ChrW$(&H4E59)
End Function

' "該当する"
Private Function ApplicableText() As String
    ApplicableText = ChrW$(&H8A72) & ChrW$(&H5F53) & ChrW$(&H3059) & ChrW$(&H308B)
End Function

' "該当しない"
Private Function NotApplicableText() As String
    NotApplicableText = ChrW$(&H8A72) & ChrW$(&H5F53) & ChrW$(&H3057) & ChrW$(&H306A) & ChrW$(&H3044)
End Function

' 施工会社名の変更/受注者用シート再生成時に、基本情報 F38/F39/F42 の現在値から
' 受注者用シートへ 甲/乙・該当する/該当しない を再転記する(ApplyContractorHeader から呼ぶ)。
Public Sub ReapplyExclusiveChoices(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub

    Dim vc As Long
    vc = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    Dim mirroredSheets As Collection
    Set mirroredSheets = ResolveMirroredSheetsForColumn(wsInfo, vc)
    If mirroredSheets.Count = 0 Then Exit Sub

    Dim ws As Worksheet
    For Each ws In mirroredSheets
        ' 行38: 部分払い 甲/乙 -> C30(結合・接尾「負担」)
        ReapplyCombinedCell ws, wsInfo.Cells(PARTIAL_PAYMENT_ROW, vc), _
                            KoText(), OtsuText(), "C30", PartialPaymentBurdenSuffixText()
        ' 行39: 労災 甲/乙 -> H30/J30
        ReapplyTwoCell ws, wsInfo.Cells(ROSAI_ROW, vc), KoText(), OtsuText(), "H30", "J30"
        ' 行42: リサイクル 該当する/該当しない -> M34/R34
        ReapplyTwoCell ws, wsInfo.Cells(RECYCLE_ROW, vc), ApplicableText(), NotApplicableText(), "M34", "R34"
    Next ws
End Sub

Private Sub ReapplyTwoCell(ByVal ws As Worksheet, ByVal srcCell As Range, _
                           ByVal topLabel As String, ByVal bottomLabel As String, _
                           ByVal topCell As String, ByVal bottomCell As String)
    Dim cur As String
    cur = CommonNzText(srcCell.mergeArea.Cells(1, 1).value)
    If Len(Trim$(cur)) = 0 Then Exit Sub
    Dim topChecked As Boolean, bottomChecked As Boolean
    ParseSelection cur, topLabel, bottomLabel, topChecked, bottomChecked
    ApplyCheckGlyph ws.Range(topCell), topChecked
    ApplyCheckGlyph ws.Range(bottomCell), bottomChecked
End Sub

Private Sub ReapplyCombinedCell(ByVal ws As Worksheet, ByVal srcCell As Range, _
                                ByVal topLabel As String, ByVal bottomLabel As String, _
                                ByVal topCell As String, ByVal suffixText As String)
    Dim cur As String
    cur = CommonNzText(srcCell.mergeArea.Cells(1, 1).value)
    If Len(Trim$(cur)) = 0 Then Exit Sub
    Dim topChecked As Boolean, bottomChecked As Boolean
    ParseSelection cur, topLabel, bottomLabel, topChecked, bottomChecked
    ApplyCombinedCheckText ws.Range(topCell), topChecked, bottomChecked, topLabel, bottomLabel, suffixText
End Sub
