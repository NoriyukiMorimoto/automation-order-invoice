Option Explicit

' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽe?ｿｽ?ｿｽ?ｿｽv?ｿｽ?ｿｽ?ｿｽ[?ｿｽg?ｿｽe?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾖの奇ｿｽ{?ｿｽ?ｿｽ?ｿｽw?ｿｽb?ｿｽ_?ｿｽ[?ｿｽ]?ｿｽL?ｿｽB
' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾉ会ｿｽ?ｿｽ?ｿｽ?ｿｽA?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾌ転?ｿｽL?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ?ｿｽﾏ更?ｿｽ?ｿｽ?ｿｽﾉゑｿｽ?ｿｽﾄ転?ｿｽL?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ?ｿｽC?ｿｽu?ｿｽ?ｿｽ?ｿｽf)?ｿｽB
' ?ｿｽ]?ｿｽL?ｿｽﾎ象シ?ｿｽ[?ｿｽg?ｿｽﾌ追会ｿｽ?ｿｽ?ｿｽ ApplyVendorSheetHeaders ?ｿｽﾌデ?ｿｽB?ｿｽX?ｿｽp?ｿｽb?ｿｽ`?ｿｽﾖ趣ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾞ。
' ?ｿｽ?ｿｽ?ｿｽC?ｿｽ?ｿｽ?ｿｽ?ｿｽ: CHANGELOG.md ?ｿｽQ?ｿｽ?ｿｽ

' ?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽﾌヘ?ｿｽb?ｿｽ_?ｿｽ[?ｿｽ]?ｿｽL?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ(?ｿｽS?ｿｽﾐ具ｿｽ?ｿｽﾊ包ｿｽ)?ｿｽB?ｿｽu?ｿｽ?ｿｽ?ｿｽb?ｿｽN?ｿｽ?ｿｽ(16/27?ｿｽs?ｿｽ?ｿｽ)?ｿｽﾍ難ｿｽ?ｿｽI?ｿｽﾉ組?ｿｽﾝ暦ｿｽ?ｿｽﾄゑｿｽ
Private Const HEADER_SOURCE_COMMON_CELLS As String = "B6,C6,C2,C9,C10,C13,C15:C16,F6"
Private Const HEADER_DATE_FONT_SIZE As Double = 14#
' ?ｿｽ?克ﾒ用?ｿｽV?ｿｽ[?ｿｽg?ｿｽ]?ｿｽL?ｿｽﾅ参?ｿｽﾆゑｿｽ?ｿｽ?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN?ｿｽs(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ)
Private Const CONTRACTOR_CONTRACT_AMOUNT_ROW As Long = 33
Private Const CONTRACTOR_CONSUMPTION_TAX_ROW As Long = 34
Private Const CONTRACTOR_CONTRACT_TOTAL_ROW As Long = 35
' ?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ ?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN: ?ｿｽ?ｿｽ\?ｿｽﾒ厄ｿｽ(12?ｿｽs?ｿｽ?ｿｽ)?ｿｽE?ｿｽZ?ｿｽ?ｿｽ(14?ｿｽs?ｿｽ?ｿｽ)
Private Const CONTRACTOR_REPRESENTATIVE_ROW As Long = 12
Private Const CONTRACTOR_ADDRESS_ROW As Long = 14
Private Const CONDITION_CHECKBOX_D_COL As Long = 4
Private Const CONDITION_CHECKBOX_X_COL As Long = 24
Private Const CONDITION_CHECKBOX_E_COL As Long = 5
Private Const CONDITION_CHECKBOX_LEFT_MAX_COL As Long = 12
Private Const CONDITION_CHECKBOX_RIGHT_MIN_COL As Long = 20
Private Const CONDITION_ROW38_BAND_MIN_ROW As Long = 37
Private Const CONDITION_ROW38_BAND_MAX_ROW As Long = 39
Private Const CONDITION_E_PAIR_MIN_ROW As Long = 34
Private Const CONDITION_E_PAIR_MAX_ROW As Long = 35

' ?ｿｽw?ｿｽ?ｿｽu?ｿｽ?ｿｽ?ｿｽb?ｿｽN?ｿｽﾌ施?ｿｽH?ｿｽ?ｿｽﾐに対会ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽe?ｿｽ?ｿｽ?ｿｽv?ｿｽ?ｿｽ?ｿｽ[?ｿｽg5?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾖヘ?ｿｽb?ｿｽ_?ｿｽ[?ｿｽ?ｿｽ]?ｿｽL?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽf?ｿｽB?ｿｽX?ｿｽp?ｿｽb?ｿｽ`?ｿｽ?ｿｽ)
Public Sub ApplyVendorSheetHeaders(ByVal wsInfo As Worksheet, _
                                   ByVal vendorIndex As Long, _
                                   ByVal aliasText As String)
    If wsInfo Is Nothing Then Exit Sub
    If Len(aliasText) = 0 Then Exit Sub

    Dim baseNames As Variant
    baseNames = mod_OrderTpl_Shared.OrderTplTemplateSheetBaseNames()

    Dim i As Long
    For i = LBound(baseNames) To UBound(baseNames)
        Dim sheetName As String
        sheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName(CStr(baseNames(i)), aliasText)
        If mod_OrderTpl_Shared.OrderTplSheetExists(sheetName) Then
            Dim wsTarget As Worksheet
            Set wsTarget = ThisWorkbook.Worksheets(sheetName)

            ApplyVendorSheetTabColor wsInfo, wsTarget, vendorIndex

            Select Case i - LBound(baseNames)
                Case 0: ApplyBreakdownHeader wsInfo, wsTarget, vendorIndex
                Case 1: ApplyContractorHeader wsInfo, wsTarget, vendorIndex
                Case 2: ApplyAcceptanceHeader wsInfo, wsTarget, vendorIndex
                Case 3: ApplyBranchCopyHeader wsInfo, wsTarget, vendorIndex
                Case 4: ApplyAttachment3Header wsInfo, wsTarget, vendorIndex
            End Select
        End If
    Next i

    Dim condSheetName As String
    condSheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
        mod_OrderTpl_Shared.OrderTplBaseNameConditionText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(condSheetName) Then
        ApplyVendorSheetTabColor wsInfo, ThisWorkbook.Worksheets(condSheetName), vendorIndex
        SetupConditionCheckboxExclusivity ThisWorkbook.Worksheets(condSheetName)
    End If
End Sub

' ?ｿｽS?ｿｽm?ｿｽ?ｿｽ?ｿｽﾐのテ?ｿｽ?ｿｽ?ｿｽv?ｿｽ?ｿｽ?ｿｽ[?ｿｽg?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾖヘ?ｿｽb?ｿｽ_?ｿｽ[?ｿｽ?ｿｽ?ｿｽﾄ転?ｿｽL?ｿｽ?ｿｽ?ｿｽ?ｿｽ
Public Sub RefreshAllVendorSheetHeaders(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo Quiet

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))

    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim companyName As String
        companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, vendorIndex)
        If companyName <> "" Then
            Dim vendorName As String
            Dim aliasText As String
            Dim workText As String
            If mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, vendorName, aliasText, workText) Then
                ApplyVendorSheetHeaders wsInfo, vendorIndex, aliasText

                Dim condSheetName As String
                condSheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
                    mod_OrderTpl_Shared.OrderTplBaseNameConditionText(), aliasText)
                If mod_OrderTpl_Shared.OrderTplSheetExists(condSheetName) Then
                    ApplyConditionSheetHeader wsInfo, ThisWorkbook.Worksheets(condSheetName), vendorIndex
                    SetupConditionCheckboxExclusivity ThisWorkbook.Worksheets(condSheetName)
                End If
            End If
        End If
    Next vendorIndex
    Exit Sub

Quiet:
    mod_OrderTpl_Shared.OrderTplLog "RefreshAllVendorSheetHeaders error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' Sheet1(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ)?ｿｽ?ｿｽWorksheet_Change?ｿｽ?ｿｽ?ｿｽ?ｿｽﾄばゑｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽB
' ?ｿｽw?ｿｽb?ｿｽ_?ｿｽ[?ｿｽ]?ｿｽL?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ(B6/C6/C2/C9/C10/C15:C16/F6?ｿｽA?ｿｽe?ｿｽu?ｿｽ?ｿｽ?ｿｽb?ｿｽN?ｿｽ?ｿｽ16/27?ｿｽs?ｿｽ?ｿｽ)?ｿｽﾌ変更?ｿｽ?ｿｽ?ｿｽe?ｿｽﾐシ?ｿｽ[?ｿｽg?ｿｽﾖ費ｿｽ?ｿｽf?ｿｽ?ｿｽ?ｿｽ?ｿｽ
Public Sub HandleBasicInfoHeaderSourceChange(ByVal wsInfo As Worksheet, ByVal target As Range)
    If wsInfo Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error GoTo Quiet

    Dim sourceRange As Range
    Set sourceRange = BuildHeaderSourceRange(wsInfo)
    If sourceRange Is Nothing Then Exit Sub
    If Intersect(target, sourceRange) Is Nothing Then Exit Sub

    RefreshAllVendorSheetHeaders wsInfo
    Exit Sub

Quiet:
    Err.Clear
End Sub

' ?ｿｽw?ｿｽb?ｿｽ_?ｿｽ[?ｿｽ]?ｿｽL?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ?ｿｽﾌ監趣ｿｽ?ｿｽﾍ囲ゑｿｽﾔゑｿｽ?ｿｽ?ｿｽ?ｿｽJ?ｿｽ?ｿｽ?ｿｽb?ｿｽp?ｿｽ[(Sheet1?ｿｽﾌ変更?ｿｽQ?ｿｽ[?ｿｽg?ｿｽ\?ｿｽz?ｿｽp)
Public Function GetBasicInfoHeaderSourceMonitorRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function
    Set GetBasicInfoHeaderSourceMonitorRange = BuildHeaderSourceRange(wsInfo)
End Function

' ?ｿｽw?ｿｽb?ｿｽ_?ｿｽ[?ｿｽ]?ｿｽL?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ?ｿｽﾌ監趣ｿｽ?ｿｽﾍ茨ｿｽ(?ｿｽ?ｿｽ?ｿｽﾊセ?ｿｽ?ｿｽ + ?ｿｽe?ｿｽu?ｿｽ?ｿｽ?ｿｽb?ｿｽN?ｿｽﾌ業者コ?ｿｽ[?ｿｽh16?ｿｽs?ｿｽ?ｿｽ/?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾔ搾ｿｽ27?ｿｽs?ｿｽ?ｿｽ)
Private Function BuildHeaderSourceRange(ByVal wsInfo As Worksheet) As Range
    Dim result As Range
    Set result = wsInfo.Range(HEADER_SOURCE_COMMON_CELLS)

    Dim vendorCount As Long
    vendorCount = mod_Construction_BasicTotals.GetBasicInfoVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)
        Set result = Union(result, _
                           wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, valueColumn), _
                           wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn), _
                           wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn), _
                           wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn), _
                           wsInfo.Cells(38, valueColumn), _
                           wsInfo.Cells(39, valueColumn), _
                           wsInfo.Cells(40, valueColumn), _
                           wsInfo.Cells(41, valueColumn), _
                           wsInfo.Cells(42, valueColumn))
    Next vendorIndex

    Set BuildHeaderSourceRange = result
End Function

' ?ｿｽV?ｿｽ[?ｿｽg?ｿｽ?ｿｽ?ｿｽo?ｿｽ?ｿｽ(?ｿｽ^?ｿｽu)?ｿｽF: ?ｿｽH?ｿｽ?ｿｽ?ｿｽ謨ｪ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ10?ｿｽs?ｿｽ?ｿｽ)?ｿｽZ?ｿｽ?ｿｽ?ｿｽﾌ塗?ｿｽ?ｿｽﾂぶゑｿｽ?ｿｽF?ｿｽ?ｿｽK?ｿｽp?ｿｽ?ｿｽ?ｿｽ?ｿｽ
Private Sub ApplyVendorSheetTabColor(ByVal wsInfo As Worksheet, _
                                     ByVal wsTarget As Worksheet, _
                                     ByVal vendorIndex As Long)
    On Error Resume Next
    Dim sourceCell As Range
    Set sourceCell = wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, _
                                  mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex))
    If sourceCell.Interior.ColorIndex = xlColorIndexNone Then
        wsTarget.Tab.ColorIndex = xlColorIndexNone
    Else
        wsTarget.Tab.Color = sourceCell.Interior.Color
    End If
    On Error GoTo 0
End Sub


' ?ｿｽ?ｿｽ?ｿｽ?ｾ細ヘ?ｿｽb?ｿｽ_?ｿｽ[?ｿｽ?ｿｽ?ｿｽﾖ奇ｿｽ{?ｿｽ?ｿｽ?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾌ難ｿｽ?ｿｽe?ｿｽ?ｿｽ]?ｿｽL?ｿｽ?ｿｽ?ｿｽ?ｿｽ
Public Sub ApplyBreakdownHeader(ByVal wsInfo As Worksheet, _
                                ByVal wsBreakdown As Worksheet, _
                                ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsBreakdown Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' C2: ?ｿｽ?ｿｽ?ｿｽX?ｿｽR?ｿｽ[?ｿｽh(?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ_?ｿｽP?ｿｽ?ｿｽ?ｿｽK?ｿｽp?ｿｽ?ｿｽ?ｿｽ?ｿｽﾌ単?ｿｽ?ｿｽ?ｿｽK?ｿｽp?ｿｽ?ｿｽ?ｿｽ?ｿｽV?ｿｽ[?ｿｽg B/C?ｿｽ?ｿｽﾆ搾ｿｽ ?ｿｽ?ｿｽ G?ｿｽ?ｿｽ)
    Dim branchOfficeCode As String
    branchOfficeCode = mod_OrderTpl_Shared.OrderTplResolveBranchOfficeCode( _
        CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value), _
        CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    WriteHeaderText wsBreakdown.Range("C2"), branchOfficeCode, False

    ' C3: ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾔ搾ｿｽ(?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN27?ｿｽs?ｿｽ?ｿｽ)
    WriteHeaderValue wsBreakdown.Range("C3"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, False

    ' B6: ?ｿｽH?ｿｽ?ｿｽ?ｿｽﾔ搾ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC9)
    WriteHeaderValue wsBreakdown.Range("B6"), wsInfo.Range("C9").value, False

    ' D6: ?ｿｽH?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC10?ｿｽA?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ?ｿｽﾌゑｿｽ?ｿｽﾟ抵ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ)
    WriteHeaderValue wsBreakdown.Range("D6"), wsInfo.Range("C10").value, True

    ' L5/L6: ?ｿｽH?ｿｽ?ｿｽ ?ｿｽ?ｿｽ/?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC15/C16?ｿｽA?ｿｽa?ｿｽ?ｿｽ\?ｿｽ?ｿｽ)
    WriteHeaderDate wsBreakdown.Range("L5"), wsInfo.Range("C15").value
    WriteHeaderDate wsBreakdown.Range("L6"), wsInfo.Range("C16").value
    wsBreakdown.Range("L5:L6").Font.Size = HEADER_DATE_FONT_SIZE

    ' O2: ?ｿｽ?成?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC2?ｿｽA?ｿｽa?ｿｽ?ｿｽ\?ｿｽ?ｿｽ)
    ' P2:Q2 created date from C2. O2 keeps template label.
    WriteHeaderDate wsBreakdown.Range("P2"), wsInfo.Range("C2").value
    wsBreakdown.Range("P2").MergeArea.Font.Size = HEADER_DATE_FONT_SIZE

    ' O3: ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽF6)
    ' P3:Q3 branch chief from F6. O3 keeps template label.
    WriteHeaderValue wsBreakdown.Range("P3"), wsInfo.Range("F6").value, True

    ' P5: ?ｿｽO?ｿｽ?ｿｽ?ｿｽ?ｿｽﾐ厄ｿｽ(?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN11?ｿｽs?ｿｽ?ｿｽ)
    WriteHeaderValue wsBreakdown.Range("P5"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' P6: ?ｿｽﾆ者コ?ｿｽ[?ｿｽh(?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN16?ｿｽs?ｿｽ?ｿｽ)
    WriteHeaderValue wsBreakdown.Range("P6"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽy?ｿｽ[?ｿｽW?ｿｽ?ｿｽ?ｿｽ?ｿｽp?ｿｽﾉヘ?ｿｽb?ｿｽ_?ｿｽ[?ｿｽs(7:10?ｿｽs?ｿｽ?ｿｽ)?ｿｽ?ｿｽ?ｿｽ^?ｿｽC?ｿｽg?ｿｽ?ｿｽ?ｿｽs?ｿｽﾉ設定す?ｿｽ?ｿｽ
    On Error Resume Next
    wsBreakdown.PageSetup.PrintTitleRows = ORDER_TPL_PRINT_TITLE_ROWS
    On Error GoTo ErrorHandler

    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownHeader done: " & wsBreakdown.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' Rail-work type text for H25:U25 copy (rail condition sheet only)
Private Function ConditionRailWorkTypeText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H8ECC, &H9053, &H5DE5, &H4E8B)
    End If
    ConditionRailWorkTypeText = cached
End Function

' Copy basic info / vendor fields into condition sheet.
' Ten shared fields + rail sheet H25:U25 from vendor row41.
Public Sub ApplyConditionSheetHeader(ByVal wsInfo As Worksheet, _
                                     ByVal wsCondition As Worksheet, _
                                     ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsCondition Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' S1:X1 created date from C2 (Gregorian yyyy/m/d, centered)
    WriteHeaderDateGregorian wsCondition.Range("S1"), wsInfo.Range("C2").value
    wsCondition.Range("S1").MergeArea.Cells(1, 1).HorizontalAlignment = xlCenter

    ' P3:S3 branch B6 / T3:X3 office C6 / T4:X4 chief F6 (bottom aligned)
    WriteHeaderValueBottom wsCondition.Range("P3"), wsInfo.Range("B6").value
    WriteHeaderValueBottom wsCondition.Range("T3"), wsInfo.Range("C6").value
    WriteHeaderValueBottom wsCondition.Range("T4"), wsInfo.Range("F6").value

    ' B5:C6 vendor name row11 (bottom aligned)
    WriteHeaderValueBottom wsCondition.Range("B5"), _
        wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value

    ' F9:I9 project no C9 / L9:X9 project name C10 (centered)
    WriteHeaderValue wsCondition.Range("F9"), wsInfo.Range("C9").value, True
    WriteHeaderValue wsCondition.Range("L9"), wsInfo.Range("C10").value, True

    ' F11:L11 start C15 / R11:X11 end C16 (Gregorian yyyy/m/d)
    WriteHeaderDateGregorian wsCondition.Range("F11"), wsInfo.Range("C15").value
    WriteHeaderDateGregorian wsCondition.Range("R11"), wsInfo.Range("C16").value

    ' D16:X16 site C13 (centered)
    WriteHeaderValue wsCondition.Range("D16"), wsInfo.Range("C13").value, True

    ' Rail-work condition sheet only: H25:U25 <- vendor row41 (left aligned)
    Dim workType As String
    workType = CommonNormalizeText(CommonNzText( _
        wsInfo.Cells(BASIC_INFO_VENDOR_WORK_TYPE_ROW, valueColumn).value))
    If StrComp(workType, ConditionRailWorkTypeText(), vbTextCompare) = 0 Then
        WriteHeaderValueLeft wsCondition.Range("H25"), wsInfo.Cells(41, valueColumn).value
    End If

    mod_OrderTpl_Shared.OrderTplLog "ApplyConditionSheetHeader done: " & wsCondition.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyConditionSheetHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' ===== Condition sheet checkbox exclusivity =====
' D/X rows (10,18-33,38) and E34/E35 pairs keep exactly one side ON.
' Turning one ON turns the other OFF; turning one OFF turns the other ON.


' Exclusive target rows (D/X: 10,18-33,38 / E vertical pair: 34,35)
Private Function IsConditionDxExclusiveRow(ByVal rowIndex As Long) As Boolean
    IsConditionDxExclusiveRow = (rowIndex = 10) Or (rowIndex >= 18 And rowIndex <= 33) Or (rowIndex = 38)
End Function

Private Function IsConditionEVerticalPairRow(ByVal rowIndex As Long) As Boolean
    IsConditionEVerticalPairRow = (rowIndex = 34) Or (rowIndex = 35)
End Function

Private Function IsConditionDxLeftSideColumn(ByVal colIndex As Long) As Boolean
    IsConditionDxLeftSideColumn = (colIndex > 0 And colIndex <= CONDITION_CHECKBOX_LEFT_MAX_COL And _
                                   colIndex <> CONDITION_CHECKBOX_E_COL)
End Function

Private Function IsConditionDxRightSideColumn(ByVal colIndex As Long) As Boolean
    IsConditionDxRightSideColumn = (colIndex >= CONDITION_CHECKBOX_RIGHT_MIN_COL)
End Function

Private Function IsConditionDxSideCheckbox(ByVal cb As Object) As Boolean
    Dim colIndex As Long
    colIndex = cb.TopLeftCell.Column
    IsConditionDxSideCheckbox = IsConditionDxLeftSideColumn(colIndex) Or _
                                IsConditionDxRightSideColumn(colIndex)
End Function

Private Function ConditionCheckboxOverlapsCell(ByVal cb As Object, _
                                               ByVal ws As Worksheet, _
                                               ByVal rowIndex As Long, _
                                               ByVal colIndex As Long) As Boolean
    On Error GoTo Fail
    Dim cell As Range
    Set cell = ws.Cells(rowIndex, colIndex)
    Dim centerX As Double
    Dim centerY As Double
    centerX = cb.Left + (cb.Width / 2#)
    centerY = cb.Top + (cb.Height / 2#)
    ConditionCheckboxOverlapsCell = (centerX >= cell.Left And centerX < (cell.Left + cell.Width) And _
                                     centerY >= cell.Top And centerY < (cell.Top + cell.Height))
    Exit Function
Fail:
    ConditionCheckboxOverlapsCell = False
End Function

Private Function ConditionEVerticalCheckboxRow(ByVal cb As Object, ByVal ws As Worksheet) As Long
    If ConditionCheckboxOverlapsCell(cb, ws, CONDITION_E_PAIR_MAX_ROW, CONDITION_CHECKBOX_E_COL) Then
        ConditionEVerticalCheckboxRow = CONDITION_E_PAIR_MAX_ROW
        Exit Function
    End If
    If ConditionCheckboxOverlapsCell(cb, ws, CONDITION_E_PAIR_MIN_ROW, CONDITION_CHECKBOX_E_COL) Then
        ConditionEVerticalCheckboxRow = CONDITION_E_PAIR_MIN_ROW
        Exit Function
    End If
    If cb.TopLeftCell.Column = CONDITION_CHECKBOX_E_COL Then
        If cb.TopLeftCell.Row >= CONDITION_E_PAIR_MIN_ROW And _
           cb.TopLeftCell.Row <= CONDITION_E_PAIR_MAX_ROW Then
            ConditionEVerticalCheckboxRow = cb.TopLeftCell.Row
        End If
    End If
End Function

Private Function IsConditionEVerticalCheckbox(ByVal cb As Object, ByVal ws As Worksheet) As Boolean
    IsConditionEVerticalCheckbox = (ConditionEVerticalCheckboxRow(cb, ws) > 0)
End Function

Private Function IsConditionRow38BandRow(ByVal rowIndex As Long) As Boolean
    IsConditionRow38BandRow = (rowIndex >= CONDITION_ROW38_BAND_MIN_ROW And _
                               rowIndex <= CONDITION_ROW38_BAND_MAX_ROW)
End Function

Private Function IsConditionExclusiveCheckbox(ByVal cb As Object, ByVal ws As Worksheet) As Boolean
    Dim rowIndex As Long
    rowIndex = cb.TopLeftCell.Row

    If IsConditionEVerticalCheckbox(cb, ws) Then
        IsConditionExclusiveCheckbox = True
        Exit Function
    End If

    If Not IsConditionDxSideCheckbox(cb) Then Exit Function

    If IsConditionDxExclusiveRow(rowIndex) Then
        IsConditionExclusiveCheckbox = True
    ElseIf IsConditionRow38BandRow(rowIndex) Then
        ' Row-38 controls may anchor to rows 37-39.
        IsConditionExclusiveCheckbox = True
    End If
End Function

' Exclusive target rows (D/X: 10,18-33,38 / E vertical pair: 34,35)
Private Function IsConditionExclusiveRow(ByVal r As Long) As Boolean
    IsConditionExclusiveRow = IsConditionDxExclusiveRow(r) Or IsConditionEVerticalPairRow(r)
End Function

' Assign exclusive-click macro to condition-sheet checkboxes
Public Sub SetupConditionCheckboxExclusivity(ByVal wsCondition As Worksheet)
    If wsCondition Is Nothing Then Exit Sub
    FixConditionE34E35Pair wsCondition
    On Error Resume Next
    Dim cb As Object
    For Each cb In wsCondition.CheckBoxes
        If IsConditionExclusiveCheckbox(cb, wsCondition) Then
            cb.OnAction = "'" & ThisWorkbook.Name & "'!ConditionCheckboxExclusiveClick"
        End If
    Next cb
    On Error GoTo 0
    NormalizeConditionCheckboxPairs wsCondition
    FixConditionE34E35Pair wsCondition
End Sub

' Checkbox click handler: flip paired control to opposite state.
Public Sub ConditionCheckboxExclusiveClick()
    On Error GoTo Done
    Dim callerName As String
    callerName = Application.Caller

    Dim clicked As Object
    Dim ws As Worksheet
    Set clicked = Nothing
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set clicked = ws.CheckBoxes(callerName)
        On Error GoTo Done
        If Not clicked Is Nothing Then Exit For
        Set clicked = Nothing
    Next ws
    If clicked Is Nothing Then Exit Sub
    If ws Is Nothing Then Exit Sub

    Dim pair As Object
    Set pair = FindConditionCheckboxPair(ws, clicked)
    If pair Is Nothing Then Exit Sub

    Application.EnableEvents = False
    If clicked.Value = xlOn Then
        pair.Value = xlOff
    Else
        pair.Value = xlOn
    End If
    Application.EnableEvents = True

Done:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' Return paired checkbox (D/X: opposite col + nearest Top / E34-E35: other row)
Private Function FindConditionCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim rowIndex As Long
    rowIndex = clicked.TopLeftCell.Row

    If IsConditionEVerticalCheckbox(clicked, ws) Then
        Set FindConditionCheckboxPair = FindConditionEVerticalCheckboxPair(ws, clicked)
        Exit Function
    End If

    If IsConditionDxSideCheckbox(clicked) Then
        If IsConditionDxExclusiveRow(rowIndex) Or IsConditionRow38BandRow(rowIndex) Then
            Set FindConditionCheckboxPair = FindConditionDxCheckboxPair(ws, clicked)
        End If
    End If
End Function

Private Function FindConditionEVerticalCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim clickedRow As Long
    clickedRow = ConditionEVerticalCheckboxRow(clicked, ws)
    If clickedRow = 0 Then Exit Function

    Dim targetRow As Long
    If clickedRow = CONDITION_E_PAIR_MIN_ROW Then
        targetRow = CONDITION_E_PAIR_MAX_ROW
    Else
        targetRow = CONDITION_E_PAIR_MIN_ROW
    End If

    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If cb.Name <> clicked.Name Then
            If ConditionEVerticalCheckboxRow(cb, ws) = targetRow Then
                Set FindConditionEVerticalCheckboxPair = cb
                Exit Function
            End If
        End If
    Next cb
End Function

Private Function FindConditionDxCheckboxPair(ByVal ws As Worksheet, ByVal clicked As Object) As Object
    Dim clickedCol As Long
    Dim clickedRow As Long
    Dim clickedTop As Double
    clickedCol = clicked.TopLeftCell.Column
    clickedRow = clicked.TopLeftCell.Row
    clickedTop = clicked.Top

    Dim clickedIsLeft As Boolean
    If IsConditionDxLeftSideColumn(clickedCol) Then
        clickedIsLeft = True
    ElseIf IsConditionDxRightSideColumn(clickedCol) Then
        clickedIsLeft = False
    Else
        Exit Function
    End If

    Dim rowMin As Long
    Dim rowMax As Long
    If clickedRow = 38 Or clickedRow = 39 Then
        rowMin = CONDITION_ROW38_BAND_MIN_ROW
        rowMax = CONDITION_ROW38_BAND_MAX_ROW
    Else
        rowMin = clickedRow
        rowMax = clickedRow
    End If

    Dim bestCb As Object
    Dim bestTopDist As Double
    bestTopDist = -1

    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If cb.Name <> clicked.Name Then
            If cb.TopLeftCell.Row >= rowMin And cb.TopLeftCell.Row <= rowMax Then
                Dim cbCol As Long
                cbCol = cb.TopLeftCell.Column
                Dim cbIsLeft As Boolean
                Dim hasSide As Boolean
                hasSide = False
                If IsConditionDxLeftSideColumn(cbCol) Then
                    cbIsLeft = True
                    hasSide = True
                ElseIf IsConditionDxRightSideColumn(cbCol) Then
                    cbIsLeft = False
                    hasSide = True
                End If

                If hasSide And (cbIsLeft <> clickedIsLeft) Then
                    Dim topDist As Double
                    topDist = Abs(cb.Top - clickedTop)
                    If bestCb Is Nothing Or topDist < bestTopDist Then
                        Set bestCb = cb
                        bestTopDist = topDist
                    End If
                End If
            End If
        End If
    Next cb

    Set FindConditionDxCheckboxPair = bestCb
End Function

Private Sub NormalizeConditionCheckboxPairs(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Dim processed As Object
    Set processed = CreateObject("Scripting.Dictionary")
    processed.CompareMode = vbTextCompare

    Application.EnableEvents = False
    Dim cb As Object
    For Each cb In ws.CheckBoxes
        If IsConditionExclusiveCheckbox(cb, ws) Then
            If Not processed.Exists(cb.Name) Then
                Dim pair As Object
                Set pair = FindConditionCheckboxPair(ws, cb)
                If Not pair Is Nothing Then
                    processed.Add cb.Name, True
                    processed.Add pair.Name, True

                    If cb.Value = xlOn And pair.Value = xlOn Then
                        pair.Value = xlOff
                    ElseIf cb.Value <> xlOn And pair.Value <> xlOn Then
                        If IsConditionEVerticalCheckbox(cb, ws) Or _
                           IsConditionDxLeftSideColumn(cb.TopLeftCell.Column) Then
                            cb.Value = xlOn
                        Else
                            pair.Value = xlOn
                        End If
                    End If
                End If
            End If
        End If
    Next cb
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Sub FixConditionE34E35Pair(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim cb34 As Object
    Dim cb35 As Object
    Dim cb As Object
    For Each cb In ws.CheckBoxes
        Select Case ConditionEVerticalCheckboxRow(cb, ws)
            Case CONDITION_E_PAIR_MIN_ROW
                Set cb34 = cb
            Case CONDITION_E_PAIR_MAX_ROW
                Set cb35 = cb
        End Select
    Next cb
    If cb34 Is Nothing Or cb35 Is Nothing Then Exit Sub

    On Error Resume Next
    cb34.OnAction = "'" & ThisWorkbook.Name & "'!ConditionCheckboxExclusiveClick"
    cb35.OnAction = "'" & ThisWorkbook.Name & "'!ConditionCheckboxExclusiveClick"
    On Error GoTo 0

    Application.EnableEvents = False
    If cb34.Value = xlOn And cb35.Value = xlOn Then
        cb35.Value = xlOff
    ElseIf cb34.Value <> xlOn And cb35.Value <> xlOn Then
        cb34.Value = xlOn
    End If
    Application.EnableEvents = True
End Sub

' ?ｿｽ?克ﾒ用/?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ/?ｿｽx?ｿｽX?ｿｽT ?ｿｽ?ｿｽ?ｿｽ?ｿｽ: S1:U1 ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾔ搾ｿｽ(27) / Q2:V2 ?ｿｽ?成?ｿｽ?ｿｽC2(?ｿｽ?ｿｽ?ｿｽ?ｿｽ) /
'   ?ｿｽs20-34(E20:?ｿｽH?ｿｽ?ｿｽ?ｿｽ?ｿｽC10 E22:?ｿｽs?ｿｽ?ｿｽ?ｿｽ{?ｿｽ?ｿｽC13 G24:?ｿｽH?ｿｽ?ｿｽ?ｿｽ?ｿｽC15(?ｿｽ?ｿｽ?ｿｽ?ｿｽ) G26:?ｿｽH?ｿｽ?ｿｽ?ｿｽ?ｿｽC16(?ｿｽ?ｿｽ?ｿｽ?ｿｽ)?ｿｽA
'   Q22:?ｿｽﾅ搾ｿｽ(35) Q23:?ｿｽﾅ費ｿｽ(33) Q24:?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ(34)?ｿｽAC30/H30/J30/M34/R34/F32/F33 ?ｿｽ?ｿｽ Reapply?ｿｽn?ｿｽo?ｿｽR)?ｿｽ?ｿｽ
'   ApplyContractorStyleCommonFields ?ｿｽﾅ、3?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾆゑｿｽ?ｿｽ?ｿｽ?ｿｽ黹搾ｿｽW?ｿｽb?ｿｽN?ｿｽ?ｿｽK?ｿｽp?ｿｽ?ｿｽ?ｿｽ?ｿｽB
' ?ｿｽ?克ﾒ用?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾅ有: E9:?ｿｽﾆ者コ?ｿｽ[?ｿｽh(16) A13:?ｿｽ?ｿｽﾐ厄ｿｽ(11) M10:?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾒ住?ｿｽ?ｿｽ
'   M11:?ｿｽ?ｿｽS?ｿｽH?ｿｽﾆ奇ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ+?ｿｽ誧ｲ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ M12:?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽX?ｿｽg?ｿｽQ?ｿｽ?ｿｽ)
Private Sub ApplyContractorHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    ' E9: ?ｿｽﾆ者コ?ｿｽ[?ｿｽh(?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN16?ｿｽs), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("E9"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' A13: ?ｿｽ?ｿｽﾐ厄ｿｽ(?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN11?ｿｽs), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("A13"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' M10/M11/M12: ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽX?ｿｽg?ｿｽQ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ ?ｿｽZ?ｿｽ?ｿｽ/?ｿｽ?ｿｽ?ｿｽ?ｿｽ/?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ)
    ApplyOfficeChiefBlock wsInfo, wsTarget

    ' ?ｿｽs38-42(?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ/?ｿｽJ?ｿｽ?ｿｽ/?ｿｽx?ｿｽ?ｿｽ?ｿｽﾞ暦ｿｽ/?ｿｽﾝ与?ｿｽi/?ｿｽ?ｿｽ?ｿｽT?ｿｽC?ｿｽN?ｿｽ?ｿｽ)?ｿｽ?ｿｽ?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｩゑｿｽ
    ' ?ｿｽ?克ﾒ用?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽE?ｿｽx?ｿｽX?ｿｽT?ｿｽﾖ再転?ｿｽL(Reapply?ｿｽ?ｿｽ?ｿｽﾅ全?ｿｽﾄの撰ｿｽ?ｿｽ?ｿｽ?ｿｽﾏシ?ｿｽ[?ｿｽg?ｿｽﾖ適?ｿｽp)
    mod_BasicInfoExclusiveChoice.ReapplyExclusiveChoices wsInfo, vendorIndex
    mod_BasicInfoSupplyLoan.ReapplySupplyLoan wsInfo, vendorIndex

    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' ?ｿｽ?克ﾒ用/?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ/?ｿｽx?ｿｽX?ｿｽT?ｿｽﾅ具ｿｽ?ｿｽﾊの転?ｿｽL(S1:U1 ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾔ搾ｿｽ / Q2:V2 ?ｿｽ?成?ｿｽ?ｿｽ(?ｿｽ?ｿｽ?ｿｽ?ｿｽ) /
' ?ｿｽs20-34: E20 ?ｿｽH?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽEE22 ?ｿｽs?ｿｽ?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽEG24/G26 ?ｿｽH?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ?ｿｽ?ｿｽ)?ｿｽEQ22-24 ?ｿｽ?ｿｽ?ｿｽ?ｾ細参?ｿｽ?ｿｽ)?ｿｽB
' ?ｿｽA3?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾆゑｿｽ?ｿｽZ?ｿｽ?ｿｽ?ｿｽ\?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽS?ｿｽ?ｿｽv?ｿｽ?ｿｽ?ｿｽﾄゑｿｽ?ｿｽ驍ｽ?ｿｽﾟ、?ｿｽ?ｿｽ?ｿｽ黹搾ｿｽW?ｿｽb?ｿｽN?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ黷ｼ?ｿｽ?ｿｽK?ｿｽp?ｿｽ?ｿｽ?ｿｽ?ｿｽB
Private Sub ApplyContractorStyleCommonFields(ByVal wsInfo As Worksheet, _
                                             ByVal wsTarget As Worksheet, _
                                             ByVal valueColumn As Long)
    ' S1:U1: ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾔ搾ｿｽ(?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN27?ｿｽs), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("S1"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, True

    ' Q2:V2: ?ｿｽ?成?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC2, ?ｿｽ?ｿｽ?ｿｽ?ｿｽ), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderDateGregorian wsTarget.Range("Q2"), wsInfo.Range("C2").value

    ' E20: ?ｿｽH?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC10), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("E20"), wsInfo.Range("C10").value, True

    ' E22: ?ｿｽs?ｿｽ?ｿｽ?ｿｽ{?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC13), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("E22"), wsInfo.Range("C13").value, True

    ' G24: ?ｿｽH?ｿｽ?ｿｽ ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC15, ?ｿｽ?ｿｽ?ｿｽ?ｿｽ), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderDateGregorian wsTarget.Range("G24"), wsInfo.Range("C15").value

    ' G26: ?ｿｽH?ｿｽ?ｿｽ ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽC16, ?ｿｽ?ｿｽ?ｿｽ?ｿｽ), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderDateGregorian wsTarget.Range("G26"), wsInfo.Range("C16").value

    ' Q22/Q23/Q24: ?ｿｽ?ｿｽ?ｿｽ?ｿｽO?ｿｽ?ｿｽ?ｿｽ[?ｿｽv?ｿｽﾌ難ｿｽ?ｿｽ?ｾ細シ?ｿｽ[?ｿｽg ?ｿｽ?ｿｽ?ｿｽv/?ｿｽv/?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ ?ｿｽs(Q?ｿｽ?ｿｽ)?ｿｽ?ｿｽ?ｿｽQ?ｿｽﾆゑｿｽ?ｿｽ髏ｶ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    ' (?ｿｽ?ｿｽ?ｿｽ?ｾ細の値?ｿｽ?ｿｽ?ｿｽX?ｿｽV?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾆ趣ｿｽ?ｿｽ?ｿｽ?ｿｽﾅ更?ｿｽV?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ)
    Dim breakdownName As String
    breakdownName = ResolveBreakdownSheetNameFromTarget(wsTarget)
    If Len(breakdownName) > 0 Then
        WriteHeaderFormulaRight wsTarget.Range("Q22"), BuildBreakdownQFormula(breakdownName, ContractorGrandTotalLabelText())
        WriteHeaderFormulaRight wsTarget.Range("Q23"), BuildBreakdownQFormula(breakdownName, ContractorNetTotalLabelText())
        WriteHeaderFormulaRight wsTarget.Range("Q24"), BuildBreakdownQFormula(breakdownName, ContractorTaxLabelText())
    Else
        WriteHeaderValueRight wsTarget.Range("Q22"), wsInfo.Cells(CONTRACTOR_CONTRACT_TOTAL_ROW, valueColumn).value
        WriteHeaderValueRight wsTarget.Range("Q23"), wsInfo.Cells(CONTRACTOR_CONTRACT_AMOUNT_ROW, valueColumn).value
        WriteHeaderValueRight wsTarget.Range("Q24"), wsInfo.Cells(CONTRACTOR_CONSUMPTION_TAX_ROW, valueColumn).value
    End If
End Sub

' wsTarget(?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾏテ?ｿｽ?ｿｽ?ｿｽv?ｿｽ?ｿｽ?ｿｽ[?ｿｽg?ｿｽV?ｿｽ[?ｿｽg)?ｿｽﾆ難ｿｽ?ｿｽ?ｿｽO?ｿｽ?ｿｽ?ｿｽ[?ｿｽv?ｿｽﾌ受注者用(?ｿｽ?ｿｽ?ｿｽ?ｿｽ)?ｿｽV?ｿｽ[?ｿｽg?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
Private Function ResolveContractorSheetFromTarget(ByVal wsTarget As Worksheet) As Worksheet
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameContractorText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then Set ResolveContractorSheetFromTarget = ThisWorkbook.Worksheets(nm)
End Function

' ?ｿｽ?克ﾒ用?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾌ鯉ｿｽ?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ?ｿｽ謫ｾ?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ?ｿｽﾝゑｿｽ?ｿｽﾈゑｿｽ/?ｿｽ?ｿｽﾈゑｿｽ?ｶ趣ｿｽ)
Private Function MirroredContractorText(ByVal wsContractor As Worksheet, ByVal address As String) As String
    If wsContractor Is Nothing Then Exit Function
    MirroredContractorText = CommonNzText(wsContractor.Range(address).MergeArea.Cells(1, 1).value)
End Function

' ?ｿｽ?克ﾒ用 M10:V ?ｿｽs?ｿｽﾌ鯉ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽe?ｿｽ?ｿｽ?ｿｽv?ｿｽ?ｿｽ?ｿｽ[?ｿｽg(M:U)?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽS?ｿｽﾉ拡?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
Private Sub EnsureOfficeChiefRowMerge(ByVal wsTarget As Worksheet, ByVal rowNo As Long)
    Const OFFICE_COL_START As Long = 13  ' M
    Const OFFICE_COL_END As Long = 22    ' V

    Dim mergeRange As Range
    Set mergeRange = wsTarget.Range(wsTarget.Cells(rowNo, OFFICE_COL_START), _
                                    wsTarget.Cells(rowNo, OFFICE_COL_END))
    mod_VendorBlockLayout.SafeUnmergeRange mergeRange
    On Error Resume Next
    mergeRange.Merge
    On Error GoTo 0
    With mergeRange.Cells(1, 1)
        .ShrinkToFit = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

' M10:?ｿｽZ?ｿｽ?ｿｽ / M11:?ｿｽ?ｿｽS?ｿｽH?ｿｽﾆ奇ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ+?ｿｽS?ｿｽp?ｿｽ?ｿｽ+?ｿｽ誧ｲ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ / M12:(?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ)?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ ?ｿｽ?ｿｽ]?ｿｽL
Private Sub ApplyOfficeChiefBlock(ByVal wsInfo As Worksheet, ByVal wsTarget As Worksheet)
    ' M10:U ?ｿｽ?ｿｽ M10:V ?ｿｽﾖ鯉ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽA?ｿｽk?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾄ全?ｿｽﾌ表?ｿｽ?ｿｽ?ｿｽﾉ設抵ｿｽ(M/13?ｿｽ?ｿｽ ?ｿｽ` V/22?ｿｽ?ｿｽ)
    Dim mergeRow As Long
    For mergeRow = 10 To 12
        EnsureOfficeChiefRowMerge wsTarget, mergeRow
    Next mergeRow

    Dim branchName As String, officeName As String
    branchName = CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value)
    officeName = CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value)

    Dim addr As String, title As String, chiefName As String
    Dim coreOffice As String, matchedOffice As String
    If Not mod_FillManagerName.GetOfficeChiefInfo(branchName, officeName, _
             addr, title, chiefName, coreOffice, matchedOffice) Then
        WriteHeaderValue wsTarget.Range("M10"), "", False
        WriteHeaderValue wsTarget.Range("M11"), "", False
        WriteHeaderValue wsTarget.Range("M12"), "", False
        Exit Sub
    End If

    Dim fw As String
    fw = ChrW$(&H3000)

    ' M10: ?ｿｽZ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("M10"), addr, False

    ' M11: ?ｿｽ?ｿｽS?ｿｽH?ｿｽﾆ奇ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ + ?ｿｽS?ｿｽp?ｿｽ?ｿｽ + ?ｿｽ誧ｲ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("M11"), CommonCompanyNameText() & fw & coreOffice, False

    ' M12: ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ=?ｿｽ誧ｲ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾈゑｿｽu?ｿｽ?ｿｽE ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽv?ｿｽA?ｿｽs?ｿｽ?ｿｽv?ｿｽﾈゑｿｽu?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ ?ｿｽ?ｿｽE ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽv
    Dim m12 As String
    If StrComp(CommonNormalizeText(matchedOffice), CommonNormalizeText(coreOffice), vbTextCompare) = 0 Then
        m12 = title & fw & chiefName
    Else
        m12 = matchedOffice & fw & title & fw & chiefName
    End If
    WriteHeaderValueRight wsTarget.Range("M12"), m12
End Sub

' ?ｿｽE?ｿｽl?ｿｽﾌ値?ｿｽ]?ｿｽL(BizUD?ｿｽS?ｿｽV?ｿｽb?ｿｽN?ｿｽK?ｿｽp)
Private Sub WriteHeaderValueRight(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
    End If
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' ?ｿｽ?ｿｽ?ｿｽl?ｿｽﾌ値?ｿｽ]?ｿｽL(BizUD?ｿｽS?ｿｽV?ｿｽb?ｿｽN?ｿｽK?ｿｽp)
Private Sub WriteHeaderValueLeft(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
    End If
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.HorizontalAlignment = xlLeft
    writeCell.VerticalAlignment = xlCenter
End Sub

' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽt?ｿｽﾌ転?ｿｽL(yyyy?ｿｽNm?ｿｽ?ｿｽd?ｿｽ?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ)
Private Sub WriteHeaderValueBottom(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
    End If

    target.MergeArea.Font.Name = BASIC_INFO_REF_FONT_NAME
    target.MergeArea.HorizontalAlignment = xlCenter
    target.MergeArea.VerticalAlignment = xlBottom
End Sub

Private Sub WriteHeaderDateGregorian(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
        If IsDate(value) Then
            writeCell.NumberFormatLocal = "yyyy" & ChrW$(&H5E74) & "m" & ChrW$(&H6708) & "d" & ChrW$(&H65E5)
        End If
    End If
    ApplyHeaderCellFormat target, True
End Sub

' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾖの転?ｿｽL?ｿｽB?ｿｽ?ｿｽ?ｿｽﾊ包ｿｽ(S1:U1/Q2:V2/?ｿｽs20-34)?ｿｽ?ｿｽ ApplyContractorStyleCommonFields?ｿｽB
'   G8:K9 ?ｿｽﾆ者コ?ｿｽ[?ｿｽh(?ｿｽ?克ﾒ用E9?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ) / B12:K12 ?ｿｽZ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ14?ｿｽs?ｿｽﾚ・?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ) /
'   B14:K14 ?ｿｽ?ｿｽﾐ厄ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ11?ｿｽs?ｿｽﾚ・?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ) / C15:I16 ?ｿｽ?ｿｽ\?ｿｽﾒ厄ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ12?ｿｽs?ｿｽﾚ・?ｿｽE?ｿｽl?ｿｽ?ｿｽ) /
'   M9:V9 ?ｿｽZ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M10?ｿｽE?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ) / M10:V10 ?ｿｽ?ｿｽS?ｿｽH?ｿｽﾆ奇ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ+?ｿｽ誧ｲ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M11?ｿｽE?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ) /
'   M11:V11 ?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M12)+?ｿｽu?ｿｽ@?ｿｽ@?ｿｽa?ｿｽv(?ｿｽE?ｿｽE?ｿｽl?ｿｽ?ｿｽ)
Private Sub ApplyAcceptanceHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    Dim wsContractor As Worksheet
    Set wsContractor = ResolveContractorSheetFromTarget(wsTarget)

    ' G8:K9: ?ｿｽﾆ者コ?ｿｽ[?ｿｽh(?ｿｽ?克ﾒ用E9), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("G8"), MirroredContractorText(wsContractor, "E9"), True

    ' B12:K12: ?ｿｽZ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ ?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN14?ｿｽs?ｿｽ?ｿｽ), ?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ
    WriteHeaderValueLeft wsTarget.Range("B12"), wsInfo.Cells(CONTRACTOR_ADDRESS_ROW, valueColumn).value

    ' B14:K14: ?ｿｽ?ｿｽﾐ厄ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ ?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN11?ｿｽs?ｿｽ?ｿｽ), ?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ
    WriteHeaderValueLeft wsTarget.Range("B14"), wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value

    ' C15:I16: ?ｿｽ?ｿｽ\?ｿｽﾒ厄ｿｽ(?ｿｽ?ｿｽ{?ｿｽ?ｿｽ?ｿｽ ?ｿｽ{?ｿｽH?ｿｽ?ｿｽﾐブ?ｿｽ?ｿｽ?ｿｽb?ｿｽN12?ｿｽs?ｿｽ?ｿｽ), ?ｿｽE?ｿｽl?ｿｽ?ｿｽ
    WriteHeaderValueRight wsTarget.Range("C15"), wsInfo.Cells(CONTRACTOR_REPRESENTATIVE_ROW, valueColumn).value

    ' M9:V9: ?ｿｽZ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M10), ?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M10")

    ' M10:V10: ?ｿｽ?ｿｽS?ｿｽH?ｿｽﾆ奇ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ+?ｿｽ誧ｲ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M11), ?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ
    WriteHeaderValueLeft wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M11")

    ' M11:V11: ?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M12) + ?ｿｽu?ｿｽ@?ｿｽ@?ｿｽa?ｿｽv, ?ｿｽE?ｿｽl?ｿｽ?ｿｽ
    Dim chiefText As String
    chiefText = MirroredContractorText(wsContractor, "M12")
    If Len(chiefText) > 0 Then chiefText = chiefText & ContractorHonorificSuffixText()
    WriteHeaderValueRight wsTarget.Range("M11"), chiefText

    mod_OrderTpl_Shared.OrderTplLog "ApplyAcceptanceHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyAcceptanceHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' ?ｿｽx?ｿｽX?ｿｽT?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾖの転?ｿｽL?ｿｽB?ｿｽ?ｿｽ?ｿｽﾊ包ｿｽ(S1:U1/Q2:V2/?ｿｽs20-34)?ｿｽ?ｿｽ ApplyContractorStyleCommonFields?ｿｽB
'   E9:I10 ?ｿｽﾆ者コ?ｿｽ[?ｿｽh(?ｿｽ?克ﾒ用E9?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ) / A13:I15 ?ｿｽ?ｿｽﾐ厄ｿｽ(?ｿｽ?克ﾒ用A13?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ) /
'   M8:V8 ?ｿｽZ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M10?ｿｽE?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ) / M9:V9 ?ｿｽ?ｿｽS?ｿｽH?ｿｽﾆ奇ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ+?ｿｽ誧ｲ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M11?ｿｽE?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ) /
'   M10:V10 ?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M12?ｿｽE?ｿｽE?ｿｽl?ｿｽ?ｿｽ)
Private Sub ApplyBranchCopyHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    Dim wsContractor As Worksheet
    Set wsContractor = ResolveContractorSheetFromTarget(wsTarget)

    ' E9:I10: ?ｿｽﾆ者コ?ｿｽ[?ｿｽh(?ｿｽ?克ﾒ用E9), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("E9"), MirroredContractorText(wsContractor, "E9"), True

    ' A13:I15: ?ｿｽ?ｿｽﾐ厄ｿｽ(?ｿｽ?克ﾒ用A13), ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
    WriteHeaderValue wsTarget.Range("A13"), MirroredContractorText(wsContractor, "A13"), True

    ' M8:V8: ?ｿｽZ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M10), ?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ
    WriteHeaderValueLeft wsTarget.Range("M8"), MirroredContractorText(wsContractor, "M10")

    ' M9:V9: ?ｿｽ?ｿｽS?ｿｽH?ｿｽﾆ奇ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ+?ｿｽ誧ｲ?ｿｽo?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M11), ?ｿｽ?ｿｽ?ｿｽl?ｿｽ?ｿｽ
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M11")

    ' M10:V10: ?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?克ﾒ用M12), ?ｿｽE?ｿｽl?ｿｽ?ｿｽ
    WriteHeaderValueRight wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M12")

    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' ?ｿｽﾊ趣ｿｽ?ｿｽV?ｿｽV?ｿｽ[?ｿｽg?ｿｽﾖの転?ｿｽL(?ｿｽ]?ｿｽL?ｿｽd?ｿｽl?ｿｽ?ｿｽ?ｿｽm?ｿｽ閧ｵ?ｿｽ?ｿｽ?ｿｽ轤ｱ?ｿｽ?ｿｽ?ｿｽﾖ趣ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ)
Private Sub ApplyAttachment3Header(ByVal wsInfo As Worksheet, _
                                   ByVal wsTarget As Worksheet, _
                                   ByVal vendorIndex As Long)
    ' ?ｿｽ]?ｿｽL?ｿｽd?ｿｽl ?ｿｽ?ｿｽ?ｿｽw?ｿｽ?ｿｽ(?ｿｽR?ｿｽs?ｿｽ[?ｿｽﾌゑｿｽ)
End Sub

' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ?ｿｽﾎ会ｿｽ?ｿｽﾌ値?ｿｽ]?ｿｽL(?ｿｽt?ｿｽH?ｿｽ?ｿｽ?ｿｽg?ｿｽK?ｿｽp?ｿｽA?ｿｽK?ｿｽv?ｿｽﾉ会ｿｽ?ｿｽ?ｿｽ?ｿｽﾄ上下?ｿｽ?ｿｽ?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ)
Private Sub WriteHeaderValue(ByVal target As Range, ByVal value As Variant, ByVal centered As Boolean)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
    End If

    ApplyHeaderCellFormat target, centered
End Sub

' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾆゑｿｽ?ｿｽﾄの転?ｿｽL(?ｿｽ?ｿｽ?ｿｽX?ｿｽR?ｿｽ[?ｿｽh?ｿｽ?ｿｽ?ｿｽA?ｿｽ?ｿｽ?ｿｽt?ｿｽ?ｿｽﾏ奇ｿｽ?ｿｽ?ｿｽh?ｿｽ?ｿｽ)
Private Sub WriteHeaderText(ByVal target As Range, ByVal textValue As String, ByVal centered As Boolean)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    writeCell.NumberFormat = "@"
    If Len(Trim$(textValue)) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = textValue
    End If

    ApplyHeaderCellFormat target, centered
End Sub

' ?ｿｽ?ｿｽ?ｿｽt?ｿｽﾌ転?ｿｽL(?ｿｽa?ｿｽ?ｿｽ\?ｿｽ?ｿｽ?ｿｽ`?ｿｽ?ｿｽ?ｿｽA?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽZ?ｿｽ?ｿｽ?ｿｽﾌゑｿｽ?ｿｽﾟ抵ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ)
Private Sub WriteHeaderDate(ByVal target As Range, ByVal value As Variant)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    If IsError(value) Then
        writeCell.ClearContents
    ElseIf Len(Trim$(CStr(value))) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = value
        If IsDate(value) Then
            writeCell.NumberFormat = mod_OrderTpl_Shared.OrderTplEraDateNumberFormatText()
        End If
    End If

    ApplyHeaderCellFormat target, True
End Sub

Private Sub ApplyHeaderCellFormat(ByVal target As Range, ByVal centered As Boolean)
    target.MergeArea.Font.Name = BASIC_INFO_REF_FONT_NAME
    If centered Then
        target.MergeArea.HorizontalAlignment = xlCenter
        target.MergeArea.VerticalAlignment = xlCenter
    End If
End Sub

' ?ｿｽ?克ﾒ用?ｿｽV?ｿｽ[?ｿｽgwsTarget?ｿｽﾆ難ｿｽ?ｿｽ?ｿｽO?ｿｽ?ｿｽ?ｿｽ[?ｿｽv?ｿｽﾌ難ｿｽ?ｿｽ?ｾ細シ?ｿｽ[?ｿｽg?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
Private Function ResolveBreakdownSheetNameFromTarget(ByVal wsTarget As Worksheet) As String
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameBreakdownText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then ResolveBreakdownSheetNameFromTarget = nm
End Function

' ?ｿｽ?ｿｽ?ｿｽ?ｾ細シ?ｿｽ[?ｿｽg?ｿｽ?ｿｽA?ｿｽ?ｿｽﾅ??ｿｽ?ｿｽx?ｿｽ?ｿｽ?ｿｽs?ｿｽ?ｿｽT?ｿｽ?ｿｽ?ｿｽA?ｿｽ?ｿｽ?ｿｽ?ｿｽQ?ｿｽ?ｿｽ?ｿｽﾔゑｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽ?ｿｽ?ｿｽﾂゑｿｽ?ｿｽ?ｿｽﾈゑｿｽ?ｿｽ?ｿｽﾎ具ｿｽ)
Private Function BuildBreakdownQFormula(ByVal sheetName As String, ByVal labelText As String) As String
    Dim q As String
    q = "'" & Replace$(sheetName, "'", "''") & "'"
    BuildBreakdownQFormula = "=IFERROR(INDEX(" & q & "!Q:Q,MATCH(""" & labelText & """," & q & "!A:A,0)),"""")"
End Function

' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ(?ｿｽE?ｿｽl)?ｿｽﾌ転?ｿｽL(BizUD?ｿｽS?ｿｽV?ｿｽb?ｿｽN?ｿｽE?ｿｽ?ｿｽ?ｿｽ?ｿｽﾘゑｿｽ)
Private Sub WriteHeaderFormulaRight(ByVal target As Range, ByVal formulaText As String)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    writeCell.Formula = formulaText
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.NumberFormat = "#,##0;-#,##0;"
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' "?ｿｽ?ｿｽ?ｿｽv"
Private Function ContractorGrandTotalLabelText() As String
    ContractorGrandTotalLabelText = ChrW$(&H5408) & ChrW$(&H8A08)
End Function

' "?ｿｽv"
Private Function ContractorNetTotalLabelText() As String
    ContractorNetTotalLabelText = ChrW$(&H8A08)
End Function

' "?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ"
Private Function ContractorTaxLabelText() As String
    ContractorTaxLabelText = ChrW$(&H6D88) & ChrW$(&H8CBB) & ChrW$(&H7A0E)
End Function

' "?ｿｽ@?ｿｽ@?ｿｽa"(?ｿｽS?ｿｽp?ｿｽ?ｿｽ2?ｿｽ?ｿｽ + ?ｿｽa?ｿｽB?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ M11:V11 ?ｿｽﾌ厄ｿｽ?ｿｽ?ｿｽ?ｿｽﾖ付?ｿｽﾑゑｿｽ?ｿｽ?ｿｽ)
Private Function ContractorHonorificSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H3000) & ChrW$(&H3000) & ChrW$(&H6BBF)
    End If
    ContractorHonorificSuffixText = cached
End Function
