Option Explicit

' ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩe?øΩ?øΩ?øΩv?øΩ?øΩ?øΩ[?øΩg?øΩe?øΩV?øΩ[?øΩg?øΩ÷ÇÃäÔøΩ{?øΩ?øΩ?øΩw?øΩb?øΩ_?øΩ[?øΩ]?øΩL?øΩB
' ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ…âÔøΩ?øΩ?øΩ?øΩA?øΩ?øΩ{?øΩ?øΩ?øΩV?øΩ[?øΩg?øΩÃì]?øΩL?øΩ?øΩ?øΩZ?øΩ?øΩ?øΩœçX?øΩ?øΩ?øΩ…ÇÔøΩ?øΩƒì]?øΩL?øΩ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩC?øΩu?øΩ?øΩ?øΩf)?øΩB
' ?øΩ]?øΩL?øΩŒè€ÉV?øΩ[?øΩg?øΩÃí«âÔøΩ?øΩ?øΩ ApplyVendorSheetHeaders ?øΩÃÉf?øΩB?øΩX?øΩp?øΩb?øΩ`?øΩ÷éÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩﬁÅB
' ?øΩ?øΩ?øΩC?øΩ?øΩ?øΩ?øΩ: CHANGELOG.md ?øΩQ?øΩ?øΩ

' ?øΩ?øΩ{?øΩ?øΩ?øΩÃÉw?øΩb?øΩ_?øΩ[?øΩ]?øΩL?øΩ?øΩ?øΩZ?øΩ?øΩ(?øΩS?øΩ–ãÔøΩ?øΩ ïÔøΩ)?øΩB?øΩu?øΩ?øΩ?øΩb?øΩN?øΩ?øΩ(16/27?øΩs?øΩ?øΩ)?øΩÕìÔøΩ?øΩI?øΩ…ëg?øΩ›óÔøΩ?øΩƒÇÔøΩ
Private Const HEADER_SOURCE_COMMON_CELLS As String = "B6,C6,C2,C9,C10,C13,C15:C16,F6"
Private Const HEADER_DATE_FONT_SIZE As Double = 14#
' ?øΩ?çé“óp?øΩV?øΩ[?øΩg?øΩ]?øΩL?øΩ≈éQ?øΩ∆ÇÔøΩ?øΩ?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN?øΩs(?øΩ?øΩ{?øΩ?øΩ?øΩ)
Private Const CONTRACTOR_CONTRACT_AMOUNT_ROW As Long = 33
Private Const CONTRACTOR_CONSUMPTION_TAX_ROW As Long = 34
Private Const CONTRACTOR_CONTRACT_TOTAL_ROW As Long = 35
' ?øΩ?øΩ{?øΩ?øΩ?øΩ ?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN: ?øΩ?øΩ\?øΩ“ñÔøΩ(12?øΩs?øΩ?øΩ)?øΩE?øΩZ?øΩ?øΩ(14?øΩs?øΩ?øΩ)
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

' ?øΩw?øΩ?øΩu?øΩ?øΩ?øΩb?øΩN?øΩÃé{?øΩH?øΩ?øΩ–Ç…ëŒâÔøΩ?øΩ?øΩ?øΩ?øΩe?øΩ?øΩ?øΩv?øΩ?øΩ?øΩ[?øΩg5?øΩV?øΩ[?øΩg?øΩ÷Éw?øΩb?øΩ_?øΩ[?øΩ?øΩ]?øΩL?øΩ?øΩ?øΩ?øΩ(?øΩf?øΩB?øΩX?øΩp?øΩb?øΩ`?øΩ?øΩ)
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

' ?øΩS?øΩm?øΩ?øΩ?øΩ–ÇÃÉe?øΩ?øΩ?øΩv?øΩ?øΩ?øΩ[?øΩg?øΩV?øΩ[?øΩg?øΩ÷Éw?øΩb?øΩ_?øΩ[?øΩ?øΩ?øΩƒì]?øΩL?øΩ?øΩ?øΩ?øΩ
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

' Sheet1(?øΩ?øΩ{?øΩ?øΩ?øΩ)?øΩ?øΩWorksheet_Change?øΩ?øΩ?øΩ?øΩƒÇŒÇÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩB
' ?øΩw?øΩb?øΩ_?øΩ[?øΩ]?øΩL?øΩ?øΩ?øΩZ?øΩ?øΩ(B6/C6/C2/C9/C10/C15:C16/F6?øΩA?øΩe?øΩu?øΩ?øΩ?øΩb?øΩN?øΩ?øΩ16/27?øΩs?øΩ?øΩ)?øΩÃïœçX?øΩ?øΩ?øΩe?øΩ–ÉV?øΩ[?øΩg?øΩ÷îÔøΩ?øΩf?øΩ?øΩ?øΩ?øΩ
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

' ?øΩw?øΩb?øΩ_?øΩ[?øΩ]?øΩL?øΩ?øΩ?øΩZ?øΩ?øΩ?øΩÃäƒéÔøΩ?øΩÕàÕÇÔøΩ‘ÇÔøΩ?øΩ?øΩ?øΩJ?øΩ?øΩ?øΩb?øΩp?øΩ[(Sheet1?øΩÃïœçX?øΩQ?øΩ[?øΩg?øΩ\?øΩz?øΩp)
Public Function GetBasicInfoHeaderSourceMonitorRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function
    Set GetBasicInfoHeaderSourceMonitorRange = BuildHeaderSourceRange(wsInfo)
End Function

' ?øΩw?øΩb?øΩ_?øΩ[?øΩ]?øΩL?øΩ?øΩ?øΩZ?øΩ?øΩ?øΩÃäƒéÔøΩ?øΩÕàÔøΩ(?øΩ?øΩ?øΩ ÉZ?øΩ?øΩ + ?øΩe?øΩu?øΩ?øΩ?øΩb?øΩN?øΩÃã∆é“ÉR?øΩ[?øΩh16?øΩs?øΩ?øΩ/?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ27?øΩs?øΩ?øΩ)
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

' ?øΩV?øΩ[?øΩg?øΩ?øΩ?øΩo?øΩ?øΩ(?øΩ^?øΩu)?øΩF: ?øΩH?øΩ?øΩ?øΩÊï™(?øΩ?øΩ{?øΩ?øΩ?øΩ10?øΩs?øΩ?øΩ)?øΩZ?øΩ?øΩ?øΩÃìh?øΩ?øΩ¬Ç‘ÇÔøΩ?øΩF?øΩ?øΩK?øΩp?øΩ?øΩ?øΩ?øΩ
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


' ?øΩ?øΩ?øΩ?æç◊Éw?øΩb?øΩ_?øΩ[?øΩ?øΩ?øΩ÷äÔøΩ{?øΩ?øΩ?øΩV?øΩ[?øΩg?øΩÃìÔøΩ?øΩe?øΩ?øΩ]?øΩL?øΩ?øΩ?øΩ?øΩ
Public Sub ApplyBreakdownHeader(ByVal wsInfo As Worksheet, _
                                ByVal wsBreakdown As Worksheet, _
                                ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsBreakdown Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' C2: ?øΩ?øΩ?øΩX?øΩR?øΩ[?øΩh(?øΩo?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ_?øΩP?øΩ?øΩ?øΩK?øΩp?øΩ?øΩ?øΩ?øΩÃíP?øΩ?øΩ?øΩK?øΩp?øΩ?øΩ?øΩ?øΩV?øΩ[?øΩg B/C?øΩ?øΩ∆çÔøΩ ?øΩ?øΩ G?øΩ?øΩ)
    Dim branchOfficeCode As String
    branchOfficeCode = mod_OrderTpl_Shared.OrderTplResolveBranchOfficeCode( _
        CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value), _
        CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    WriteHeaderText wsBreakdown.Range("C2"), branchOfficeCode, False

    ' C3: ?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ(?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN27?øΩs?øΩ?øΩ)
    WriteHeaderValue wsBreakdown.Range("C3"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, False

    ' B6: ?øΩH?øΩ?øΩ?øΩ‘çÔøΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC9)
    WriteHeaderValue wsBreakdown.Range("B6"), wsInfo.Range("C9").value, False

    ' D6: ?øΩH?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC10?øΩA?øΩ?øΩ?øΩ?øΩ?øΩZ?øΩ?øΩ?øΩÃÇÔøΩ?øΩﬂíÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)
    WriteHeaderValue wsBreakdown.Range("D6"), wsInfo.Range("C10").value, True

    ' L5/L6: ?øΩH?øΩ?øΩ ?øΩ?øΩ/?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC15/C16?øΩA?øΩa?øΩ?øΩ\?øΩ?øΩ)
    WriteHeaderDate wsBreakdown.Range("L5"), wsInfo.Range("C15").value
    WriteHeaderDate wsBreakdown.Range("L6"), wsInfo.Range("C16").value
    wsBreakdown.Range("L5:L6").Font.Size = HEADER_DATE_FONT_SIZE

    ' O2: ?øΩ?ê¨?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC2?øΩA?øΩa?øΩ?øΩ\?øΩ?øΩ)
    ' P2:Q2 created date from C2. O2 keeps template label.
    WriteHeaderDate wsBreakdown.Range("P2"), wsInfo.Range("C2").value
    wsBreakdown.Range("P2").MergeArea.Font.Size = HEADER_DATE_FONT_SIZE

    ' O3: ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩF6)
    ' P3:Q3 branch chief from F6. O3 keeps template label.
    WriteHeaderValue wsBreakdown.Range("P3"), wsInfo.Range("F6").value, True

    ' P5: ?øΩO?øΩ?øΩ?øΩ?øΩ–ñÔøΩ(?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN11?øΩs?øΩ?øΩ)
    WriteHeaderValue wsBreakdown.Range("P5"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' P6: ?øΩ∆é“ÉR?øΩ[?øΩh(?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN16?øΩs?øΩ?øΩ)
    WriteHeaderValue wsBreakdown.Range("P6"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' ?øΩ?øΩ?øΩ?øΩ?øΩy?øΩ[?øΩW?øΩ?øΩ?øΩ?øΩp?øΩ…Éw?øΩb?øΩ_?øΩ[?øΩs(7:10?øΩs?øΩ?øΩ)?øΩ?øΩ?øΩ^?øΩC?øΩg?øΩ?øΩ?øΩs?øΩ…ê›íËÇ∑?øΩ?øΩ
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

' ?øΩ?çé“óp/?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ/?øΩx?øΩX?øΩT ?øΩ?øΩ?øΩ?øΩ: S1:U1 ?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ(27) / Q2:V2 ?øΩ?ê¨?øΩ?øΩC2(?øΩ?øΩ?øΩ?øΩ) /
'   ?øΩs20-34(E20:?øΩH?øΩ?øΩ?øΩ?øΩC10 E22:?øΩs?øΩ?øΩ?øΩ{?øΩ?øΩC13 G24:?øΩH?øΩ?øΩ?øΩ?øΩC15(?øΩ?øΩ?øΩ?øΩ) G26:?øΩH?øΩ?øΩ?øΩ?øΩC16(?øΩ?øΩ?øΩ?øΩ)?øΩA
'   Q22:?øΩ≈çÔøΩ(35) Q23:?øΩ≈îÔøΩ(33) Q24:?øΩ?øΩ?øΩ?øΩ?øΩ(34)?øΩAC30/H30/J30/M34/R34/F32/F33 ?øΩ?øΩ Reapply?øΩn?øΩo?øΩR)?øΩ?øΩ
'   ApplyContractorStyleCommonFields ?øΩ≈ÅA3?øΩV?øΩ[?øΩg?øΩ∆ÇÔøΩ?øΩ?øΩ?øΩÍÉçÔøΩW?øΩb?øΩN?øΩ?øΩK?øΩp?øΩ?øΩ?øΩ?øΩB
' ?øΩ?çé“óp?øΩV?øΩ[?øΩg?øΩ≈óL: E9:?øΩ∆é“ÉR?øΩ[?øΩh(16) A13:?øΩ?øΩ–ñÔøΩ(11) M10:?øΩ?øΩ?øΩ?øΩ?øΩ“èZ?øΩ?øΩ
'   M11:?øΩ?øΩS?øΩH?øΩ∆äÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ+?øΩÓä≤?øΩo?øΩ?øΩ?øΩ?øΩ M12:?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ(?øΩo?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩX?øΩg?øΩQ?øΩ?øΩ)
Private Sub ApplyContractorHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ApplyContractorStyleCommonFields wsInfo, wsTarget, valueColumn

    ' E9: ?øΩ∆é“ÉR?øΩ[?øΩh(?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN16?øΩs), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("E9"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' A13: ?øΩ?øΩ–ñÔøΩ(?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN11?øΩs), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("A13"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' M10/M11/M12: ?øΩo?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩX?øΩg?øΩQ?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ ?øΩZ?øΩ?øΩ/?øΩ?øΩ?øΩ?øΩ/?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ)
    ApplyOfficeChiefBlock wsInfo, wsTarget

    ' ?øΩs38-42(?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ/?øΩJ?øΩ?øΩ/?øΩx?øΩ?øΩ?øΩﬁóÔøΩ/?øΩ›ó^?øΩi/?øΩ?øΩ?øΩT?øΩC?øΩN?øΩ?øΩ)?øΩ?øΩ?øΩ?øΩ{?øΩ?øΩ?©ÇÔøΩ
    ' ?øΩ?çé“óp?øΩE?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩE?øΩx?øΩX?øΩT?øΩ÷çƒì]?øΩL(Reapply?øΩ?øΩ?øΩ≈ëS?øΩƒÇÃêÔøΩ?øΩ?øΩ?øΩœÉV?øΩ[?øΩg?øΩ÷ìK?øΩp)
    mod_BasicInfoExclusiveChoice.ReapplyExclusiveChoices wsInfo, vendorIndex
    mod_BasicInfoSupplyLoan.ReapplySupplyLoan wsInfo, vendorIndex

    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' ?øΩ?çé“óp/?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ/?øΩx?øΩX?øΩT?øΩ≈ãÔøΩ?øΩ ÇÃì]?øΩL(S1:U1 ?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ / Q2:V2 ?øΩ?ê¨?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ) /
' ?øΩs20-34: E20 ?øΩH?øΩ?øΩ?øΩ?øΩ?øΩEE22 ?øΩs?øΩ?øΩ?øΩ{?øΩ?øΩ?øΩEG24/G26 ?øΩH?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ)?øΩEQ22-24 ?øΩ?øΩ?øΩ?æç◊éQ?øΩ?øΩ)?øΩB
' ?øΩA3?øΩV?øΩ[?øΩg?øΩ∆ÇÔøΩ?øΩZ?øΩ?øΩ?øΩ\?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩS?øΩ?øΩv?øΩ?øΩ?øΩƒÇÔøΩ?øΩÈÇΩ?øΩﬂÅA?øΩ?øΩ?øΩÍÉçÔøΩW?øΩb?øΩN?øΩ?øΩ?øΩ?øΩ?øΩÍÇº?øΩ?øΩK?øΩp?øΩ?øΩ?øΩ?øΩB
Private Sub ApplyContractorStyleCommonFields(ByVal wsInfo As Worksheet, _
                                             ByVal wsTarget As Worksheet, _
                                             ByVal valueColumn As Long)
    ' S1:U1: ?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ(?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN27?øΩs), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("S1"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, True

    ' Q2:V2: ?øΩ?ê¨?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC2, ?øΩ?øΩ?øΩ?øΩ), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderDateGregorian wsTarget.Range("Q2"), wsInfo.Range("C2").value

    ' E20: ?øΩH?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC10), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("E20"), wsInfo.Range("C10").value, True

    ' E22: ?øΩs?øΩ?øΩ?øΩ{?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC13), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("E22"), wsInfo.Range("C13").value, True

    ' G24: ?øΩH?øΩ?øΩ ?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC15, ?øΩ?øΩ?øΩ?øΩ), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderDateGregorian wsTarget.Range("G24"), wsInfo.Range("C15").value

    ' G26: ?øΩH?øΩ?øΩ ?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩC16, ?øΩ?øΩ?øΩ?øΩ), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderDateGregorian wsTarget.Range("G26"), wsInfo.Range("C16").value

    ' Q22/Q23/Q24: ?øΩ?øΩ?øΩ?øΩO?øΩ?øΩ?øΩ[?øΩv?øΩÃìÔøΩ?øΩ?æç◊ÉV?øΩ[?øΩg ?øΩ?øΩ?øΩv/?øΩv/?øΩ?øΩ?øΩ?øΩ?øΩ ?øΩs(Q?øΩ?øΩ)?øΩ?øΩ?øΩQ?øΩ∆ÇÔøΩ?øΩÈê∂?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
    ' (?øΩ?øΩ?øΩ?æç◊ÇÃíl?øΩ?øΩ?øΩX?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ∆éÔøΩ?øΩ?øΩ?øΩ≈çX?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ)
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

' wsTarget(?øΩ?øΩ?øΩ?øΩ?øΩœÉe?øΩ?øΩ?øΩv?øΩ?øΩ?øΩ[?øΩg?øΩV?øΩ[?øΩg)?øΩ∆ìÔøΩ?øΩ?øΩO?øΩ?øΩ?øΩ[?øΩv?øΩÃéÛíçé“óp(?øΩ?øΩ?øΩ?øΩ)?øΩV?øΩ[?øΩg?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
Private Function ResolveContractorSheetFromTarget(ByVal wsTarget As Worksheet) As Worksheet
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameContractorText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then Set ResolveContractorSheetFromTarget = ThisWorkbook.Worksheets(nm)
End Function

' ?øΩ?çé“óp?øΩV?øΩ[?øΩg?øΩÃåÔøΩ?øΩ?øΩ?øΩZ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩl?øΩ?øΩ?øΩÊìæ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩ›ÇÔøΩ?øΩ»ÇÔøΩ/?øΩ?øΩ»ÇÔøΩ?∂éÔøΩ)
Private Function MirroredContractorText(ByVal wsContractor As Worksheet, ByVal address As String) As String
    If wsContractor Is Nothing Then Exit Function
    MirroredContractorText = CommonNzText(wsContractor.Range(address).MergeArea.Cells(1, 1).value)
End Function

' ?øΩ?çé“óp M10:V ?øΩs?øΩÃåÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩe?øΩ?øΩ?øΩv?øΩ?øΩ?øΩ[?øΩg(M:U)?øΩ?øΩ?øΩ?øΩ?øΩ?øΩS?øΩ…äg?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
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

' M10:?øΩZ?øΩ?øΩ / M11:?øΩ?øΩS?øΩH?øΩ∆äÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ+?øΩS?øΩp?øΩ?øΩ+?øΩÓä≤?øΩo?øΩ?øΩ?øΩ?øΩ / M12:(?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ ?øΩ?øΩ]?øΩL
Private Sub ApplyOfficeChiefBlock(ByVal wsInfo As Worksheet, ByVal wsTarget As Worksheet)
    ' M10:U ?øΩ?øΩ M10:V ?øΩ÷åÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩA?øΩk?øΩ?øΩ?øΩ?øΩ?øΩƒëS?øΩÃï\?øΩ?øΩ?øΩ…ê›íÔøΩ(M/13?øΩ?øΩ ?øΩ` V/22?øΩ?øΩ)
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

    ' M10: ?øΩZ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("M10"), addr, False

    ' M11: ?øΩ?øΩS?øΩH?øΩ∆äÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ + ?øΩS?øΩp?øΩ?øΩ + ?øΩÓä≤?øΩo?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("M11"), CommonCompanyNameText() & fw & coreOffice, False

    ' M12: ?øΩo?øΩ?øΩ?øΩ?øΩ=?øΩÓä≤?øΩo?øΩ?øΩ?øΩ?øΩ?øΩ»ÇÔøΩu?øΩ?øΩE ?øΩ?øΩ?øΩ?øΩ?øΩv?øΩA?øΩs?øΩ?øΩv?øΩ»ÇÔøΩu?øΩo?øΩ?øΩ?øΩ?øΩ ?øΩ?øΩE ?øΩ?øΩ?øΩ?øΩ?øΩv
    Dim m12 As String
    If StrComp(CommonNormalizeText(matchedOffice), CommonNormalizeText(coreOffice), vbTextCompare) = 0 Then
        m12 = title & fw & chiefName
    Else
        m12 = matchedOffice & fw & title & fw & chiefName
    End If
    WriteHeaderValueRight wsTarget.Range("M12"), m12
End Sub

' ?øΩE?øΩl?øΩÃíl?øΩ]?øΩL(BizUD?øΩS?øΩV?øΩb?øΩN?øΩK?øΩp)
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

' ?øΩ?øΩ?øΩl?øΩÃíl?øΩ]?øΩL(BizUD?øΩS?øΩV?øΩb?øΩN?øΩK?øΩp)
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

' ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩt?øΩÃì]?øΩL(yyyy?øΩNm?øΩ?øΩd?øΩ?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ)
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

' ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩV?øΩ[?øΩg?øΩ÷ÇÃì]?øΩL?øΩB?øΩ?øΩ?øΩ ïÔøΩ(S1:U1/Q2:V2/?øΩs20-34)?øΩ?øΩ ApplyContractorStyleCommonFields?øΩB
'   G8:K9 ?øΩ∆é“ÉR?øΩ[?øΩh(?øΩ?çé“ópE9?øΩE?øΩ?øΩ?øΩ?øΩ) / B12:K12 ?øΩZ?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩ14?øΩs?øΩ⁄ÅE?øΩ?øΩ?øΩl?øΩ?øΩ) /
'   B14:K14 ?øΩ?øΩ–ñÔøΩ(?øΩ?øΩ{?øΩ?øΩ?øΩ11?øΩs?øΩ⁄ÅE?øΩ?øΩ?øΩl?øΩ?øΩ) / C15:I16 ?øΩ?øΩ\?øΩ“ñÔøΩ(?øΩ?øΩ{?øΩ?øΩ?øΩ12?øΩs?øΩ⁄ÅE?øΩE?øΩl?øΩ?øΩ) /
'   M9:V9 ?øΩZ?øΩ?øΩ(?øΩ?çé“ópM10?øΩE?øΩ?øΩ?øΩl?øΩ?øΩ) / M10:V10 ?øΩ?øΩS?øΩH?øΩ∆äÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ+?øΩÓä≤?øΩo?øΩ?øΩ?øΩ?øΩ(?øΩ?çé“ópM11?øΩE?øΩ?øΩ?øΩl?øΩ?øΩ) /
'   M11:V11 ?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ(?øΩ?çé“ópM12)+?øΩu?øΩ@?øΩ@?øΩa?øΩv(?øΩE?øΩE?øΩl?øΩ?øΩ)
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

    ' G8:K9: ?øΩ∆é“ÉR?øΩ[?øΩh(?øΩ?çé“ópE9), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("G8"), MirroredContractorText(wsContractor, "E9"), True

    ' B12:K12: ?øΩZ?øΩ?øΩ(?øΩ?øΩ{?øΩ?øΩ?øΩ ?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN14?øΩs?øΩ?øΩ), ?øΩ?øΩ?øΩl?øΩ?øΩ
    WriteHeaderValueLeft wsTarget.Range("B12"), wsInfo.Cells(CONTRACTOR_ADDRESS_ROW, valueColumn).value

    ' B14:K14: ?øΩ?øΩ–ñÔøΩ(?øΩ?øΩ{?øΩ?øΩ?øΩ ?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN11?øΩs?øΩ?øΩ), ?øΩ?øΩ?øΩl?øΩ?øΩ
    WriteHeaderValueLeft wsTarget.Range("B14"), wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value

    ' C15:I16: ?øΩ?øΩ\?øΩ“ñÔøΩ(?øΩ?øΩ{?øΩ?øΩ?øΩ ?øΩ{?øΩH?øΩ?øΩ–Éu?øΩ?øΩ?øΩb?øΩN12?øΩs?øΩ?øΩ), ?øΩE?øΩl?øΩ?øΩ
    WriteHeaderValueRight wsTarget.Range("C15"), wsInfo.Cells(CONTRACTOR_REPRESENTATIVE_ROW, valueColumn).value

    ' M9:V9: ?øΩZ?øΩ?øΩ(?øΩ?çé“ópM10), ?øΩ?øΩ?øΩl?øΩ?øΩ
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M10")

    ' M10:V10: ?øΩ?øΩS?øΩH?øΩ∆äÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ+?øΩÓä≤?øΩo?øΩ?øΩ?øΩ?øΩ(?øΩ?çé“ópM11), ?øΩ?øΩ?øΩl?øΩ?øΩ
    WriteHeaderValueLeft wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M11")

    ' M11:V11: ?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ(?øΩ?çé“ópM12) + ?øΩu?øΩ@?øΩ@?øΩa?øΩv, ?øΩE?øΩl?øΩ?øΩ
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

' ?øΩx?øΩX?øΩT?øΩV?øΩ[?øΩg?øΩ÷ÇÃì]?øΩL?øΩB?øΩ?øΩ?øΩ ïÔøΩ(S1:U1/Q2:V2/?øΩs20-34)?øΩ?øΩ ApplyContractorStyleCommonFields?øΩB
'   E9:I10 ?øΩ∆é“ÉR?øΩ[?øΩh(?øΩ?çé“ópE9?øΩE?øΩ?øΩ?øΩ?øΩ) / A13:I15 ?øΩ?øΩ–ñÔøΩ(?øΩ?çé“ópA13?øΩE?øΩ?øΩ?øΩ?øΩ) /
'   M8:V8 ?øΩZ?øΩ?øΩ(?øΩ?çé“ópM10?øΩE?øΩ?øΩ?øΩl?øΩ?øΩ) / M9:V9 ?øΩ?øΩS?øΩH?øΩ∆äÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ+?øΩÓä≤?øΩo?øΩ?øΩ?øΩ?øΩ(?øΩ?çé“ópM11?øΩE?øΩ?øΩ?øΩl?øΩ?øΩ) /
'   M10:V10 ?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ(?øΩ?çé“ópM12?øΩE?øΩE?øΩl?øΩ?øΩ)
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

    ' E9:I10: ?øΩ∆é“ÉR?øΩ[?øΩh(?øΩ?çé“ópE9), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("E9"), MirroredContractorText(wsContractor, "E9"), True

    ' A13:I15: ?øΩ?øΩ–ñÔøΩ(?øΩ?çé“ópA13), ?øΩ?øΩ?øΩ?øΩ
    WriteHeaderValue wsTarget.Range("A13"), MirroredContractorText(wsContractor, "A13"), True

    ' M8:V8: ?øΩZ?øΩ?øΩ(?øΩ?çé“ópM10), ?øΩ?øΩ?øΩl?øΩ?øΩ
    WriteHeaderValueLeft wsTarget.Range("M8"), MirroredContractorText(wsContractor, "M10")

    ' M9:V9: ?øΩ?øΩS?øΩH?øΩ∆äÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ+?øΩÓä≤?øΩo?øΩ?øΩ?øΩ?øΩ(?øΩ?çé“ópM11), ?øΩ?øΩ?øΩl?øΩ?øΩ
    WriteHeaderValueLeft wsTarget.Range("M9"), MirroredContractorText(wsContractor, "M11")

    ' M10:V10: ?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ(?øΩ?çé“ópM12), ?øΩE?øΩl?øΩ?øΩ
    WriteHeaderValueRight wsTarget.Range("M10"), MirroredContractorText(wsContractor, "M12")

    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBranchCopyHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
End Sub

' ?øΩ éÔøΩ?øΩV?øΩV?øΩ[?øΩg?øΩ÷ÇÃì]?øΩL(?øΩ]?øΩL?øΩd?øΩl?øΩ?øΩ?øΩm?øΩËÇµ?øΩ?øΩ?øΩÁÇ±?øΩ?øΩ?øΩ÷éÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)
Private Sub ApplyAttachment3Header(ByVal wsInfo As Worksheet, _
                                   ByVal wsTarget As Worksheet, _
                                   ByVal vendorIndex As Long)
    ' O1: é{çHâÔé–ñº(äÓñ{èÓïÒ é{çHâÔé–óÒ11çsñ⁄)ÇâEãlÇﬂÅEBIZ UDÉSÉVÉbÉNÇ≈ì]ãL
    Dim vendorCol As Long
    vendorCol = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)
    Dim vendorName As String
    vendorName = GetBasicInfoCellText(wsInfo, wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, vendorCol).Address)
    WriteAttachment3VendorName wsTarget.Range("O1"), vendorName

    ' 36Å`41çsñ⁄ F/HóÒÅEJ/LóÒÇÃîrëºÉ`ÉFÉbÉNÉ{ÉbÉNÉXêßå‰
    SetupAttachment3CheckboxExclusivity wsTarget

    ' E49:I49Å`E52:I52 É_ÉuÉãÉNÉäÉbÉNÇ…ÇÊÇÈéYã∆îpä¸ï®èàóùé{ê›ëIë
    SetupAttachment3SanpaiFacilityDoubleClickHint wsTarget

    ' J54:M54 éYîpçsJRã‡äzçáåv
    RefreshAttachment3SanpaiJrTotal wsTarget
End Sub

' O1(é{çHâÔé–ñº)ÇÃì]ãL(âEãlÇﬂÅEBIZ UDÉSÉVÉbÉN)
Private Sub WriteAttachment3VendorName(ByVal target As Range, ByVal vendorName As String)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)

    writeCell.NumberFormat = "@"
    If Len(Trim$(vendorName)) = 0 Then
        writeCell.ClearContents
    Else
        writeCell.value = vendorName
    End If

    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' ===== ï éÜáV 36Å`41çsñ⁄ F/HóÒÅEJ/LóÒ îrëºÉ`ÉFÉbÉNÉ{ÉbÉNÉX =====
' FóÒ=óL / HóÒ=ñ≥ ÇÃÉyÉAÇ∆ÅAJóÒÅELóÒÇÃÉyÉAÇÇªÇÍÇºÇÍîrëºâªÇ∑ÇÈÅB
' Ç≥ÇÁÇ… HóÒ(ñ≥)Ç™É`ÉFÉbÉNÇ≥ÇÍÇƒÇ¢ÇÈçsÇÕ JóÒÅELóÒÇÃÇ«ÇøÇÁÇ‡É`ÉFÉbÉNÇ≈Ç´Ç»Ç¢ÅB
Private Const ATTACHMENT3_CHECK_ROW_MIN As Long = 36
Private Const ATTACHMENT3_CHECK_ROW_MAX As Long = 41
Private Const ATTACHMENT3_COL_F As Long = 6
Private Const ATTACHMENT3_COL_H As Long = 8
Private Const ATTACHMENT3_COL_J As Long = 10
Private Const ATTACHMENT3_COL_L As Long = 12

Public Sub SetupAttachment3CheckboxExclusivity(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Dim cb As Object
    For Each cb In ws.CheckBoxes
        Dim r As Long, c As Long
        r = cb.TopLeftCell.Row
        c = cb.TopLeftCell.Column
        If r >= ATTACHMENT3_CHECK_ROW_MIN And r <= ATTACHMENT3_CHECK_ROW_MAX Then
            If c = ATTACHMENT3_COL_F Or c = ATTACHMENT3_COL_H Or _
               c = ATTACHMENT3_COL_J Or c = ATTACHMENT3_COL_L Then
                cb.OnAction = "'" & ThisWorkbook.Name & "'!Attachment3CheckboxClick"
            End If
        End If
    Next cb
    On Error GoTo 0

    NormalizeAttachment3CheckboxRows ws
End Sub

' É`ÉFÉbÉNÉ{ÉbÉNÉXÉNÉäÉbÉNéûÇÃîrëºêßå‰ñ{ëÃ(äeÉ`ÉFÉbÉNÉ{ÉbÉNÉXÇÃ OnAction Ç©ÇÁåƒÇŒÇÍÇÈ)
Public Sub Attachment3CheckboxClick()
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

    Dim rowIndex As Long, colIndex As Long
    rowIndex = clicked.TopLeftCell.Row
    colIndex = clicked.TopLeftCell.Column

    Application.EnableEvents = False

    Select Case colIndex
        Case ATTACHMENT3_COL_F
            If clicked.value = xlOn Then
                SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_H, False
            End If
        Case ATTACHMENT3_COL_H
            If clicked.value = xlOn Then
                SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_F, False
                ' H(ñ≥)ÇëIëÇµÇΩçsÇÕ JóÒÅELóÒÇÃÇ«ÇøÇÁÇ‡É`ÉFÉbÉNïsâ¬Ç∆Ç∑ÇÈ
                SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_J, False
                SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_L, False
            End If
        Case ATTACHMENT3_COL_J
            If clicked.value = xlOn Then
                If IsAttachment3CheckboxOn(ws, rowIndex, ATTACHMENT3_COL_H) Then
                    clicked.value = xlOff
                Else
                    SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_L, False
                End If
            End If
        Case ATTACHMENT3_COL_L
            If clicked.value = xlOn Then
                If IsAttachment3CheckboxOn(ws, rowIndex, ATTACHMENT3_COL_H) Then
                    clicked.value = xlOff
                Else
                    SetAttachment3Checkbox ws, rowIndex, ATTACHMENT3_COL_J, False
                End If
            End If
    End Select

    Application.EnableEvents = True
    Exit Sub

Done:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Function FindAttachment3Checkbox(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As Object
    Dim cb As Object
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If cb.TopLeftCell.Row = rowIndex And cb.TopLeftCell.Column = colIndex Then
            Set FindAttachment3Checkbox = cb
            Exit Function
        End If
    Next cb
    On Error GoTo 0
End Function

Private Sub SetAttachment3Checkbox(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long, ByVal turnOn As Boolean)
    Dim cb As Object
    Set cb = FindAttachment3Checkbox(ws, rowIndex, colIndex)
    If cb Is Nothing Then Exit Sub
    If turnOn Then
        cb.value = xlOn
    Else
        cb.value = xlOff
    End If
End Sub

Private Function IsAttachment3CheckboxOn(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As Boolean
    Dim cb As Object
    Set cb = FindAttachment3Checkbox(ws, rowIndex, colIndex)
    If cb Is Nothing Then Exit Function
    IsAttachment3CheckboxOn = (cb.value = xlOn)
End Function

' ä˘ë∂ÇÃëIëì‡óeÇ…ñµèÇ(F/Hóºï˚É`ÉFÉbÉNÅAH+J/LÇÃïπópÅAJ/Lóºï˚É`ÉFÉbÉNìô)Ç™Ç†ÇÈèÍçáÇÃÇ›ï‚ê≥Ç∑ÇÈÅB
' êVãKÇ…ä˘íËílÇïtó^Ç∑ÇÈÇ±Ç∆ÇÕÇµÇ»Ç¢ÅB
Private Sub NormalizeAttachment3CheckboxRows(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next

    Dim r As Long
    For r = ATTACHMENT3_CHECK_ROW_MIN To ATTACHMENT3_CHECK_ROW_MAX
        If IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_F) And _
           IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_H) Then
            SetAttachment3Checkbox ws, r, ATTACHMENT3_COL_H, False
        End If

        If IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_H) Then
            SetAttachment3Checkbox ws, r, ATTACHMENT3_COL_J, False
            SetAttachment3Checkbox ws, r, ATTACHMENT3_COL_L, False
        ElseIf IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_J) And _
               IsAttachment3CheckboxOn(ws, r, ATTACHMENT3_COL_L) Then
            SetAttachment3Checkbox ws, r, ATTACHMENT3_COL_L, False
        End If
    Next r

    On Error GoTo 0
    Application.EnableEvents = prevEvents
End Sub

' ===== ï éÜáV E49:I49Å`E52:I52 éYã∆îpä¸ï®èàóùé{ê›ëIë =====
Public Const ATTACHMENT3_SANPAI_ROW_MIN As Long = 49
Public Const ATTACHMENT3_SANPAI_ROW_MAX As Long = 52
Public Const ATTACHMENT3_SANPAI_NAME_COL As Long = 5    ' E
Public Const ATTACHMENT3_SANPAI_VALUE_COL As Long = 10  ' J
Private Const SANPAI_FACILITY_MASTER_START_ROW As Long = 59

' É_ÉuÉãÉNÉäÉbÉNëŒè€îÕàÕÇÕê∂ê¨éûì_Ç≈èëéÆÅEï€åÏìôÇÃí«â¡ê›íËÇÕïsóvÇÃÇΩÇﬂåªèÛÇÕãÛé¿ëïÅB
' (ëŒè€îªíËÅEëIëèàóùÇÕ ThisWorkbook ë§Ç©ÇÁ mod_OrderTpl_Header ÇÃä÷êîÇíºê⁄åƒÇ—èoÇ∑)
Private Sub SetupAttachment3SanpaiFacilityDoubleClickHint(ByVal ws As Worksheet)
    ' ó\ñÒ(è´óàÇÃèëéÆïtó^ìôÇ…îıÇ¶ÇΩÉtÉbÉN)
End Sub

' Attachment3(ï éÜáV)ÉVÅ[ÉgÇ©Ç«Ç§Ç©ÇîªíËÇ∑ÇÈ
Public Function IsAttachment3Sheet(ByVal ws As Worksheet) As Boolean
    IsAttachment3Sheet = False
    If ws Is Nothing Then Exit Function
    Dim baseName As String, aliasText As String
    If mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(ws, baseName, aliasText) Then
        IsAttachment3Sheet = (StrComp(baseName, mod_OrderTpl_Shared.OrderTplBaseNameAttachment3Text(), vbTextCompare) = 0)
    End If
End Function

' E49:I49Å`E52:I52 (éYã∆îpä¸ï®èàóùé{ê›ëIëÉZÉã)Ç©Ç«Ç§Ç©ÇîªíËÇ∑ÇÈ
Public Function IsAttachment3SanpaiFacilityTarget(ByVal ws As Worksheet, ByVal target As Range) As Boolean
    IsAttachment3SanpaiFacilityTarget = False
    If ws Is Nothing Or target Is Nothing Then Exit Function
    If Not IsAttachment3Sheet(ws) Then Exit Function

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.MergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    Dim topLeft As Range
    Set topLeft = targetArea.Cells(1, 1)

    If topLeft.Row < ATTACHMENT3_SANPAI_ROW_MIN Or topLeft.Row > ATTACHMENT3_SANPAI_ROW_MAX Then Exit Function
    IsAttachment3SanpaiFacilityTarget = (topLeft.Column = ATTACHMENT3_SANPAI_NAME_COL)
End Function

' frmSubconSelector.frm ÇçƒóòópÇµÇƒéYã∆îpä¸ï®èàóùé{ê›ÇëIëÇµÅA
' EóÒ(é{ê›ñº)ÅEJóÒ(ëŒâûílÅAã∆é“É}ÉXÉ^BóÒ)Ç÷èëÇ´çûÇﬁÅB
Public Sub RequestAttachment3SanpaiFacilitySelection(ByVal ws As Worksheet, ByVal target As Range)
    If ws Is Nothing Or target Is Nothing Then Exit Sub

    Dim rowIndex As Long
    rowIndex = target.Row

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Sub

    Dim branchName As String
    branchName = GetBasicInfoCellText(wsInfo, BASIC_INFO_BRANCH_CELL)

    Dim names As Collection
    Dim companionMap As Object
    Set names = New Collection
    Set companionMap = CreateObject("Scripting.Dictionary")
    companionMap.CompareMode = vbTextCompare

    If Not LoadSanpaiFacilityMasterRows(branchName, names, companionMap) Then
        MsgBox "éYã∆îpä¸ï®èàóùé{ê›ÇÃÉ}ÉXÉ^Ç™éÊìæÇ≈Ç´Ç‹ÇπÇÒÇ≈ÇµÇΩÅB", vbExclamation
        Exit Sub
    End If
    If names.Count = 0 Then
        MsgBox "éYã∆îpä¸ï®èàóùé{ê›ÇÃÉ}ÉXÉ^Ç…ÉfÅ[É^Ç™Ç†ÇËÇ‹ÇπÇÒÅB", vbExclamation
        Exit Sub
    End If

    Dim arr() As String
    ReDim arr(1 To names.Count)
    Dim i As Long
    For i = 1 To names.Count
        arr(i) = CStr(names(i))
    Next i

    Dim f As New frmSubconSelector
    f.Caption = Attachment3SanpaiFacilityCaptionText()
    f.SetCompanies arr
    f.Show vbModal

    Dim confirmed As Boolean, chosen As String
    confirmed = f.confirmed
    chosen = f.SelectedCompany
    Unload f
    If Not confirmed Or chosen = "" Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    ws.Cells(rowIndex, ATTACHMENT3_SANPAI_NAME_COL).MergeArea.Cells(1, 1).value = chosen

    Dim companionValue As String
    If companionMap.Exists(chosen) Then companionValue = CStr(companionMap(chosen))
    ws.Cells(rowIndex, ATTACHMENT3_SANPAI_VALUE_COL).MergeArea.Cells(1, 1).value = companionValue

    Application.EnableEvents = prevEvents
End Sub

' ã∆é“É}ÉXÉ^(ëSé–î≈).xlsx ÇÃéxìXñºÉVÅ[ÉgÇ©ÇÁ A59à»â∫(é{ê›ñº)ÅEB59à»â∫(ëŒâûíl)Çì«Ç›çûÇﬁÅB
Private Function LoadSanpaiFacilityMasterRows(ByVal branchName As String, _
                                              ByVal names As Collection, _
                                              ByVal companionMap As Object) As Boolean
    Dim connection As Object
    Dim recordset As Object
    On Error GoTo Cleanup

    If Trim$(branchName) = "" Then GoTo Cleanup

    Dim masterPath As String
    masterPath = mod_Construction_BasicTotals.ResolveVendorMasterPath()
    If masterPath = "" Then GoTo Cleanup

    Set connection = CommonOpenExcelAdoConnection(masterPath)
    If connection Is Nothing Then GoTo Cleanup

    Dim actualSheetName As String
    actualSheetName = mod_Construction_OutputLayout.FindAdoWorksheetName(connection, branchName)
    If actualSheetName = "" Then GoTo Cleanup

    Dim tableRangeName As String
    tableRangeName = "[" & Replace$(actualSheetName, "]", "]]") & "$A" & _
                      SANPAI_FACILITY_MASTER_START_ROW & ":B10000]"

    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT F1, F2 FROM " & tableRangeName, connection, 0, 1, 1

    Do Until recordset.EOF
        Dim nameVal As String, companionVal As String
        nameVal = Trim$(CommonNzText(recordset.fields(0).value))
        companionVal = Trim$(CommonNzText(recordset.fields(1).value))
        If nameVal <> "" Then
            names.Add nameVal
            If Not companionMap.Exists(nameVal) Then companionMap.Add nameVal, companionVal
        End If
        recordset.MoveNext
    Loop

    LoadSanpaiFacilityMasterRows = True

Cleanup:
    If Err.Number <> 0 Then Err.Clear
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
End Function

' "éYã∆îpä¸ï®èàóùé{ê›ëIë"
Private Function Attachment3SanpaiFacilityCaptionText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H7523, &H696D, &H5EC3, &H68C4, &H7269, _
                                      &H51E6, &H7406, &H65BD, &H8A2D, &H9078, &H629E)
    End If
    Attachment3SanpaiFacilityCaptionText = cached
End Function

' ===== ï éÜáV J54:M54 éYîpçsJRã‡äzçáåv =====
' é{çséwé¶èë(çHéñ)/é{çsí ímèë(çHéñ)ÉVÅ[ÉgÇÃéYîpçs(é{çHâÔé–Ç™ìñèâëIëÇ≈Ç´Ç»Ç©Ç¡ÇΩçs)ÇÃ
' JRã‡äzóÒçáåvÇÅAåÖãÊêÿÇËêÆêîÇ≈ J54:M54(åãçáÉZÉã)Ç÷ì]ãLÇ∑ÇÈÅB
Private Sub RefreshAttachment3SanpaiJrTotal(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim total As Double
    total = mod_Construction_BasicTotals.SumSanpaiJrAmount()

    Dim writeCell As Range
    Set writeCell = ws.Range("J54").MergeArea.Cells(1, 1)
    writeCell.value = total
    writeCell.NumberFormat = "#,##0;-#,##0;"
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' ?øΩ?øΩ?øΩ?øΩ?øΩZ?øΩ?øΩ?øΩŒâÔøΩ?øΩÃíl?øΩ]?øΩL(?øΩt?øΩH?øΩ?øΩ?øΩg?øΩK?øΩp?øΩA?øΩK?øΩv?øΩ…âÔøΩ?øΩ?øΩ?øΩƒè„â∫?øΩ?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)
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

' ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ∆ÇÔøΩ?øΩƒÇÃì]?øΩL(?øΩ?øΩ?øΩX?øΩR?øΩ[?øΩh?øΩ?øΩ?øΩA?øΩ?øΩ?øΩt?øΩ?øΩœäÔøΩ?øΩ?øΩh?øΩ?øΩ)
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

' ?øΩ?øΩ?øΩt?øΩÃì]?øΩL(?øΩa?øΩ?øΩ\?øΩ?øΩ?øΩ`?øΩ?øΩ?øΩA?øΩ?øΩ?øΩ?øΩ?øΩZ?øΩ?øΩ?øΩÃÇÔøΩ?øΩﬂíÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)
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

' ?øΩ?çé“óp?øΩV?øΩ[?øΩgwsTarget?øΩ∆ìÔøΩ?øΩ?øΩO?øΩ?øΩ?øΩ[?øΩv?øΩÃìÔøΩ?øΩ?æç◊ÉV?øΩ[?øΩg?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
Private Function ResolveBreakdownSheetNameFromTarget(ByVal wsTarget As Worksheet) As String
    Dim baseName As String, aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(wsTarget, baseName, aliasText) Then Exit Function
    If aliasText = "" Then Exit Function
    Dim nm As String
    nm = mod_OrderTpl_Shared.OrderTplBuildSheetName(mod_OrderTpl_Shared.OrderTplBaseNameBreakdownText(), aliasText)
    If mod_OrderTpl_Shared.OrderTplSheetExists(nm) Then ResolveBreakdownSheetNameFromTarget = nm
End Function

' ?øΩ?øΩ?øΩ?æç◊ÉV?øΩ[?øΩg?øΩ?øΩA?øΩ?øΩ≈??øΩ?øΩx?øΩ?øΩ?øΩs?øΩ?øΩT?øΩ?øΩ?øΩA?øΩ?øΩ?øΩ?øΩQ?øΩ?øΩ?øΩ‘ÇÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩ¬ÇÔøΩ?øΩ?øΩ»ÇÔøΩ?øΩ?øΩŒãÔøΩ)
Private Function BuildBreakdownQFormula(ByVal sheetName As String, ByVal labelText As String) As String
    Dim q As String
    q = "'" & Replace$(sheetName, "'", "''") & "'"
    BuildBreakdownQFormula = "=IFERROR(INDEX(" & q & "!Q:Q,MATCH(""" & labelText & """," & q & "!A:A,0)),"""")"
End Function

' ?øΩ?øΩ?øΩ?øΩ(?øΩE?øΩl)?øΩÃì]?øΩL(BizUD?øΩS?øΩV?øΩb?øΩN?øΩE?øΩ?øΩ?øΩ?øΩÿÇÔøΩ)
Private Sub WriteHeaderFormulaRight(ByVal target As Range, ByVal formulaText As String)
    Dim writeCell As Range
    Set writeCell = target.MergeArea.Cells(1, 1)
    writeCell.Formula = formulaText
    writeCell.Font.Name = BASIC_INFO_REF_FONT_NAME
    writeCell.NumberFormat = "#,##0;-#,##0;"
    writeCell.HorizontalAlignment = xlRight
    writeCell.VerticalAlignment = xlCenter
End Sub

' "?øΩ?øΩ?øΩv"
Private Function ContractorGrandTotalLabelText() As String
    ContractorGrandTotalLabelText = ChrW$(&H5408) & ChrW$(&H8A08)
End Function

' "?øΩv"
Private Function ContractorNetTotalLabelText() As String
    ContractorNetTotalLabelText = ChrW$(&H8A08)
End Function

' "?øΩ?øΩ?øΩ?øΩ?øΩ"
Private Function ContractorTaxLabelText() As String
    ContractorTaxLabelText = ChrW$(&H6D88) & ChrW$(&H8CBB) & ChrW$(&H7A0E)
End Function

' "?øΩ@?øΩ@?øΩa"(?øΩS?øΩp?øΩ?øΩ2?øΩ?øΩ + ?øΩa?øΩB?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ M11:V11 ?øΩÃñÔøΩ?øΩ?øΩ?øΩ÷ït?øΩ—ÇÔøΩ?øΩ?øΩ)
Private Function ContractorHonorificSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H3000) & ChrW$(&H3000) & ChrW$(&H6BBF)
    End If
    ContractorHonorificSuffixText = cached
End Function
