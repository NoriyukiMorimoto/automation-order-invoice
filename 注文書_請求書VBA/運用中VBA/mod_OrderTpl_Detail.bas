Option Explicit

' ?øΩ?øΩ?øΩ?æç◊ÉV?øΩ[?øΩg?øΩÃñÔøΩ?øΩ◊ïÔøΩ?øΩ]?øΩL?øΩG?øΩ?øΩ?øΩW?øΩ?øΩ?øΩB
' ?øΩ{?øΩH?øΩw?øΩ?øΩ?øΩ?øΩ(?øΩH?øΩ?øΩ)/?øΩ{?øΩH?øΩ ím?øΩ?øΩ(?øΩH?øΩ?øΩ)?øΩE?øΩ{?øΩH?øΩw?øΩ?øΩ?øΩ?øΩ(?øΩn?øΩ?øΩ)/?øΩ{?øΩH?øΩ ím?øΩ?øΩ(?øΩn?øΩ?øΩ)?øΩÃéÊçû?øΩœÇ›ÉV?øΩ[?øΩg?øΩ?øΩ?øΩ?øΩ
' ?øΩ{?øΩH?øΩ?øΩ?øΩ(?øΩ∆é“É}?øΩX?øΩ^A?øΩ?øΩ ?øΩ∆é“ñÔøΩ)?øΩ≈íÔøΩ?øΩo?øΩ?øΩ?øΩA?øΩ_?øΩ?øΩ?øΩ?øΩÊñº_?øΩ«óÔøΩ?øΩ?øΩ?øΩÃÉZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩP?øΩ Ç≈ì]?øΩL?øΩ?øΩ?øΩ?øΩB
' ?øΩH?øΩ?øΩ?øΩÊï™?øΩ?øΩ?øΩO?øΩ?øΩ?øΩH?øΩ?øΩ?øΩÃèÍçá: ?øΩH?øΩ?øΩ?øΩV?øΩ[?øΩg ?øΩ?øΩ 2?øΩs?øΩ?øΩ ?øΩ?øΩ ?øΩn?øΩ⁄ÉV?øΩ[?øΩg(?øΩO?øΩ?øΩ?øΩËå≥?øΩ?øΩ–Ç≈íÔøΩ?øΩo)?øΩB
' ?øΩH?øΩ?øΩ?øΩÊï™?øΩ?øΩ?øΩn?øΩ⁄çH?øΩ?øΩ?øΩÃèÍçá: ?øΩn?øΩ⁄ÉV?øΩ[?øΩg(?øΩn?øΩ⁄âÔøΩ–Ç≈íÔøΩ?øΩo)?øΩÃÇ›ÅB
' ?øΩ?øΩ?øΩC?øΩ?øΩ?øΩ?øΩ: CHANGELOG.md ?øΩQ?øΩ?øΩ

Private Const SRC_FIELD_SEIRI As Long = 1
Private Const SRC_FIELD_TYPE As Long = 2
Private Const SRC_FIELD_DAYNIGHT As Long = 3
Private Const SRC_FIELD_UNIT As Long = 4
Private Const SRC_FIELD_QTY As Long = 5
Private Const SRC_FIELD_PRICE As Long = 6

Private Const FMT_GROUP_GENERAL As Long = 0
Private Const FMT_GROUP_INTEGER As Long = 1
Private Const FMT_GROUP_DECIMAL As Long = 2
Private Const SUMMARY_EXTRA_ROWS As Long = 5   ' ?øΩl?øΩ?øΩ/?øΩv/?øΩ?øΩ?øΩ?øΩ?øΩ/?øΩ?øΩ?øΩv/?øΩr?øΩ?øΩ?øΩp?øΩ?øΩ
Private Const SUMMARY_NUMBER_FORMAT As String = "#,##0"
Private Const SUMMARY_ZERO_HIDE_FORMAT As String = "#,##0;-#,##0;"
Private Const DETAIL_AMOUNT_NUMBER_FORMAT As String = "#,##0;-#,##0;"            ' ?øΩ?øΩ?øΩ?øΩÿÇÔøΩE?øΩ[?øΩ?øΩ?øΩ?øΩ\?øΩ?øΩ
Private Const DETAIL_QTY_DECIMAL_NUMBER_FORMAT As String = "#,##0.00;-#,##0.00;" ' ?øΩ?øΩ?øΩ?øΩ2?øΩ?øΩ?øΩE?øΩ[?øΩ?øΩ?øΩ?øΩ\?øΩ?øΩ
Private Const DETAIL_FONT_SIZE As Double = 11#
Private Const DETAIL_COL_A_WIDTH As Double = 7#
Private Const DETAIL_COL_QTY_4 As Long = 15          ' O?øΩ?øΩ(?øΩ?øΩ4?øΩ?øΩ?øΩ óÔøΩ)
Private Const DETAIL_COL_PRICE_4 As Long = 16        ' P?øΩ?øΩ(?øΩ?øΩ4?øΩP?øΩ?øΩ?øΩ?øΩ)
Private Const DETAIL_COL_PRICE_BASE As Long = 7        ' G?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ?øΩP?øΩ?øΩ)

Private Function DiscountLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5024, &H5F15)
    End If
    DiscountLabelText = cached
End Function

Private Function NetTotalLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H8A08)
    End If
    NetTotalLabelText = cached
End Function

Private Function TaxLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6D88, &H8CBB, &H7A0E)
    End If
    TaxLabelText = cached
End Function

Private Function GrandTotalLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5408, &H8A08)
    End If
    GrandTotalLabelText = cached
End Function

' ?øΩ?øΩ?øΩ?æç◊ÉV?øΩ[?øΩg?øΩÃñÔøΩ?øΩ◊ïÔøΩ(11?øΩs?øΩ⁄Å`?øΩ?øΩ?øΩv?øΩs?øΩÃíÔøΩ?øΩO)?øΩ÷ì]?øΩL?øΩ?øΩ?øΩ?øΩ
Public Sub ApplyBreakdownDetails(ByVal wsBreakdown As Worksheet, _
                                 ByVal vendorName As String, _
                                 ByVal officialName As String, _
                                 ByVal branchName As String, _
                                 ByVal workTypeText As String)
    If wsBreakdown Is Nothing Then Exit Sub
    If CommonNormalizeText(vendorName) = "" And CommonNormalizeText(officialName) = "" Then Exit Sub

    On Error GoTo ErrorHandler

    ' ?øΩ∆é“É}?øΩX?øΩ^?øΩÃñÔøΩ?øΩ?øΩ(?øΩ?øΩ?øΩK?øΩ?øΩ/?øΩ?øΩ?øΩÃÇÃÇ«ÇÔøΩ?øΩ?øΩÃï\?øΩL?øΩ≈ÇÔøΩ?øΩ?øΩv?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)
    Dim aliasMap As Object
    Set aliasMap = mod_Construction_BasicTotals.GetVendorAliasMap(branchName)

    Dim targetKeys As Object
    Set targetKeys = CreateObject("Scripting.Dictionary")
    targetKeys.CompareMode = vbTextCompare
    AddTargetKey targetKeys, vendorName, aliasMap
    AddTargetKey targetKeys, officialName, aliasMap
    If targetKeys.Count = 0 Then Exit Sub

    Dim subtotalRow As Long
    ResetDetailArea wsBreakdown, subtotalRow
    If subtotalRow = 0 Then
        mod_OrderTpl_Shared.OrderTplLog "subtotal row not found: " & wsBreakdown.Name
        Exit Sub
    End If

    Dim isWeldingWork As Boolean
    isWeldingWork = (InStr(1, CommonNormalizeText(workTypeText), WELDING_WORK_TYPE_KEYWORD, vbTextCompare) > 0)

    ' ?øΩ?øΩ?øΩo: ?øΩO?øΩ?øΩ?øΩH?øΩ?øΩ?øΩÕçH?øΩ?øΩ?øΩV?øΩ[?øΩg(A?øΩ?øΩ ?øΩ{?øΩH?øΩ∆éÔøΩ)+?øΩn?øΩ⁄ÉV?øΩ[?øΩg(B?øΩ?øΩ ?øΩO?øΩ?øΩ?øΩËå≥?øΩ?øΩ?øΩ)?øΩA
    '       ?øΩn?øΩ⁄çH?øΩ?øΩ?øΩÕón?øΩ⁄ÉV?øΩ[?øΩg(A?øΩ?øΩ ?øΩn?øΩ⁄âÔøΩ?øΩ)?øΩÃÇÔøΩ
    Dim detailT0 As Double
    detailT0 = mod_Construction_Import_Shared.LogCIStart()

    Dim worksSections As Collection
    Dim weldSections As Collection
    If isWeldingWork Then
        Set weldSections = CollectSourceSections( _
            mod_OrderTpl_Shared.OrderTplFindWeldingSourceSheet(), WELD_COL_WELDING_VENDOR, targetKeys, aliasMap, True)
    Else
        Set worksSections = CollectSourceSections( _
            mod_OrderTpl_Shared.OrderTplFindWorksSourceSheet(), COL_VENDOR, targetKeys, aliasMap, False)
        Set weldSections = CollectSourceSections( _
            mod_OrderTpl_Shared.OrderTplFindWeldingSourceSheet(), WELD_COL_TRACK_VENDOR, targetKeys, aliasMap, True)
    End If
    mod_Construction_Import_Shared.LogCIElapsed "ApplyBreakdownDetails: collect", detailT0

    Dim worksLineCount As Long
    Dim weldLineCount As Long
    worksLineCount = CountBlockLines(worksSections)
    weldLineCount = CountBlockLines(weldSections)

    Dim totalLines As Long
    totalLines = worksLineCount + weldLineCount
    If worksLineCount > 0 And weldLineCount > 0 Then totalLines = totalLines + 2
    ' ñæç◊çsÇ™ñ≥Ç≠ÇƒÇ‡èWåvÉtÉbÉ^Å[(è¨åv/ílà¯/åv/è¡îÔê≈/çáåv)ÇÕïKÇ∏ê∂ê¨Ç∑ÇÈ(ìùçáÉtÉçÅ[)ÅB
    ' ñæç◊Ç™óLÇÈèÍçáÇÃÇ›ÅAçsà íuí≤êÆÇ∆ñæç◊ì]ãLÇé¿é{Ç∑ÇÈÅB
    If totalLines <= 0 Then mod_OrderTpl_Shared.OrderTplLog "no detail rows (summary only): " & wsBreakdown.Name & " vendor=" & vendorName

    If totalLines > 0 Then
    PositionSubtotalRow wsBreakdown, subtotalRow, totalLines

    ' ?øΩ]?øΩL?øΩl?øΩÃëg?øΩ›óÔøΩ?øΩ?øΩ
    Dim valuesAF() As Variant
    Dim valuesO() As Variant
    Dim valuesP() As Variant
    ReDim valuesAF(1 To totalLines, 1 To 7)
    ReDim valuesO(1 To totalLines, 1 To 1)
    ReDim valuesP(1 To totalLines, 1 To 1)

    Dim headerLineRows As Collection
    Dim integerLineRows As Collection
    Dim decimalLineRows As Collection
    Set headerLineRows = New Collection
    Set integerLineRows = New Collection
    Set decimalLineRows = New Collection

    Dim lineCursor As Long
    lineCursor = 0
    WriteBlockLines worksSections, False, lineCursor, valuesAF, valuesO, valuesP, _
                    headerLineRows, integerLineRows, decimalLineRows
    If worksLineCount > 0 And weldLineCount > 0 Then lineCursor = lineCursor + 2
    WriteBlockLines weldSections, True, lineCursor, valuesAF, valuesO, valuesP, _
                    headerLineRows, integerLineRows, decimalLineRows

    ' ?øΩÍä??øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
    Dim startRow As Long
    Dim endRow As Long
    startRow = ORDER_TPL_DETAIL_START_ROW
    endRow = startRow + totalLines - 1

    wsBreakdown.Range(wsBreakdown.Cells(startRow, 1), wsBreakdown.Cells(endRow, 7)).value = valuesAF
    ' ?øΩ?øΩ4?øΩO?øΩ?øΩ?øΩ[?øΩv: O?øΩ?øΩ=?øΩ?øΩ?øΩ ÅAP?øΩ?øΩ=?øΩP?øΩ?øΩ?øΩ?øΩ?øΩ¬ï ÉZ?øΩ?øΩ?øΩ÷ì]?øΩL?øΩ?øΩ?øΩ?øΩ(2?øΩ?øΩÍä??øΩ?øΩ?øΩ?øΩP?øΩ?øΩ÷ìÔøΩ?øΩ?øΩP?øΩ?øΩ?øΩ?øΩO?øΩ?øΩ÷ÇÔøΩ?øΩ?øΩ?øΩÍç??øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)
    wsBreakdown.Range(wsBreakdown.Cells(startRow, DETAIL_COL_QTY_4), wsBreakdown.Cells(endRow, DETAIL_COL_QTY_4)).value = valuesO
    wsBreakdown.Range(wsBreakdown.Cells(startRow, DETAIL_COL_PRICE_4), wsBreakdown.Cells(endRow, DETAIL_COL_PRICE_4)).value = valuesP

    ' ?øΩ?øΩ?øΩz?øΩ?øΩ(H/K/N/Q)?øΩ÷êÔøΩ?øΩ?øΩ(?øΩ?øΩ?øΩ Å~?øΩP?øΩ?øΩ)?øΩ?øΩ›íËÇ∑?øΩ?øΩ(?øΩ?øΩ?øΩv?øΩs?øΩÃíÔøΩ?øΩO?øΩ‹ÇÔøΩ)
    Dim formulaColumn As Variant
    For Each formulaColumn In Array("H", "K", "N", "Q")
        wsBreakdown.Range(formulaColumn & startRow & ":" & formulaColumn & (subtotalRow - 1)).FormulaR1C1 = _
            "=RC[-2]*RC[-1]"
    Next formulaColumn

    CleanupDetailAmountEmptyRows wsBreakdown, startRow, subtotalRow - 1

    detailT0 = mod_Construction_Import_Shared.LogCIStart()
    ApplyDetailFormats wsBreakdown, startRow, endRow, headerLineRows, integerLineRows, decimalLineRows

    ' ÈáëÈ°çÂ??(H/K/N/Q)„ÅÆÂ∞èÊï∞Ë°®Á§∫„ÅØ„ÄÅApplyDetailFormats „ÅÆÊï¥Êï∞Êõ∏Âºè„?ÆÂæå„Å´ÈÅ©Áî®„Åó„Å¶ÂÑ™ÂÖà„Åï„Åõ„Çã
    ApplyDetailAmountColumnFormats wsBreakdown, startRow, subtotalRow - 1

    ' B:C?øΩ?øΩÃåÔøΩ?øΩ?øΩ(11?øΩs?øΩ⁄Å`?øΩ?øΩ?øΩv?øΩs?øΩÃíÔøΩ?øΩO?øΩB?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩo?øΩ?øΩ?øΩs?øΩ?øΩA:D?øΩ?øΩ?øΩ?øΩ?øΩÃÇÔøΩ?øΩﬂèÔøΩ?øΩ?øΩ)
    MergeDetailNameColumns wsBreakdown, startRow, subtotalRow - 1, headerLineRows
    mod_Construction_Import_Shared.LogCIElapsed "ApplyBreakdownDetails: format/merge", detailT0
    End If

    detailT0 = mod_Construction_Import_Shared.LogCIStart()
    BuildSummaryBlock wsBreakdown, subtotalRow, totalLines
    mod_Construction_Import_Shared.LogCIElapsed "ApplyBreakdownDetails: summary", detailT0

    ' 11?øΩs?øΩ⁄à»ç~(?øΩW?øΩv?øΩu?øΩ?øΩ?øΩb?øΩN?øΩ‹ÇÔøΩ)?øΩÃÉt?øΩH?øΩ?øΩ?øΩg?øΩT?øΩC?øΩY?øΩ?øΩ11?øΩ|?øΩC?øΩ?øΩ?øΩg?øΩ÷ìÔøΩ?øΩÍÇ∑?øΩ?øΩ
    wsBreakdown.Range(wsBreakdown.Cells(ORDER_TPL_DETAIL_START_ROW, 1), _
                      wsBreakdown.Cells(subtotalRow + SUMMARY_EXTRA_ROWS, 17)).Font.Size = DETAIL_FONT_SIZE

    ' A?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ)?øΩÃóÒïùÇÔøΩ?øΩ≈íÔøΩl7.00?øΩ…ê›íËÇ∑?øΩ?øΩ
    wsBreakdown.Columns(1).ColumnWidth = DETAIL_COL_A_WIDTH

    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownDetails done: " & wsBreakdown.Name & _
        " rows=" & totalLines & " works=" & worksLineCount & " weld=" & weldLineCount
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownDetails error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' ?øΩ?øΩ?øΩ◊ïÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ(?øΩe?øΩ?øΩ?øΩv?øΩ?øΩ?øΩ[?øΩg?øΩs?øΩ?øΩ)?øΩ÷ñﬂÇÔøΩ?øΩA?øΩl?øΩ?øΩ?øΩN?øΩ?øΩ?øΩA?øΩ?øΩ?øΩ?øΩ
Private Sub ResetDetailArea(ByVal wsBreakdown As Worksheet, ByRef subtotalRow As Long)
    subtotalRow = FindSubtotalRow(wsBreakdown)
    If subtotalRow = 0 Then Exit Sub

    ' ?øΩ»ëO?øΩ…êÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩW?øΩv?øΩu?øΩ?øΩ?øΩb?øΩN(?øΩl?øΩ?øΩ/?øΩv/?øΩ?øΩ?øΩ?øΩ?øΩ/?øΩ?øΩ?øΩv/?øΩr?øΩ?øΩ?øΩp?øΩ?íçs)?øΩ?øΩ?øΩÌèúÔøΩ?øΩ?øΩ?øΩ
    If IsSummaryLabelRow(wsBreakdown, subtotalRow + 1, DiscountLabelText()) Then
        wsBreakdown.Rows((subtotalRow + 1) & ":" & (subtotalRow + SUMMARY_EXTRA_ROWS)).Delete
    End If

    Dim defaultLastRow As Long
    defaultLastRow = ORDER_TPL_DETAIL_START_ROW + ORDER_TPL_DETAIL_DEFAULT_ROWS - 1

    If subtotalRow - 1 > defaultLastRow Then
        wsBreakdown.Rows((defaultLastRow + 1) & ":" & (subtotalRow - 1)).Delete
        subtotalRow = defaultLastRow + 1
    End If

    Dim clearLastRow As Long
    clearLastRow = subtotalRow - 1

    ' ?øΩO?øΩ?øΩ]?øΩL?øΩÃÉZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩo?øΩ?øΩ(A:D?øΩ?øΩ?øΩ?øΩ)?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩƒÇÔøΩ?øΩ?øΩl?øΩ?øΩ?øΩN?øΩ?øΩ?øΩA?øΩ?øΩ?øΩ?øΩ
    mod_VendorBlockLayout.SafeUnmergeRange wsBreakdown.Range( _
        wsBreakdown.Cells(ORDER_TPL_DETAIL_START_ROW, 1), wsBreakdown.Cells(clearLastRow, 4))

    With wsBreakdown
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 1), .Cells(clearLastRow, 7)).ClearContents
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, DETAIL_COL_QTY_4), .Cells(clearLastRow, DETAIL_COL_QTY_4)).ClearContents
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, DETAIL_COL_PRICE_4), .Cells(clearLastRow, DETAIL_COL_PRICE_4)).ClearContents
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 6), .Cells(clearLastRow, 7)).NumberFormat = "General"
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, DETAIL_COL_QTY_4), .Cells(clearLastRow, DETAIL_COL_PRICE_4)).NumberFormat = "General"
    End With
End Sub

' A?øΩ?øΩÃÅu?øΩ?øΩ?øΩv?øΩv?øΩs?øΩ?øΩT?øΩ?øΩ
Private Function FindSubtotalRow(ByVal wsBreakdown As Worksheet) As Long
    Dim subtotalText As String
    subtotalText = mod_OrderTpl_Shared.OrderTplSubtotalLabelText()

    Dim r As Long
    For r = ORDER_TPL_DETAIL_START_ROW To ORDER_TPL_DETAIL_START_ROW + 2000
        If StrComp(CommonRemoveAllSpaces(CommonNormalizeText(CommonNzText(wsBreakdown.Cells(r, 1).value))), _
                   subtotalText, vbTextCompare) = 0 Then
            FindSubtotalRow = r
            Exit Function
        End If
    Next r
End Function

' ÂΩìË©≤Ê•≠ËÄ?„ÅÆÂÜ?Ë®≥ÊòéÁ¥∞„Ç∑„Éº„Éà„?Æ„ÄåË®à„ÄçË°?(Â∞èË®?+2)„ÅÆQÂàóÂÄ§„ÇíËøî„Åô„ÄÇÂ≠òÂú®„Åó„Å™„Åë„Çå„Å∞ Empty„Ä?
' Âü∫Êú¨ÊÉ?Â†±33Ë°?(Â•ëÁ¥?ÈáëÈ°çÁ®éÊäú)„ÇíÂ??Ë®≥ÊòéÁ¥∞„ÄåË®à„Äç„Å´‰∏ÄËá¥„Åï„Åõ„Çã„Åü„ÇÅ„Å´‰Ωø„Å?„Ä?
Public Function OrderTplGetBreakdownNetTotalQForVendor(ByVal wsInfo As Worksheet, _
                                                       ByVal vendorIndex As Long) As Variant
    On Error GoTo Done
    If wsInfo Is Nothing Then Exit Function

    Dim branchName As String
    branchName = CommonNormalizeText(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))

    Dim companyName As String
    companyName = mod_OrderTpl_Shared.OrderTplGetVendorCompanyName(wsInfo, vendorIndex)
    If companyName = "" Then Exit Function

    Dim vendorName As String, aliasText As String, workText As String
    If Not mod_OrderTpl_Shared.OrderTplResolveVendorMasterInfo(branchName, companyName, _
            vendorName, aliasText, workText) Then Exit Function
    If aliasText = "" Then Exit Function

    Dim sheetName As String
    sheetName = mod_OrderTpl_Shared.OrderTplBuildSheetName( _
        mod_OrderTpl_Shared.OrderTplBaseNameBreakdownText(), aliasText)
    If Not mod_OrderTpl_Shared.OrderTplSheetExists(sheetName) Then Exit Function

    Dim wsBreakdown As Worksheet
    Set wsBreakdown = ThisWorkbook.Worksheets(sheetName)

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(wsBreakdown)
    If subtotalRow = 0 Then Exit Function

    ' „ÄåË®à„ÄçË°å„?ÆÊï∞Âº?(=Â∞èË®?+ÂÄ§Âºï„Åç)„ÇíÊúÄÊñ∞Âåñ„Åó„Å¶„Åã„ÇâQÂÄ§„ÇíÂèñÂæó„Åô„Ç?
    On Error Resume Next
    wsBreakdown.Calculate
    On Error GoTo Done
    OrderTplGetBreakdownNetTotalQForVendor = wsBreakdown.Range("Q" & (subtotalRow + 2)).value

Done:
End Function

' ?øΩ?øΩ?øΩv?øΩs?øΩ?øΩ?øΩu?øΩ≈èI?øΩf?øΩ[?øΩ^?øΩs?øΩ?øΩ?øΩ?øΩ2?øΩs?øΩ?ØÇÔøΩ?øΩ íu?øΩv?øΩ÷à⁄ìÔøΩ?øΩ?øΩ?øΩ?øΩ(?øΩs?øΩ?øΩ?øΩÕçs?øΩ}?øΩ?øΩ?øΩA?øΩ]?øΩ?øΩÕçs?øΩÌè?)
Private Sub PositionSubtotalRow(ByVal wsBreakdown As Worksheet, _
                                ByRef subtotalRow As Long, _
                                ByVal totalLines As Long)
    Dim desiredRow As Long
    desiredRow = ORDER_TPL_DETAIL_START_ROW + totalLines + 2

    If desiredRow > subtotalRow Then
        Dim insertCount As Long
        insertCount = desiredRow - subtotalRow

        ' ?øΩ?øΩ?øΩ?øΩ?øΩÕèÔøΩ?øΩ?øΩs?øΩ?øΩ?øΩ?øΩp?øΩ?øΩ?øΩB?øΩ?øΩ?øΩz?øΩ?øΩ?øΩ?øΩ(H/K/N/Q)?øΩÕì]?øΩL?øΩ?øΩ…ëS?øΩs?øΩ÷ê›íËÇ∑?øΩ?øΩ
        wsBreakdown.Rows(subtotalRow & ":" & (subtotalRow + insertCount - 1)).Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ElseIf desiredRow < subtotalRow Then
        wsBreakdown.Rows(desiredRow & ":" & (subtotalRow - 1)).Delete
    End If

    subtotalRow = desiredRow
End Sub

' ?øΩW?øΩv?øΩu?øΩ?øΩ?øΩb?øΩN?øΩ?øΩ?øΩ\?øΩz?øΩ?øΩ?øΩ?øΩ: ?øΩ?øΩ?øΩv/?øΩl?øΩ?øΩ/?øΩv/?øΩ?øΩ?øΩ?øΩ?øΩ/?øΩ?øΩ?øΩv + ?øΩr?øΩ?øΩ?øΩp?øΩ?íçs
Private Sub BuildSummaryBlock(ByVal wsBreakdown As Worksheet, _
                              ByVal subtotalRow As Long, _
                              ByVal totalLines As Long)
    Dim lastDataRow As Long
    ' ñæç◊0çsÇ≈Ç‡SUMîÕàÕÇ™îjí]ÇµÇ»Ç¢ÇÊÇ§ÅAèWåvíºëOçs(è¨åvçsÇÃ1çsè„)ÇäÓèÄÇ…Ç∑ÇÈ
    lastDataRow = subtotalRow - 1

    ' ?øΩ?øΩ?øΩv?øΩs?øΩÃâÔøΩ?øΩ?øΩ4?øΩs(?øΩl?øΩ?øΩ/?øΩv/?øΩ?øΩ?øΩ?øΩ?øΩ/?øΩ?øΩ?øΩv)+?øΩr?øΩ?øΩ?øΩp?øΩ?íçs?øΩ?øΩ}?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
    wsBreakdown.Rows((subtotalRow + 1) & ":" & (subtotalRow + SUMMARY_EXTRA_ROWS)).Insert _
        Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' ?øΩ?øΩ?øΩ?øΩ≈óÔøΩ(?øΩ?øΩ{?øΩ?øΩ?øΩB34?øΩu?øΩ?øΩ?øΩ?øΩ?øΩ(10%)?øΩF?øΩv?øΩÃÉJ?øΩb?øΩR?øΩ?øΩ)?øΩ?øΩ?øΩÊìæ?øΩ?øΩ?øΩ?øΩ
    Dim taxRateText As String
    taxRateText = Trim$(Str$(mod_Construction_BasicTotals.ResolveBasicInfoTaxRate( _
        CommonGetBasicInfoWorksheet())))

    ' ?øΩ?øΩ?øΩx?øΩ?øΩ(A:C?øΩ?øΩ?øΩ?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)
    WriteSummaryLabel wsBreakdown, subtotalRow, mod_OrderTpl_Shared.OrderTplSubtotalLabelText()
    WriteSummaryLabel wsBreakdown, subtotalRow + 1, DiscountLabelText()
    WriteSummaryLabel wsBreakdown, subtotalRow + 2, NetTotalLabelText()
    WriteSummaryLabel wsBreakdown, subtotalRow + 3, TaxLabelText()
    WriteSummaryLabel wsBreakdown, subtotalRow + 4, GrandTotalLabelText()

    ' H/K/N/Q ?øΩ?øΩ: ?øΩ?øΩ?øΩv=SUM(11:?øΩ≈èI?øΩf?øΩ[?øΩ^?øΩs)?øΩA?øΩl?øΩ?øΩ=-MOD(?øΩ?øΩ?øΩv,1000)?øΩA?øΩv=?øΩ?øΩ?øΩv+?øΩl?øΩ?øΩ?øΩA
    ' ?øΩ?øΩ?øΩ?øΩ?øΩ=ROUNDDOWN(?øΩv*?øΩ≈óÔøΩ,0)?øΩA?øΩ?øΩ?øΩv=?øΩv+?øΩ?øΩ?øΩ?øΩ?øΩ
    Dim summaryColumns As Variant
    summaryColumns = Array("H", "K", "N", "Q")

    Dim i As Long
    For i = LBound(summaryColumns) To UBound(summaryColumns)
        Dim colLetter As String
        colLetter = CStr(summaryColumns(i))

        wsBreakdown.Range(colLetter & subtotalRow).Formula = _
            "=ROUND(SUM(" & colLetter & ORDER_TPL_DETAIL_START_ROW & ":" & colLetter & lastDataRow & "),0)"
        wsBreakdown.Range(colLetter & (subtotalRow + 1)).Formula = _
            "=-MOD(" & colLetter & subtotalRow & ",1000)"
        wsBreakdown.Range(colLetter & (subtotalRow + 2)).Formula = _
            "=" & colLetter & subtotalRow & "+" & colLetter & (subtotalRow + 1)
        wsBreakdown.Range(colLetter & (subtotalRow + 3)).Formula = _
            "=ROUNDDOWN(" & colLetter & (subtotalRow + 2) & "*" & taxRateText & ",0)"
        wsBreakdown.Range(colLetter & (subtotalRow + 4)).Formula = _
            "=" & colLetter & (subtotalRow + 2) & "+" & colLetter & (subtotalRow + 3)
    Next i

    ApplySummaryNumberFormats wsBreakdown, subtotalRow

    ' ?øΩt?øΩH?øΩ?øΩ?øΩg
    wsBreakdown.Range(wsBreakdown.Cells(subtotalRow, 1), _
                      wsBreakdown.Cells(subtotalRow + 4, 17)).Font.Name = BASIC_INFO_REF_FONT_NAME

    ApplySummaryBorders wsBreakdown, subtotalRow
End Sub

' ?øΩ?øΩ?øΩx?øΩ?øΩ?øΩZ?øΩ?øΩ(A:C?øΩ?øΩ?øΩ?øΩ?øΩE?øΩ„â∫?øΩ?øΩ?øΩE?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)?øΩ÷ÇÃèÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
Private Sub WriteSummaryLabel(ByVal wsBreakdown As Worksheet, _
                              ByVal rowIndex As Long, _
                              ByVal labelText As String)
    Dim labelRange As Range
    Set labelRange = wsBreakdown.Range(wsBreakdown.Cells(rowIndex, 1), wsBreakdown.Cells(rowIndex, 4))

    On Error Resume Next
    If Not labelRange.MergeCells Then
        labelRange.UnMerge
        labelRange.Merge
    End If
    On Error GoTo 0

    labelRange.Cells(1, 1).value = labelText
    labelRange.HorizontalAlignment = xlCenter
    labelRange.VerticalAlignment = xlCenter
End Sub

' ?øΩ?øΩ?øΩ◊ïÔøΩ?øΩ{?øΩW?øΩv?øΩu?øΩ?øΩ?øΩb?øΩN?øΩÃår?øΩ?øΩ?øΩ?çï?øΩ?øΩ?øΩ?øΩe?øΩ?øΩ?øΩv?øΩ?øΩ?øΩ[?øΩg?øΩÃól?øΩ?øΩ?øΩ…çÔøΩ?øΩÌÇπ?øΩƒàÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩB
' ?øΩŒè€ÇÕñÔøΩ?øΩ◊äJ?øΩn?øΩs?øΩ`?øΩ?øΩ?øΩv?øΩs(?øΩ?øΩ?øΩv+4)?øΩ‹Ç≈ÅEA:Q ?øΩS?øΩ?øΩB?øΩs?øΩÃë}?øΩ?øΩ/?øΩÌèúÔøΩ≈ïÔøΩ?øΩÍÇΩ?øΩr?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩZ?øΩb?øΩg?øΩ?øΩ?øΩƒçƒï`?øΩ?øΩB
'   ?øΩ?øΩ?øΩr?øΩ?øΩ: ?øΩ?øΩ?øΩ◊ÇÔøΩ xlHairline?øΩB?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ ?øΩ?øΩ?øΩv?øΩs?øΩÃèÔøΩ=?øΩ?øΩd?øΩ?øΩ / ?øΩ?øΩ?øΩv?øΩs?øΩÃâÔøΩ=xlMedium?øΩB
'   ?øΩc?øΩr?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ): ?øΩO?øΩg?øΩ?øΩ(A)?øΩE?øΩO?øΩg?øΩE(Q)?øΩE?øΩ?øΩ?øΩz?øΩO?øΩ?øΩ?øΩ[?øΩv?øΩ?øΩ?øΩE(F/I/L/O ?øΩÃçÔøΩ)?øΩB
'   ?øΩc?øΩr?øΩ?øΩ(?øΩ◊êÔøΩ): ?øΩ?øΩ?øΩÃëÔøΩ?øΩÃóÒã´äE(E/G/H/J/K/M/N/P/Q ?øΩÃçÔøΩ)?øΩB
Private Sub ApplySummaryBorders(ByVal wsBreakdown As Worksheet, _
                                ByVal subtotalRow As Long)
    Dim firstRow As Long, grandTotalRow As Long
    firstRow = ORDER_TPL_DETAIL_START_ROW
    grandTotalRow = subtotalRow + 4

    Dim fullRange As Range
    Set fullRange = wsBreakdown.Range(wsBreakdown.Cells(firstRow, 1), _
                                      wsBreakdown.Cells(grandTotalRow, 17))

    ' ?øΩ?øΩ?øΩ?øΩ?øΩr?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?äÉZ?øΩb?øΩg
    fullRange.Borders(xlEdgeLeft).LineStyle = xlNone
    fullRange.Borders(xlEdgeRight).LineStyle = xlNone
    fullRange.Borders(xlEdgeTop).LineStyle = xlNone
    fullRange.Borders(xlEdgeBottom).LineStyle = xlNone
    fullRange.Borders(xlInsideHorizontal).LineStyle = xlNone
    fullRange.Borders(xlInsideVertical).LineStyle = xlNone

    ' ?øΩ?øΩ?øΩr?øΩ?øΩ: ?øΩ?øΩ?øΩ◊ÇÔøΩ xlHairline(?øΩs?øΩ‘ÅE?øΩ?øΩ[)
    fullRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    fullRange.Borders(xlInsideHorizontal).Weight = xlHairline
    fullRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    fullRange.Borders(xlEdgeTop).Weight = xlHairline

    ' ?øΩc?øΩr?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ): ?øΩO?øΩg?øΩ?øΩ(A)?øΩE?øΩ?øΩ?øΩz?øΩO?øΩ?øΩ?øΩ[?øΩv?øΩ?øΩ?øΩE(F/I/L/O ?øΩÃçÔøΩ)
    Dim mediumCol As Variant
    For Each mediumCol In Array(1, 6, 9, 12, 15)
        With wsBreakdown.Range(wsBreakdown.Cells(firstRow, CLng(mediumCol)), _
                               wsBreakdown.Cells(grandTotalRow, CLng(mediumCol))).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Next mediumCol
    ' ?øΩO?øΩg?øΩE(Q)=?øΩ?øΩ?øΩ?øΩ
    With wsBreakdown.Range(wsBreakdown.Cells(firstRow, 17), _
                           wsBreakdown.Cells(grandTotalRow, 17)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    ' ?øΩc?øΩr?øΩ?øΩ(?øΩ◊êÔøΩ): ?øΩ?øΩ?øΩÃëÔøΩ?øΩÃóÒã´äE(E/G/H/J/K/M/N/P/Q ?øΩÃçÔøΩ)
    Dim thinCol As Variant
    For Each thinCol In Array(5, 7, 8, 10, 11, 13, 14, 16, 17)
        With wsBreakdown.Range(wsBreakdown.Cells(firstRow, CLng(thinCol)), _
                               wsBreakdown.Cells(grandTotalRow, CLng(thinCol))).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next thinCol

    ' ?øΩ?øΩ?øΩv?øΩs?øΩÃèÔøΩ=?øΩ?øΩd?øΩ?øΩ (A:Q)
    wsBreakdown.Range(wsBreakdown.Cells(subtotalRow, 1), _
                      wsBreakdown.Cells(subtotalRow, 17)).Borders(xlEdgeTop).LineStyle = xlDouble

    ' ?øΩ?øΩ?øΩv?øΩs?øΩÃâÔøΩ=?øΩ?øΩ?øΩ?øΩ (A:Q)
    With wsBreakdown.Range(wsBreakdown.Cells(grandTotalRow, 1), _
                           wsBreakdown.Cells(grandTotalRow, 17)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
End Sub

' A?øΩ?øΩZ?øΩ?øΩ?øΩÃï\?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?™éw?øΩËÉâÔøΩx?øΩ?øΩ?øΩ∆àÔøΩv?øΩ?øΩ?øΩÈÇ©
Private Function IsSummaryLabelRow(ByVal wsBreakdown As Worksheet, _
                                   ByVal rowIndex As Long, _
                                   ByVal labelText As String) As Boolean
    IsSummaryLabelRow = (StrComp( _
        CommonRemoveAllSpaces(CommonNormalizeText(CommonNzText(wsBreakdown.Cells(rowIndex, 1).value))), _
        labelText, vbTextCompare) = 0)
End Function

' ?øΩÊçûÔøΩœÇ›ÉV?øΩ[?øΩg?øΩ?øΩ?øΩ?øΩ{?øΩH?øΩ?øΩ–Ç≈íÔøΩ?øΩo?øΩ?øΩ?øΩA?øΩ_?øΩ?øΩ?øΩ?øΩÊñº_?øΩ«óÔøΩ?øΩ?øΩ?øΩÃÉZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩÍóóÔøΩ?øΩ?øΩ?øΩ?øΩ
' ?øΩﬂÇÔøΩl: Collection of Array(?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩo?øΩ?øΩ, ?øΩsCollection)?øΩB?øΩs = Array(?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ, ?øΩH?øΩ?øΩ?øΩ?øΩ?øΩ, ?øΩ?øΩ?øΩ?øΩ?øΩ, ?øΩP?øΩ?øΩ, ?øΩ?øΩ?øΩ?øΩ, JR?øΩP?øΩ?øΩ)
Private Function CollectSourceSections(ByVal wsSource As Worksheet, _
                                       ByVal vendorColumn As Long, _
                                       ByVal targetKeys As Object, _
                                       ByVal aliasMap As Object, _
                                       ByVal isWeldingSource As Boolean) As Collection
    If wsSource Is Nothing Then Exit Function

    Dim columnOffset As Long
    If isWeldingSource Then columnOffset = WELDING_OUTPUT_COL_OFFSET

    Dim lastRow As Long
    lastRow = mod_Construction_LineMapping.GetLastDataRow(wsSource, COL_SEIRI + columnOffset)
    If lastRow < 2 Then Exit Function

    Dim sourceValues As Variant
    sourceValues = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, COL_JR_PRICE + columnOffset)).value

    ' ?øΩ{?øΩH?øΩ?øΩ–ï íP?øΩ?øΩ?øΩ?øΩ(?øΩw?øΩb?øΩ_?øΩ[?øΩu?øΩ?øΩ–ñÔøΩ+?øΩP?øΩ?øΩ?øΩv)?øΩ?ºäÒÇπÇ≈ìÔøΩ?øΩËÇ∑?øΩ?øΩ
    Dim vendorPriceColumn As Long
    vendorPriceColumn = FindVendorUnitPriceColumn(wsSource, targetKeys, aliasMap)
    If vendorPriceColumn = 0 Then
        mod_OrderTpl_Shared.OrderTplLog "vendor price column not found on: " & wsSource.Name
    End If

    Dim vendorPriceValues As Variant
    If vendorPriceColumn > 0 Then
        vendorPriceValues = wsSource.Range(wsSource.Cells(2, vendorPriceColumn), _
                                           wsSource.Cells(lastRow, vendorPriceColumn)).value
    End If

    Dim sectionKeys As Collection
    Set sectionKeys = New Collection

    Dim sectionMap As Object
    Set sectionMap = CreateObject("Scripting.Dictionary")
    sectionMap.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To UBound(sourceValues, 1)
        ' éYîpèàóùçs(ìñèâÇÕé{çHâÔé–ÇëIëÇ≈Ç´Ç»Ç©Ç¡ÇΩçHéÌ)ÇÕì‡ñÛñæç◊Ç÷ì]ãLÇµÇ»Ç¢
        If InStr(1, CommonRemoveAllSpaces(CommonNzText(sourceValues(i, COL_TYPE + columnOffset))), _
                 SANPAI_KEYWORD, vbTextCompare) > 0 Then GoTo NextSourceRow

        Dim rowVendorKey As String
        rowVendorKey = mod_Construction_BasicTotals.ResolveVendorCanonicalKey( _
            CommonNzText(sourceValues(i, vendorColumn)), aliasMap)
        If rowVendorKey <> "" And targetKeys.Exists(rowVendorKey) Then
            Dim sectionLabel As String
            sectionLabel = BuildSectionLabel( _
                CommonNzText(sourceValues(i, COL_LINE + columnOffset)), _
                CommonNzText(sourceValues(i, COL_MGR + columnOffset)), _
                isWeldingSource)

            Dim sectionKey As String
            sectionKey = CommonRemoveAllSpaces(sectionLabel)

            Dim sectionRows As Collection
            If sectionMap.Exists(sectionKey) Then
                Set sectionRows = sectionMap(sectionKey)
            Else
                Set sectionRows = New Collection
                sectionMap.Add sectionKey, sectionRows
                sectionKeys.Add Array(sectionLabel, sectionKey)
            End If

            ' ?øΩP?øΩ?øΩ?øΩÕé{?øΩH?øΩ?øΩ–ï íP?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ–ñÔøΩ+?øΩP?øΩ?øΩ)?øΩ?øΩ?øΩ?øΩÊìæ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩW?øΩJ?øΩE?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩÕãÔøΩ)
            Dim vendorPriceValue As Variant
            vendorPriceValue = Empty
            If vendorPriceColumn > 0 Then vendorPriceValue = vendorPriceValues(i, 1)

            sectionRows.Add Array( _
                sourceValues(i, COL_SEIRI + columnOffset), _
                sourceValues(i, COL_TYPE + columnOffset), _
                sourceValues(i, COL_DAYNIGHT + columnOffset), _
                sourceValues(i, COL_UNIT + columnOffset), _
                sourceValues(i, COL_QTY + columnOffset), _
                vendorPriceValue)
        End If
NextSourceRow:
    Next i

    If sectionKeys.Count = 0 Then Exit Function

    Dim result As Collection
    Set result = New Collection

    Dim keyPair As Variant
    For Each keyPair In sectionKeys
        Dim sortedRows As Collection
        Set sortedRows = SortSectionRows(sectionMap(CStr(keyPair(1))))
        result.Add Array(CStr(keyPair(0)), sortedRows)
    Next keyPair

    Set CollectSourceSections = result
End Function

' ?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩo?øΩ?øΩ: ?øΩ_?øΩ?øΩ?øΩ?øΩÊñº(?øΩ⁄îÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)[_?øΩ?øΩ?øΩ[?øΩ?øΩ?øΩn?øΩ?øΩ]_?øΩ«óÔøΩ?øΩ?øΩ?øΩ?øΩ
Private Function BuildSectionLabel(ByVal lineText As String, _
                                   ByVal managerRoomText As String, _
                                   ByVal isWeldingSource As Boolean) As String
    Dim label As String
    label = mod_OrderTpl_Shared.OrderTplStripLineSuffix(lineText, isWeldingSource)
    If isWeldingSource Then
        label = label & "_" & mod_OrderTpl_Shared.OrderTplRailWeldingLabelText()
    End If

    ' ?øΩ{?øΩs?øΩ ím?øΩ?øΩ?øΩ…ÇÕä«óÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩﬂÅA?øΩ?øΩÃèÍçá?øΩÕÅu_?øΩv?øΩ?øΩt?øΩ?øΩ?øΩ»ÇÔøΩ
    Dim managerRoomName As String
    managerRoomName = CommonNormalizeText(managerRoomText)
    If Len(managerRoomName) > 0 Then label = label & "_" & managerRoomName

    BuildSectionLabel = label
End Function

' ?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ\?øΩ[?øΩg: ?øΩ?øΩ1?øΩL?øΩ[ ?øΩ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ)?øΩA?øΩ?øΩ2?øΩL?øΩ[ ?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ(?øΩ?øΩ?øΩ?øΩ)
Private Function SortSectionRows(ByVal sectionRows As Collection) As Collection
    Dim rowCount As Long
    rowCount = sectionRows.Count

    If rowCount <= 1 Then
        Set SortSectionRows = sectionRows
        Exit Function
    End If

    Dim rowsArray() As Variant
    ReDim rowsArray(1 To rowCount)

    Dim i As Long
    For i = 1 To rowCount
        rowsArray(i) = sectionRows(i)
    Next i

    Dim j As Long
    For i = 2 To rowCount
        Dim currentRow As Variant
        currentRow = rowsArray(i)
        j = i - 1
        Do While j >= 1
            If CompareDetailRows(rowsArray(j), currentRow) > 0 Then
                rowsArray(j + 1) = rowsArray(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        rowsArray(j + 1) = currentRow
    Next i

    Dim result As Collection
    Set result = New Collection
    For i = 1 To rowCount
        result.Add rowsArray(i)
    Next i
    Set SortSectionRows = result
End Function

Private Function CompareDetailRows(ByVal leftRow As Variant, ByVal rightRow As Variant) As Long
    Dim leftRank As Long
    Dim rightRank As Long
    leftRank = DayNightRank(CommonNzText(leftRow(SRC_FIELD_DAYNIGHT - 1)))
    rightRank = DayNightRank(CommonNzText(rightRow(SRC_FIELD_DAYNIGHT - 1)))

    If leftRank <> rightRank Then
        CompareDetailRows = Sgn(leftRank - rightRank)
        Exit Function
    End If

    Dim leftSeiri As Double
    Dim rightSeiri As Double
    leftSeiri = Val(StrConv(CommonNzText(leftRow(SRC_FIELD_SEIRI - 1)), vbNarrow))
    rightSeiri = Val(StrConv(CommonNzText(rightRow(SRC_FIELD_SEIRI - 1)), vbNarrow))
    CompareDetailRows = Sgn(leftSeiri - rightSeiri)
End Function

Private Function DayNightRank(ByVal dayNightText As String) As Long
    If InStr(1, CommonNormalizeText(dayNightText), mod_OrderTpl_Shared.OrderTplDayFirstText(), vbTextCompare) > 0 Then
        DayNightRank = 0
    Else
        DayNightRank = 1
    End If
End Function

Private Function CountBlockLines(ByVal sections As Collection) As Long
    If sections Is Nothing Then Exit Function

    Dim total As Long
    Dim sectionItem As Variant
    For Each sectionItem In sections
        If total > 0 Then total = total + 1   ' ?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ‘ÇÃãÛîíçs
        total = total + 1                     ' ?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩo?øΩ?øΩ?øΩs
        total = total + sectionItem(1).Count  ' ?øΩf?øΩ[?øΩ^?øΩs
    Next sectionItem
    CountBlockLines = total
End Function

' ?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩÍóóÔøΩ?øΩ]?øΩL?øΩp?øΩz?øΩ?øΩ÷ìW?øΩJ?øΩ?øΩ?øΩ?øΩ
Private Sub WriteBlockLines(ByVal sections As Collection, _
                            ByVal isWeldingSource As Boolean, _
                            ByRef lineCursor As Long, _
                            ByRef valuesAF() As Variant, _
                            ByRef valuesO() As Variant, _
                            ByRef valuesP() As Variant, _
                            ByVal headerLineRows As Collection, _
                            ByVal integerLineRows As Collection, _
                            ByVal decimalLineRows As Collection)
    If sections Is Nothing Then Exit Sub

    Dim isFirstSection As Boolean
    isFirstSection = True

    Dim sectionItem As Variant
    For Each sectionItem In sections
        If Not isFirstSection Then lineCursor = lineCursor + 1   ' ?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ‘ÇÃãÛîíçs
        isFirstSection = False

        lineCursor = lineCursor + 1
        valuesAF(lineCursor, 1) = CStr(sectionItem(0))
        headerLineRows.Add lineCursor

        Dim dataRow As Variant
        For Each dataRow In sectionItem(1)
            lineCursor = lineCursor + 1
            valuesAF(lineCursor, 1) = dataRow(SRC_FIELD_SEIRI - 1)
            valuesAF(lineCursor, 2) = dataRow(SRC_FIELD_TYPE - 1)
            valuesAF(lineCursor, 4) = dataRow(SRC_FIELD_DAYNIGHT - 1)
            valuesAF(lineCursor, 5) = dataRow(SRC_FIELD_UNIT - 1)
            valuesAF(lineCursor, 6) = NumericOrValue(dataRow(SRC_FIELD_QTY - 1))
            valuesAF(lineCursor, 7) = NumericOrValue(dataRow(SRC_FIELD_PRICE - 1))
            valuesO(lineCursor, 1) = valuesAF(lineCursor, 6)
            valuesP(lineCursor, 1) = valuesAF(lineCursor, 7)

            Select Case QuantityFormatGroup(CommonNzText(dataRow(SRC_FIELD_UNIT - 1)), isWeldingSource)
                Case FMT_GROUP_INTEGER
                    integerLineRows.Add lineCursor
                Case FMT_GROUP_DECIMAL
                    decimalLineRows.Add lineCursor
            End Select
        Next dataRow
    Next sectionItem
End Sub

' ?øΩ?øΩ?øΩ ÇÃï\?øΩ?øΩ?øΩ`?øΩ?øΩ: ?øΩn?øΩ⁄ÇÕåÔøΩ?øΩÿÇÔøΩ»ÇÔøΩ(General)?øΩA?øΩH?øΩ?øΩ?øΩÕíP?øΩ Ç…âÔøΩ?øΩ?øΩ?øΩƒêÔøΩ?øΩ?øΩ(0)/?øΩ?øΩ?øΩ?øΩ2?øΩ?øΩ(0.00)
Private Function QuantityFormatGroup(ByVal unitText As String, ByVal isWeldingSource As Boolean) As Long
    If isWeldingSource Then
        QuantityFormatGroup = FMT_GROUP_GENERAL
    ElseIf mod_OrderTpl_Shared.OrderTplIsIntegerUnit(unitText) Then
        QuantityFormatGroup = FMT_GROUP_INTEGER
    ElseIf mod_OrderTpl_Shared.OrderTplIsDecimalUnit(unitText) Then
        QuantityFormatGroup = FMT_GROUP_DECIMAL
    Else
        QuantityFormatGroup = FMT_GROUP_GENERAL
    End If
End Function

Private Function NumericOrValue(ByVal sourceValue As Variant) As Variant
    If IsNumeric(sourceValue) Then
        NumericOrValue = CDbl(sourceValue)
    Else
        NumericOrValue = sourceValue
    End If
End Function

' ?øΩ?øΩ?øΩ?øΩ?øΩK?øΩp: ?øΩt?øΩH?øΩ?øΩ?øΩg?øΩE?øΩz?øΩu?øΩE?øΩk?øΩ?øΩ?øΩ\?øΩ?øΩ?øΩE?øΩ\?øΩ?øΩ?øΩ`?øΩ?øΩ
Private Sub ApplyDetailFormats(ByVal wsBreakdown As Worksheet, _
                               ByVal startRow As Long, _
                               ByVal endRow As Long, _
                               ByVal headerLineRows As Collection, _
                               ByVal integerLineRows As Collection, _
                               ByVal decimalLineRows As Collection)
    With wsBreakdown
        ' ?øΩ?øΩ?øΩÕîÕàÕëS?øΩÃÇÔøΩ BIZ UD?øΩS?øΩV?øΩb?øΩN?øΩ?øΩ
        .Range(.Cells(startRow, 1), .Cells(endRow, 17)).Font.Name = BASIC_INFO_REF_FONT_NAME

        ' A?øΩ?øΩ: ?øΩ?øΩ?øΩ?øΩ?øΩ‘çÔøΩ?øΩÕçÔøΩ?øΩE?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
        .Range(.Cells(startRow, 1), .Cells(endRow, 1)).HorizontalAlignment = xlCenter

        ' B?øΩ?øΩ: ?øΩH?øΩ?øΩ?øΩ?øΩﬁÇÕçÔøΩ?øΩl?øΩﬂÅE?øΩk?øΩ?øΩ?øΩ?øΩ?øΩƒëS?øΩÃÇÔøΩ\?øΩ?øΩ
        With .Range(.Cells(startRow, 2), .Cells(endRow, 2))
            .HorizontalAlignment = xlLeft
            .WrapText = False
            .ShrinkToFit = True
        End With

        ' D?øΩ?øΩ: ?øΩ?øΩ?øΩ?øΩ ÇÕçÔøΩ?øΩE?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ
        .Range(.Cells(startRow, 4), .Cells(endRow, 4)).HorizontalAlignment = xlCenter

        ' F?øΩ`Q?øΩ?øΩ: ?øΩ?øΩ?øΩ?øΩÿÇÔøΩE?øΩ[?øΩ?øΩ?øΩl?øΩ?øΩ\?øΩ?øΩ?øΩB?øΩ?øΩ?øΩ óÔøΩ(F/I/L/O)?øΩÃäÔøΩ?øΩ?øΩ=?øΩ?øΩ?øΩ?øΩ?øΩA?øΩ?øΩ?øΩ?øΩ3?øΩ?øΩ?øΩÃçs?øΩÕåÔøΩ≈è„èë?øΩ?øΩ
        .Range(.Cells(startRow, 6), .Cells(endRow, 17)).NumberFormat = DETAIL_AMOUNT_NUMBER_FORMAT
    End With

    ' ?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩo?øΩ?øΩ?øΩs(?øΩ_?øΩ?øΩ?øΩ?øΩÊñº?øΩE?øΩ«óÔøΩ?øΩ?øΩ?øΩ?øΩ): A:D?øΩ?øΩ?øΩ?øΩ?øΩE?øΩ?øΩ?øΩl?øΩﬂÅE?øΩk?øΩ?øΩ?øΩ?øΩ?øΩƒÉZ?øΩ?øΩ?øΩ?øΩ?øΩ…éÔøΩ?øΩﬂÇÔøΩ
    Dim lineIndex As Variant
    For Each lineIndex In headerLineRows
        Dim headerRowIndex As Long
        headerRowIndex = startRow + CLng(lineIndex) - 1
        With wsBreakdown.Range(wsBreakdown.Cells(headerRowIndex, 1), wsBreakdown.Cells(headerRowIndex, 4))
            .Merge
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .ShrinkToFit = True
        End With
    Next lineIndex

    ' ?øΩ?øΩ?øΩ óÔøΩ(F/I/L/O): ?øΩP?øΩ ÇÔøΩ?øΩ?øΩ?øΩ?øΩ2?øΩ?øΩ?øΩŒèÔøΩ(m/m3/M/?øΩu/t)?øΩÃçs?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ2?øΩ?øΩ?øΩ\?øΩ?øΩ?øΩ÷è„èë?øΩ?øΩ
    ApplyNumberFormatToLineRows wsBreakdown, startRow, 6, decimalLineRows, DETAIL_QTY_DECIMAL_NUMBER_FORMAT
    ApplyNumberFormatToLineRows wsBreakdown, startRow, 9, decimalLineRows, DETAIL_QTY_DECIMAL_NUMBER_FORMAT
    ApplyNumberFormatToLineRows wsBreakdown, startRow, 12, decimalLineRows, DETAIL_QTY_DECIMAL_NUMBER_FORMAT
    ApplyNumberFormatToLineRows wsBreakdown, startRow, 15, decimalLineRows, DETAIL_QTY_DECIMAL_NUMBER_FORMAT
End Sub

Private Sub ApplyNumberFormatToLineRows(ByVal wsBreakdown As Worksheet, _
                                        ByVal startRow As Long, _
                                        ByVal columnIndex As Long, _
                                        ByVal lineRows As Collection, _
                                        ByVal formatText As String)
    If lineRows Is Nothing Then Exit Sub
    If lineRows.Count = 0 Then Exit Sub

    Dim targetRange As Range
    Dim unionCount As Long

    Dim lineIndex As Variant
    For Each lineIndex In lineRows
        Dim targetCell As Range
        Set targetCell = wsBreakdown.Cells(startRow + CLng(lineIndex) - 1, columnIndex)
        If targetRange Is Nothing Then
            Set targetRange = targetCell
        Else
            Set targetRange = Union(targetRange, targetCell)
        End If
        unionCount = unionCount + 1
        If unionCount >= 100 Then
            targetRange.NumberFormat = formatText
            Set targetRange = Nothing
            unionCount = 0
        End If
    Next lineIndex

    If Not targetRange Is Nothing Then targetRange.NumberFormat = formatText
End Sub

' ?øΩÊçûÔøΩ?øΩ?øΩV?øΩ[?øΩg?øΩ?øΩ1?øΩs?øΩ⁄ÇÔøΩ?øΩ?øΩu?øΩ?øΩ–ñÔøΩ+?øΩP?øΩ?øΩ?øΩv?øΩw?øΩb?øΩ_?øΩ[?øΩ?øΩ?ºäÒÇπÇ≈íT?øΩ?øΩ
Private Function FindVendorUnitPriceColumn(ByVal wsSource As Worksheet, _
                                           ByVal targetKeys As Object, _
                                           ByVal aliasMap As Object) As Long
    If wsSource Is Nothing Then Exit Function
    If targetKeys Is Nothing Then Exit Function

    Dim lastColumn As Long
    lastColumn = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    If lastColumn < 1 Then Exit Function

    Dim suffixText As String
    suffixText = mod_OrderTpl_Shared.OrderTplUnitPriceHeaderSuffixText()

    ' ÉwÉbÉ_Å[çsÇàÍäáì«çûÇµÇƒÉZÉãâùïúÇîÇØÇÈ
    Dim headerArr As Variant
    If lastColumn = 1 Then
        ReDim headerArr(1 To 1, 1 To 1)
        headerArr(1, 1) = wsSource.Cells(1, 1).value
    Else
        headerArr = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastColumn)).value
    End If

    Dim c As Long
    For c = 1 To lastColumn
        Dim headerText As String
        headerText = CommonRemoveAllSpaces(CommonNormalizeText(CommonNzText(headerArr(1, c))))
        If Len(headerText) > Len(suffixText) Then
            If StrComp(Right$(headerText, Len(suffixText)), suffixText, vbTextCompare) = 0 Then
                Dim vendorKey As String
                vendorKey = mod_Construction_BasicTotals.ResolveVendorCanonicalKey( _
                    Left$(headerText, Len(headerText) - Len(suffixText)), aliasMap)
                If vendorKey <> "" Then
                    If targetKeys.Exists(vendorKey) Then
                        FindVendorUnitPriceColumn = c
                        Exit Function
                    End If
                End If
            End If
        End If
    Next c
End Function

' B:C?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩ◊ïÔøΩ?øΩÃÇ›ÅB?øΩZ?øΩN?øΩV?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩo?øΩ?øΩ?øΩs?øΩ?øΩA:D?øΩ?øΩ?øΩ?øΩ?øΩÃÇÔøΩ?øΩﬂèÔøΩ?øΩO)
Private Sub MergeDetailNameColumns(ByVal wsBreakdown As Worksheet, _
                                   ByVal firstRow As Long, _
                                   ByVal lastRow As Long, _
                                   ByVal headerLineRows As Collection)
    If lastRow < firstRow Then Exit Sub

    Dim headerRowMap As Object
    Set headerRowMap = CreateObject("Scripting.Dictionary")

    Dim lineIndex As Variant
    If Not headerLineRows Is Nothing Then
        For Each lineIndex In headerLineRows
            headerRowMap(firstRow + CLng(lineIndex) - 1) = True
        Next lineIndex
    End If

    ' èëéÆÇÕàÍäáÅAMerge ÇÕñ¢åãçáçsÇÃÇ›
    Dim formatRange As Range
    Set formatRange = wsBreakdown.Range(wsBreakdown.Cells(firstRow, 2), wsBreakdown.Cells(lastRow, 3))
    formatRange.WrapText = False
    formatRange.ShrinkToFit = True
    formatRange.HorizontalAlignment = xlLeft

    Dim r As Long
    For r = firstRow To lastRow
        If Not headerRowMap.Exists(r) Then
            With wsBreakdown.Range(wsBreakdown.Cells(r, 2), wsBreakdown.Cells(r, 3))
                If Not .MergeCells Then .Merge
            End With
        End If
    Next r
End Sub

' ?øΩ?øΩ?øΩ?æç◊ÉV?øΩ[?øΩg?øΩÃëO?øΩ?øΩ(I?øΩ?øΩ)/?øΩ?øΩ?øΩ?øΩ(L?øΩ?øΩ)/?øΩ?øΩ4?øΩ?øΩ?øΩ?øΩ(O?øΩ?øΩ)?øΩ÷ÇÃêÔøΩ?øΩ ìÔøΩ?øΩÕÇÔøΩ?øΩƒéÔøΩ?øΩ?øΩ?øΩA?øΩP?øΩ?øΩ?øΩ?øΩ(J/M/P)?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩÕÇÔøΩ?øΩ?øΩB
' I12/L12 ?øΩÃìÔøΩ?øΩÕóL?øΩ?øΩ?øΩ…âÔøΩ?øΩ?øΩ?øΩƒèW?øΩv?øΩu?øΩ?øΩ?øΩb?øΩN?øΩ?øΩ K/N ?øΩ?øΩ[?øΩ?øΩ?øΩ\?øΩ?øΩ?øΩ?øΩ?øΩX?øΩV?øΩ?øΩ?øΩ?øΩB
' ThisWorkbook.Workbook_SheetChange ?øΩ?øΩ?øΩ?øΩƒÇŒÇÔøΩ?øΩ
Public Sub HandleBreakdownQuantityCellChange(ByVal sh As Object, ByVal target As Range)
    If TypeName(sh) <> "Worksheet" Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error GoTo Quiet

    Dim ws As Worksheet
    Set ws = sh

    Dim baseName As String
    Dim aliasText As String
    If Not mod_OrderTpl_Shared.OrderTplIsGeneratedSheet(ws, baseName, aliasText) Then Exit Sub
    If StrComp(baseName, mod_OrderTpl_Shared.OrderTplBaseNameBreakdownText(), vbTextCompare) <> 0 Then Exit Sub

    If Not Intersect(target, ws.Range("I12:L12")) Is Nothing Then
        RefreshBreakdownSummaryNumberFormats ws
    End If

    Dim hitRange As Range
    Set hitRange = Intersect(target, Union(ws.Columns(9), ws.Columns(12), ws.Columns(DETAIL_COL_QTY_4)))
    If hitRange Is Nothing Then Exit Sub

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    Set hitRange = Intersect(hitRange, _
        ws.Range(ws.Cells(ORDER_TPL_DETAIL_START_ROW, 1), ws.Cells(subtotalRow - 1, DETAIL_COL_PRICE_4)))
    If hitRange Is Nothing Then Exit Sub

    Dim prevEnableEvents As Boolean
    prevEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim changedCell As Range
    For Each changedCell In hitRange.Cells
        If changedCell.Column = 9 Or changedCell.Column = 12 Or changedCell.Column = DETAIL_COL_QTY_4 Then
            Dim priceCell As Range
            Set priceCell = ws.Cells(changedCell.Row, changedCell.Column + 1)

            Dim quantityValue As Variant
            quantityValue = changedCell.value
            If IsNumeric(quantityValue) And Len(Trim$(CStr(quantityValue))) > 0 Then
                ' ?øΩP?øΩ?øΩ?øΩÕìÔøΩ?øΩ?øΩ?øΩP?øΩ?øΩ(G?øΩ?øΩ=?øΩ?øΩ–ï íP?øΩ?øΩ)?øΩ∆ìÔøΩ?øΩl?øΩ?øΩ?øΩ?øΩÕÇÔøΩ?øΩ?øΩ
                priceCell.value = ws.Cells(changedCell.Row, DETAIL_COL_PRICE_BASE).value
            Else
                priceCell.ClearContents
            End If
        End If
    Next changedCell

    Application.EnableEvents = prevEnableEvents
    Exit Sub

Quiet:
    Application.EnableEvents = True
    Err.Clear
End Sub

' ?øΩ?øΩ?øΩz?øΩ?øΩ(H/K/N/Q)?øΩ÷É[?øΩ?øΩ?øΩ?øΩ\?øΩ?øΩ?øΩÃï\?øΩ?øΩ?øΩ`?øΩ?øΩ?øΩ?øΩK?øΩp?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩv?øΩ?øΩ?øΩO?øΩÃãÔøΩs?øΩ?øΩ?øΩ‹ÇÔøΩ)
Private Sub ApplyDetailAmountColumnFormats(ByVal wsBreakdown As Worksheet, _
                                           ByVal startRow As Long, _
                                           ByVal endRow As Long)
    Dim colLetter As Variant
    For Each colLetter In Array("H", "K", "N", "Q")
        wsBreakdown.Range(CStr(colLetter) & startRow & ":" & CStr(colLetter) & endRow).NumberFormat = _
            DETAIL_AMOUNT_NUMBER_FORMAT
    Next colLetter
End Sub

' A?øΩ?™ãÛóìÇÃçs?øΩÕãÔøΩ?øΩz?øΩ?øΩ?øΩ0?øΩ?øΩ\?øΩ?øΩ?øΩ?øΩ?øΩ»ÇÔøΩ(?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ ÇÔøΩ?øΩN?øΩ?øΩ?øΩA)
Private Sub CleanupDetailAmountEmptyRows(ByVal wsBreakdown As Worksheet, _
                                         ByVal startRow As Long, _
                                         ByVal endRow As Long)
    Dim r As Long
    For r = startRow To endRow
        If Len(Trim$(CommonNzText(wsBreakdown.Cells(r, 1).value))) = 0 Then
            Dim amountCol As Variant
            For Each amountCol In Array(8, 11, 14, 17)
                wsBreakdown.Cells(r, CLng(amountCol)).ClearContents
            Next amountCol
        End If
    Next r
End Sub

' ?øΩW?øΩv?øΩu?øΩ?øΩ?øΩb?øΩN(?øΩ?øΩ?øΩv?øΩ`?øΩ?øΩ?øΩv)?øΩÃãÔøΩ?øΩz?øΩ\?øΩ?øΩ?øΩ`?øΩ?øΩ?øΩBK/N?øΩ?øΩ?øΩ I12/L12 ?øΩÃìÔøΩ?øΩÕóL?øΩ?øΩ?øΩ≈É[?øΩ?øΩ?øΩ\?øΩ?øΩ?øΩ?øΩÿëÔøΩ
Private Sub ApplySummaryNumberFormats(ByVal wsBreakdown As Worksheet, ByVal subtotalRow As Long)
    Dim r1 As Long, r2 As Long
    r1 = subtotalRow
    r2 = subtotalRow + 4

    wsBreakdown.Range("H" & r1 & ":H" & r2).NumberFormat = SUMMARY_ZERO_HIDE_FORMAT
    wsBreakdown.Range("Q" & r1 & ":Q" & r2).NumberFormat = SUMMARY_ZERO_HIDE_FORMAT

    If BreakdownHeaderQtyCellHasValue(wsBreakdown.Range("I12")) Then
        wsBreakdown.Range("K" & r1 & ":K" & r2).NumberFormat = SUMMARY_NUMBER_FORMAT
    Else
        wsBreakdown.Range("K" & r1 & ":K" & r2).NumberFormat = SUMMARY_ZERO_HIDE_FORMAT
    End If

    If BreakdownHeaderQtyCellHasValue(wsBreakdown.Range("L12")) Then
        wsBreakdown.Range("N" & r1 & ":N" & r2).NumberFormat = SUMMARY_NUMBER_FORMAT
    Else
        wsBreakdown.Range("N" & r1 & ":N" & r2).NumberFormat = SUMMARY_ZERO_HIDE_FORMAT
    End If
End Sub

Public Sub RefreshBreakdownSummaryNumberFormats(ByVal wsBreakdown As Worksheet)
    If wsBreakdown Is Nothing Then Exit Sub

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(wsBreakdown)
    If subtotalRow = 0 Then Exit Sub

    ApplySummaryNumberFormats wsBreakdown, subtotalRow
End Sub

' ?øΩA?øΩN?øΩe?øΩB?øΩu?øΩ»ìÔøΩ?øΩ?æç◊ÉV?øΩ[?øΩg?øΩ÷ÅA?øΩr?øΩ?øΩ?øΩEA?øΩ?ùÅE?øΩW?øΩv?øΩ?øΩ?øΩl?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩËìÆ?øΩ≈çƒìK?øΩp?øΩ?øΩ?øΩ?øΩ(?øΩ?øΩ?øΩ?øΩ/?øΩÍä??øΩC?øΩ?øΩ?øΩp)?øΩB
' ?øΩ?øΩ?øΩ?øΩ?øΩœÇ›ÇÃìÔøΩ?øΩ?æç◊ÉV?øΩ[?øΩg?øΩ?øΩ\?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ‘ÇÔøΩ Alt+F8 ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩs?øΩ?øΩ?øΩ?øΩB
Public Sub ApplyBreakdownFormattingToActiveSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then
        MsgBox BreakdownSubtotalNotFoundText(), vbExclamation
        Exit Sub
    End If

    Dim prevScreen As Boolean
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo Restore
    ApplySummaryBorders ws, subtotalRow
    ApplySummaryNumberFormats ws, subtotalRow
    ws.Range(ws.Cells(ORDER_TPL_DETAIL_START_ROW, 1), _
             ws.Cells(subtotalRow + SUMMARY_EXTRA_ROWS, 17)).Font.Size = DETAIL_FONT_SIZE
    ws.Columns(1).ColumnWidth = DETAIL_COL_A_WIDTH

Restore:
    Application.ScreenUpdating = prevScreen
    If Err.Number <> 0 Then
        MsgBox "Err " & Err.Number & ": " & Err.Description, vbExclamation
        Exit Sub
    End If
    MsgBox BreakdownFormattingDoneText() & vbCrLf & ws.Name & " / subtotalRow=" & subtotalRow, vbInformation
End Sub

' "?øΩ?øΩ?øΩv?øΩs?øΩ?øΩ?øΩ?øΩ?øΩ¬ÇÔøΩ?øΩ?øΩ‹ÇÔøΩ?øΩ?øΩB?øΩ?øΩ?øΩ?æç◊ÉV?øΩ[?øΩg?øΩ?øΩ?øΩA?øΩN?øΩe?øΩB?øΩu?øΩ…ÇÔøΩ?øΩƒéÔøΩ?øΩs?øΩ?øΩ?øΩƒÇÔøΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩB"
Private Function BreakdownSubtotalNotFoundText() As String
    BreakdownSubtotalNotFoundText = ChrW$(&H5C0F) & ChrW$(&H8A08) & ChrW$(&H884C) & ChrW$(&H304C) & _
        ChrW$(&H898B) & ChrW$(&H3064) & ChrW$(&H304B) & ChrW$(&H308A) & ChrW$(&H307E) & ChrW$(&H305B) & _
        ChrW$(&H3093) & ChrW$(&H3002) & ChrW$(&H5185) & ChrW$(&H8A33) & ChrW$(&H660E) & ChrW$(&H7D30) & _
        ChrW$(&H30B7) & ChrW$(&H30FC) & ChrW$(&H30C8) & ChrW$(&H3092) & ChrW$(&H30A2) & ChrW$(&H30AF) & _
        ChrW$(&H30C6) & ChrW$(&H30A3) & ChrW$(&H30D6) & ChrW$(&H306B) & ChrW$(&H3057) & ChrW$(&H3066) & _
        ChrW$(&H5B9F) & ChrW$(&H884C) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & _
        ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

' "?øΩr?øΩ?øΩ?øΩE?øΩ?ùÅE?øΩ?øΩ?øΩ?øΩ?øΩ?øΩ?øΩƒìK?øΩp?øΩ?øΩ?øΩ‹ÇÔøΩ?øΩ?øΩ?øΩB"
Private Function BreakdownFormattingDoneText() As String
    BreakdownFormattingDoneText = ChrW$(&H7F6B) & ChrW$(&H7DDA) & ChrW$(&H30FB) & ChrW$(&H5217) & _
        ChrW$(&H5E45) & ChrW$(&H30FB) & ChrW$(&H66F8) & ChrW$(&H5F0F) & ChrW$(&H3092) & ChrW$(&H518D) & _
        ChrW$(&H9069) & ChrW$(&H7528) & ChrW$(&H3057) & ChrW$(&H307E) & ChrW$(&H3057) & ChrW$(&H305F) & _
        ChrW$(&H3002)
End Function

Private Function BreakdownHeaderQtyCellHasValue(ByVal targetCell As Range) As Boolean
    If targetCell Is Nothing Then Exit Function

    Dim valueText As String
    valueText = Trim$(CommonNzText(targetCell.MergeArea.Cells(1, 1).value))
    BreakdownHeaderQtyCellHasValue = (Len(valueText) > 0)
End Function

Private Sub AddTargetKey(ByVal targetKeys As Object, _
                         ByVal nameText As String, _
                         ByVal aliasMap As Object)
    Dim keyText As String
    keyText = mod_Construction_BasicTotals.ResolveVendorCanonicalKey(nameText, aliasMap)
    If keyText = "" Then Exit Sub
    If Not targetKeys.Exists(keyText) Then targetKeys.Add keyText, True
End Sub
