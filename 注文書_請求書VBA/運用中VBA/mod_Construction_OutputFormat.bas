Option Explicit

' ????: CHANGELOG.md ??
' mod_Construction_OutputFormat (split from mod_Construction_Order_Import)

'  ConstructionIntegerNumberFormat
'  Integer display format (negative values in red). ChrW builds [aka].
Public Function ConstructionIntegerNumberFormat() As String
    Static cached As String
    If Len(cached) = 0 Then
        cached = "#,##0;[" & ChrW$(&H8D64) & "]-#,##0"
    End If
    ConstructionIntegerNumberFormat = cached
End Function

' NumberFormatLocal を安全に設定する。結合セルにまたがる範囲や保護セル等で
' 実行時エラー1004が出ても、取込全体を止めずスキップ(表示書式のみの影響)する。
Public Sub SafeSetRangeNumberFormatLocal(ByVal target As Range, ByVal formatText As String)
    On Error Resume Next
    target.NumberFormatLocal = formatText
    If Err.Number <> 0 Then
        LogCI "NumberFormatLocal設定をスキップ(結合/保護セル?): Err " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub ApplySanpaiRowRestrictionsCore(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim fillColor As Long
    fillColor = mod_Construction_BasicTotals.GetSanpaiFillColor()

    Dim restrictedCount As Long
    Dim r As Long
    Dim vendorCol As Variant
    Dim vendorColumns As Collection
    Set vendorColumns = mod_Construction_OutputLayout.OutputSheetVendorColumnsCore(ws)
    For r = 2 To lastRow
        If mod_Construction_BasicTotals.IsSanpaiRow(ws, r) Then
            For Each vendorCol In vendorColumns
                With ws.Cells(r, CLng(vendorCol))
                    .ClearContents
                    .Interior.Color = fillColor
                    With .Validation
                        .Delete
                        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
                             Formula1:="=FALSE"
                        .IgnoreBlank = True
                        .InCellDropdown = False
                        .ShowInput = False
                        .ErrorTitle = "���͕s��"
                        .ErrorMessage = "�Y�p�����̍s�͎{�H��Ђ���͂ł��܂���B"
                        .ShowError = True
                    End With
                End With
            Next vendorCol
            restrictedCount = restrictedCount + 1
        End If
    Next r

    LogCI "�Y�p�����s: A��h��Ԃ��E���͕s��=" & restrictedCount & " �s"
End Sub

Public Sub WriteTotalCells(ByVal ws As Worksheet, ByVal totalRow As Long, _
                            ByVal labelColumn As Long, ByVal labelText As String, _
                            ByVal sumColumn As Long, ByVal sumLastRow As Long, _
                            Optional ByVal useRoundHalfUp As Boolean = False)
    With ws.Cells(totalRow, labelColumn)
        .value = labelText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
    End With

    With ws.Cells(totalRow, sumColumn)
        If useRoundHalfUp Then
            .FormulaR1C1 = "=ROUND(SUM(R2C:R" & sumLastRow & "C),0)"
        Else
            .FormulaR1C1 = "=ROUNDDOWN(SUM(R2C:R" & sumLastRow & "C),0)"
        End If
    End With
    SafeSetRangeNumberFormatLocal ws.Cells(totalRow, sumColumn), ConstructionIntegerNumberFormat()

    DrawDoubleBorder ws.Cells(totalRow, labelColumn), RGB(0, 0, 0)
    DrawDoubleBorder ws.Cells(totalRow, sumColumn), RGB(255, 0, 0)
End Sub

Public Function IsOutputTotalLabelText(ByVal text As String) As Boolean
    Dim normalized As String
    normalized = Trim$(CommonNormalizeText(CommonNzText(text)))
    IsOutputTotalLabelText = (normalized = "JR���v") Or _
        (Len(normalized) >= 2 And Right$(normalized, 2) = "���v")
End Function

Public Sub ClearDoubleBorder(ByVal target As Range)
    Dim edgeId As Variant
    For Each edgeId In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With target.Borders(edgeId)
            .LineStyle = xlLineStyleNone
        End With
    Next edgeId
End Sub

Public Sub ClearOutputTotalPairFormatting(ByVal ws As Worksheet, _
                                           ByVal rowIndex As Long, _
                                           ByVal labelColumn As Long, _
                                           ByVal sumColumn As Long, _
                                           ByVal currentTotalRow As Long, _
                                           ByVal lastDataRow As Long, _
                                           ByVal seiriColumn As Long)
    If rowIndex = currentTotalRow Then Exit Sub
    If Not IsOutputTotalLabelText(CommonNzText(ws.Cells(rowIndex, labelColumn).value)) Then Exit Sub

    ClearDoubleBorder ws.Cells(rowIndex, labelColumn)
    ClearDoubleBorder ws.Cells(rowIndex, sumColumn)
    ws.Cells(rowIndex, labelColumn).ClearContents
    ws.Cells(rowIndex, sumColumn).ClearContents
End Sub

Public Sub ClearStaleOutputTotalFormatting(ByVal ws As Worksheet, _
                                            ByVal lastDataRow As Long, _
                                            ByVal totalRow As Long, _
                                            ByVal subconFirstCol As Long, _
                                            ByVal subconColumnCount As Long, _
                                            Optional ByVal seiriColumn As Long = 0, _
                                            Optional ByVal jrLabelColumn As Long = 0, _
                                            Optional ByVal jrSumColumn As Long = 0)
    Dim scanLastRow As Long
    scanLastRow = lastDataRow + 5
    If Not ws.UsedRange Is Nothing Then
        Dim usedLastRow As Long
        usedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If usedLastRow > scanLastRow Then scanLastRow = usedLastRow
    End If

    If seiriColumn = 0 Then seiriColumn = mod_Construction_OutputLayout.OutputSheetSeiriColumnCore(ws)
    If jrLabelColumn = 0 Then jrLabelColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_JR_PRICE)
    If jrSumColumn = 0 Then jrSumColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_JR_AMOUNT)

    Dim rowIndex As Long
    For rowIndex = 2 To scanLastRow
        ClearOutputTotalPairFormatting ws, rowIndex, jrLabelColumn, jrSumColumn, _
            totalRow, lastDataRow, seiriColumn

        If subconColumnCount > 0 Then
            Dim colIndex As Long
            For colIndex = subconFirstCol To subconFirstCol + subconColumnCount - 1 Step 2
                ClearOutputTotalPairFormatting ws, rowIndex, colIndex, colIndex + 1, _
                    totalRow, lastDataRow, seiriColumn
            Next colIndex
        End If
    Next rowIndex
End Sub

Public Sub DrawDoubleBorder(ByVal target As Range, ByVal lineColor As Long)
    Dim edgeId As Variant
    For Each edgeId In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With target.Borders(edgeId)
            .LineStyle = xlDouble
            .Color = lineColor
        End With
    Next edgeId
End Sub

Public Sub WriteOutputTotalRows( _
    ByVal ws As Worksheet, _
    ByVal vendorNames As Collection, _
    Optional ByVal subconFirstCol As Long = 0, _
    Optional ByVal subconColumnCount As Long = -1, _
    Optional ByVal dataKeyColumn As Long = 0, _
    Optional ByVal jrLabelColumn As Long = 0, _
    Optional ByVal jrSumColumn As Long = 0)

    If dataKeyColumn = 0 Then dataKeyColumn = mod_Construction_OutputLayout.OutputSheetSeiriColumnCore(ws)
    If jrLabelColumn = 0 Then jrLabelColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_JR_PRICE)
    If jrSumColumn = 0 Then jrSumColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_JR_AMOUNT)
    If subconFirstCol = 0 Then subconFirstCol = mod_Construction_OutputLayout.OutputSheetSubconPriceFirstColCore(ws)
    If subconColumnCount < 0 Then
        If vendorNames Is Nothing Then
            subconColumnCount = 0
        Else
            subconColumnCount = vendorNames.Count * 2
        End If
    End If

    Dim lastRow As Long
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws, dataKeyColumn)
    If lastRow < 2 Then Exit Sub

    ClearStaleOutputTotalFormatting ws, lastRow, lastRow + 1, subconFirstCol, _
        subconColumnCount, dataKeyColumn, jrLabelColumn, jrSumColumn

    WriteTotalCells ws, lastRow + 1, jrLabelColumn, "JR���v", jrSumColumn, lastRow

    If Not vendorNames Is Nothing Then
        Dim vendorIndex As Long
        Dim priceColumn As Long
        For vendorIndex = 1 To vendorNames.Count
            priceColumn = subconFirstCol + ((vendorIndex - 1) * 2)
            WriteTotalCells ws, lastRow + 1, _
                            priceColumn, CStr(vendorNames(vendorIndex)) & "���v", _
                            priceColumn + 1, lastRow, useRoundHalfUp:=True
        Next vendorIndex
    End If
End Sub

Public Sub WriteJrTotalRow(ByVal ws As Worksheet)
    Dim emptyVendors As New Collection
    WriteOutputTotalRows ws, emptyVendors, 0, 0
End Sub

Public Sub WritePurchaseNoticeJrTotalRow(ByVal ws As Worksheet)
    Dim emptyVendors As New Collection
    WriteOutputTotalRows ws, emptyVendors, 0, 0, _
        PURCHASE_NOTICE_SEIRI_COL, PURCHASE_NOTICE_JR_PRICE_COL, PURCHASE_NOTICE_JR_AMOUNT_COL
End Sub

Public Sub FillReferenceUnitPrices(ByVal ws As Worksheet, _
                                    ByVal guidanceDocumentName As String)
    Dim lastRow As Long
    lastRow = mod_Construction_LineMapping.GetLastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    Dim lineSheetMap As Object
    Set lineSheetMap = mod_Construction_LineMapping.BuildConstructionLineSheetMap()

    Dim sheetPriceCaches As Object
    Set sheetPriceCaches = CreateObject("Scripting.Dictionary")
    sheetPriceCaches.CompareMode = vbTextCompare

    Dim matchedCount As Long, unresolvedLineCount As Long, missingRecordCount As Long
    Dim pendingCollectCount As Long
    Dim seiriColumn As Long
    Dim dayNightColumn As Long
    Dim qtyColumn As Long
    Dim autoPriceColumn As Long
    Dim autoAmountColumn As Long
    Dim compareColumn As Long
    seiriColumn = mod_Construction_OutputLayout.OutputSheetSeiriColumnCore(ws)
    dayNightColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_DAYNIGHT)
    qtyColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_QTY)
    autoPriceColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_AUTO_PRICE)
    autoAmountColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_AUTO_AMOUNT)
    compareColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_PRICE_COMPARE)

    Dim typeColumn As Long
    Dim unitColumn As Long
    Dim lineColumn As Long
    typeColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_TYPE)
    unitColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_UNIT)
    lineColumn = mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_LINE)

    ' �����ԍ����Ώې���(�P��)�V�[�g�ɖ��o�^�̍s�����V�[�g�ʂɎ��W����
    Dim pendingByLineSheet As Object
    Set pendingByLineSheet = CreateObject("Scripting.Dictionary")
    pendingByLineSheet.CompareMode = vbTextCompare

    Dim isWeldingSheet As Boolean
    Dim weldingPriceSheetName As String
    isWeldingSheet = mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws)
    If isWeldingSheet Then
        weldingPriceSheetName = mod_Construction_OutputLayout.ResolveWeldingPriceSheetName()
        If weldingPriceSheetName = "" Then
            LogCI "���[���n�ڒP��: �V�[�g���̉����Ɏ��s(��{���B6/C6 �܂��� �P���K�p����}�X�^����v)"
        ElseIf Not mod_Construction_OutputLayout.SheetExistsByName(weldingPriceSheetName) Then
            LogCI "���[���n�ڒP��: �V�[�g�u" & weldingPriceSheetName & "�v��������܂���"
        End If
    End If

    Dim r As Long
    For r = 2 To lastRow
        Dim unitPriceSheetName As String
        Dim recordKey As String
        Dim rawLineText As String
        rawLineText = ""
        If isWeldingSheet Then
            unitPriceSheetName = weldingPriceSheetName
            rawLineText = CommonNzText(ws.Cells(r, lineColumn).value)
            recordKey = mod_Construction_OutputLayout.BuildWeldingLineSeiriLookupKey(rawLineText, ws.Cells(r, seiriColumn).value)
        Else
            rawLineText = CommonNzText(ws.Cells(r, lineColumn).value)
            unitPriceSheetName = mod_Construction_LineMapping.ResolveUnitPriceSheetName(lineSheetMap, rawLineText, False)
            recordKey = mod_Construction_LineMapping.NormalizeRecordKey(ws.Cells(r, seiriColumn).value)
        End If

        Dim referencePrice As Variant
        referencePrice = Empty

        Dim recordMissing As Boolean
        recordMissing = False

        If unitPriceSheetName = "" Then
            unresolvedLineCount = unresolvedLineCount + 1
        Else
            Dim priceRows As Object
            Set priceRows = mod_Construction_LineMapping.GetUnitPriceRows(unitPriceSheetName, sheetPriceCaches)

            If priceRows Is Nothing Or recordKey = "" Then
                missingRecordCount = missingRecordCount + 1
                If recordKey <> "" Then recordMissing = True
            Else
                Dim resolvedRecordKey As String
                If isWeldingSheet Then
                    resolvedRecordKey = mod_Construction_LineMapping.FindWeldingPriceRecordKey(priceRows, rawLineText, ws.Cells(r, seiriColumn).value)
                Else
                    resolvedRecordKey = recordKey
                End If

                If resolvedRecordKey <> "" And priceRows.Exists(resolvedRecordKey) Then
                    Dim dayNightPrices As Variant
                    dayNightPrices = priceRows(resolvedRecordKey)
                    referencePrice = mod_Construction_LineMapping.SelectDayNightPrice(CommonNzText(ws.Cells(r, dayNightColumn).value), dayNightPrices)
                    If Not IsEmpty(referencePrice) Then matchedCount = matchedCount + 1
                Else
                    missingRecordCount = missingRecordCount + 1
                    recordMissing = True
                End If
            End If
        End If

        ' �H��(��n��)�V�[�g�̂�: �����ԍ����P���}�X�^�ɖ����s�����V�[�g�֓o�^�ΏۂƂ��Ď��W����B
        ' G���(�O��)�� NormalizeLineLookupText �ŃV�[�g���ƍ����ɏ����ς݁B�ǋL���͍̂s���B
        If recordMissing And (Not isWeldingSheet) And unitPriceSheetName <> "" Then
            CollectMissingSeiriForLineSheet pendingByLineSheet, unitPriceSheetName, recordKey, _
                ws.Cells(r, seiriColumn).value, _
                ws.Cells(r, typeColumn).value, _
                ws.Cells(r, unitColumn).value
            pendingCollectCount = pendingCollectCount + 1
        End If

        If IsEmpty(referencePrice) Or IsError(referencePrice) Then
            ws.Cells(r, autoPriceColumn).ClearContents
        Else
            ws.Cells(r, autoPriceColumn).value = referencePrice
        End If
        WritePriceComparison ws, r, unitPriceSheetName, True, guidanceDocumentName
    Next r

    ' �����ԍ����P���}�X�^�ɖ������������A�Ώې���V�[�g�̍ŉ����֓o�^����
    If Not isWeldingSheet Then
        If pendingCollectCount > 0 Then
            LogCI "�����ԍ����o�^�̒ǋL���=" & pendingCollectCount & _
                  " (����V�[�g��=" & pendingByLineSheet.Count & ")"
        End If
        RegisterMissingSeiriToLineSheets pendingByLineSheet
    End If

    ws.Range(ws.Cells(2, autoAmountColumn), ws.Cells(lastRow, autoAmountColumn)).FormulaR1C1 = _
        "=IF(OR(RC[" & (autoPriceColumn - autoAmountColumn) & "]="""",RC[" & (qtyColumn - autoAmountColumn) & "]=""""),"""",RC[" & (autoPriceColumn - autoAmountColumn) & "]*RC[" & (qtyColumn - autoAmountColumn) & "])"

    SafeSetRangeNumberFormatLocal ws.Range(ws.Cells(2, autoPriceColumn), ws.Cells(lastRow, autoAmountColumn)), ConstructionIntegerNumberFormat()
    ws.Range(ws.Cells(1, compareColumn), ws.Cells(lastRow, compareColumn)).HorizontalAlignment = xlCenter

    LogCI "�Q�ƒP����v=" & matchedCount & _
          " / ���斢����=" & unresolvedLineCount & _
          " / �����ԍ�����v=" & missingRecordCount
End Sub

Public Sub RefreshConstructionReferenceUnitPricesOnExistingSheetsCore()
    Dim g As New clsPerfGuard
    mod_Construction_LineMapping.ClearProjectLineNameAliasCache

    On Error GoTo Cleanup
    g.Suspend DisableEvents:=False

    Dim ws As Worksheet
    Dim refreshedCount As Long
    For Each ws In ThisWorkbook.Worksheets
        If IsConstructionDocumentOutputSheet(ws) Then
            Dim guidanceDocumentName As String
            guidanceDocumentName = mod_Construction_LineMapping.ResolveGuidanceDocumentNameFromOutputSheet(ws)
            FillReferenceUnitPrices ws, guidanceDocumentName
            If mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then
                ApplyPriceGuidanceColumnLayoutAtColumns _
                    ws, mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_PRICE_COMPARE), mod_Construction_OutputLayout.OutputSheetColCore(ws, COL_PRICE_GUIDANCE)
            Else
                ApplyPriceGuidanceColumnLayout ws
            End If
            refreshedCount = refreshedCount + 1
        End If
    Next ws

    LogCI "�����{�H�w�������V�[�g�̎Q�ƒP���ēǍ�: �Ώ�=" & refreshedCount

Cleanup:
    g.Restore
End Sub

Public Sub RefreshConstructionReferencePricesForUnitPriceChangeCore( _
    ByVal wsUnitPrice As Worksheet, ByVal changedRange As Range)

    If wsUnitPrice Is Nothing Or changedRange Is Nothing Then Exit Sub
    If Not mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) Then Exit Sub

    Dim monitored As Range
    Set monitored = Application.Intersect(changedRange, _
        Application.Union(wsUnitPrice.Columns(COL_SEIRI), _
                          wsUnitPrice.Columns(5), _
                          wsUnitPrice.Columns(6)))
    If monitored Is Nothing Then Exit Sub

    Dim changedPriceRows As Object
    Set changedPriceRows = CreateObject("Scripting.Dictionary")
    changedPriceRows.CompareMode = vbTextCompare

    Dim changedCell As Range
    Dim recordKey As String
    For Each changedCell In monitored.Cells
        If changedCell.Row >= UNIT_PRICE_DATA_START_ROW Then
            recordKey = mod_Construction_LineMapping.NormalizeRecordKey(wsUnitPrice.Cells(changedCell.Row, COL_SEIRI).value)
            If recordKey <> "" Then
                changedPriceRows(recordKey) = Array( _
                    wsUnitPrice.Cells(changedCell.Row, 5).value, _
                    wsUnitPrice.Cells(changedCell.Row, 6).value)
            End If
        End If
    Next changedCell

    If changedPriceRows.Count = 0 Then Exit Sub

    Dim lineSheetMap As Object
    Set lineSheetMap = mod_Construction_LineMapping.BuildConstructionLineSheetMap()

    Dim wsTarget As Worksheet
    For Each wsTarget In ThisWorkbook.worksheets
        If IsConstructionDocumentOutputSheet(wsTarget) Then
            RefreshConstructionReferencePricesOnSheet _
                wsTarget, wsUnitPrice.Name, changedPriceRows, lineSheetMap
        End If
    Next wsTarget
End Sub

Public Function IsConstructionDocumentOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If Not mod_Construction_OutputLayout.IsConstructionVendorOutputSheet(ws) Then Exit Function
    If mod_Construction_BasicTotals.FindHeaderColumn(ws, "�����ԍ�") = 0 Then Exit Function
    IsConstructionDocumentOutputSheet = (mod_Construction_BasicTotals.FindHeaderColumn(ws, "�P����r") > 0)
End Function

Public Sub RefreshConstructionReferencePricesOnSheet( _
    ByVal ws As Worksheet, _
    ByVal unitPriceSheetName As String, _
    ByVal changedPriceRows As Object, _
    ByVal lineSheetMap As Object)

    ' �n�ڃV�[�g�̎Q�ƒP���̓��[���n�ڒP���V�[�g�R���̂��߁A�H���P���V�[�g�ύX�ł͍Čv�Z���Ȃ��B
    If mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then Exit Sub

    Dim seiriColumn As Long
    Dim dayNightColumn As Long
    Dim quantityColumn As Long
    Dim jrPriceColumn As Long
    Dim comparisonColumn As Long
    seiriColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "�����ԍ�")
    dayNightColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "�����")
    quantityColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "����")
    jrPriceColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "JR�P��")
    comparisonColumn = mod_Construction_BasicTotals.FindHeaderColumn(ws, "�P����r")

    If seiriColumn = 0 Or dayNightColumn = 0 Or quantityColumn = 0 Then Exit Sub
    If jrPriceColumn = 0 Or comparisonColumn < 3 Then Exit Sub

    Dim autoPriceColumn As Long
    Dim autoAmountColumn As Long
    Dim guidanceColumn As Long
    autoPriceColumn = comparisonColumn - 2
    autoAmountColumn = comparisonColumn - 1
    guidanceColumn = comparisonColumn + 1

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, seiriColumn).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim isWeldingSheet As Boolean
    isWeldingSheet = mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws)

    Dim normalizedSourceSheet As String
    normalizedSourceSheet = mod_Construction_LineMapping.NormalizeLineLookupText(unitPriceSheetName, isWeldingSheet)

    Dim r As Long
    Dim recordKey As String
    Dim resolvedSheetName As String
    Dim referencePrice As Variant
    For r = 2 To lastRow
        recordKey = mod_Construction_LineMapping.NormalizeRecordKey(ws.Cells(r, seiriColumn).value)
        If recordKey <> "" And changedPriceRows.Exists(recordKey) Then
            resolvedSheetName = mod_Construction_LineMapping.ResolveUnitPriceSheetName( _
                lineSheetMap, CommonNzText(ws.Cells(r, mod_Construction_BasicTotals.FindHeaderColumn(ws, "�_����於")).value), _
                isWeldingSheet)

            If mod_Construction_LineMapping.NormalizeLineLookupText(resolvedSheetName, isWeldingSheet) = normalizedSourceSheet Then
                referencePrice = mod_Construction_LineMapping.SelectDayNightPrice( _
                    CommonNzText(ws.Cells(r, dayNightColumn).value), _
                    changedPriceRows(recordKey))

                If IsEmpty(referencePrice) Or IsError(referencePrice) Then
                    ws.Cells(r, autoPriceColumn).ClearContents
                Else
                    ws.Cells(r, autoPriceColumn).value = referencePrice
                End If

                ws.Cells(r, autoAmountColumn).FormulaR1C1 = _
                    "=IF(OR(RC[" & (autoPriceColumn - autoAmountColumn) & "]=""""," & _
                    "RC[" & (quantityColumn - autoAmountColumn) & "]=""""),""""," & _
                    "RC[" & (autoPriceColumn - autoAmountColumn) & "]*" & _
                    "RC[" & (quantityColumn - autoAmountColumn) & "])"

                WritePriceComparisonAtColumns _
                    ws, r, unitPriceSheetName, autoPriceColumn, jrPriceColumn, _
                    comparisonColumn, guidanceColumn, True
            End If
        End If
    Next r

    SafeSetRangeNumberFormatLocal ws.Range(ws.Cells(2, autoPriceColumn), _
             ws.Cells(lastRow, autoAmountColumn)), ConstructionIntegerNumberFormat()
    ws.Range(ws.Cells(1, comparisonColumn), _
             ws.Cells(lastRow, comparisonColumn)).HorizontalAlignment = xlCenter
    ApplyPriceGuidanceColumnLayoutAtColumns ws, comparisonColumn, guidanceColumn
End Sub

