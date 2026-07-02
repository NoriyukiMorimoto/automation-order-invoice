Option Explicit

Private Const BASIC_INFO_VENDOR_NAME_ROW As Long = 11
Private Const BASIC_INFO_VENDOR_BLOCK_TOP_ROW As Long = 10
Private Const BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW As Long = 29
Private Const BASIC_INFO_VENDOR_RAIL_PATTERN_ROW As Long = 30
Private Const BASIC_INFO_VENDOR_WELDING_RATIO_ROW As Long = 31
Private Const BASIC_INFO_YEAR_CELL As String = "B4"
Private Const BASIC_INFO_BILLING_COUNT_CELL As String = "F4"
Private Const BASIC_INFO_PRICE_KIND_CELL As String = "C22"
Private Const VENDOR_UNIT_PRICE_JR_HEADER_ROW As Long = 4
Private Const VENDOR_UNIT_PRICE_JR_HEADER_COL_START As Long = 5
Private Const VENDOR_UNIT_PRICE_JR_HEADER_COL_END As Long = 6
Private Const VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW As Long = 1
Private Const VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_FONT_SIZE As Long = 11
Private Const VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_NUMBER_FORMAT As String = "0.0%"
Private Const VENDOR_UNIT_PRICE_HEADER_ROW As Long = 4
Private Const VENDOR_UNIT_PRICE_NAME_ROW As Long = 5
Private Const VENDOR_UNIT_PRICE_LABEL_ROW As Long = 6
Private Const VENDOR_UNIT_PRICE_DATA_START_ROW As Long = 7
Private Const VENDOR_UNIT_PRICE_FIRST_DAY_COL As Long = 7
Private Const VENDOR_UNIT_PRICE_REF_UNIT_COL As Long = 5
Private Const VENDOR_UNIT_PRICE_REF_WIDTH_COL As Long = 6
Private Const VENDOR_UNIT_PRICE_WORK_TYPE_COL As Long = 3
Private Const VENDOR_UNIT_PRICE_LAST_ROW_COL As Long = 2
Private Const VENDOR_UNIT_PRICE_INITIAL_FILL_LAST_COL As Long = 10
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_R As Long = 128
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_G As Long = 128
Private Const VENDOR_UNIT_PRICE_FILL_COLOR_B As Long = 128
Private Const VENDOR_UNIT_PRICE_NUMBER_FORMAT As String = "#,##0"
Private Const MAX_VENDOR_BLOCK_COUNT As Long = 10
Private Const VENDOR_ROW_NAME_INDEX As Long = 0
Private Const VENDOR_ROW_UNIT_PRICE_NAME_INDEX As Long = 12

Public Function BuildVendorUnitPriceFormula(ByVal wsUnitPrice As Worksheet, _
                                             ByVal rowIndex As Long, _
                                             ByVal sourceCol As Long, _
                                             ByVal ratioAddress As String) As String
    Dim unitCellRef As String
    unitCellRef = wsUnitPrice.Cells(rowIndex, sourceCol).Address(False, False)

    BuildVendorUnitPriceFormula = "=ROUND(" & unitCellRef & "*(" & ratioAddress & ")," & _
                                  "-INT(LOG10(" & unitCellRef & "*(" & ratioAddress & ")))+2)"
End Function

Public Function BuildVendorUnitPriceFormulaR1C1(ByVal isDayColumn As Boolean, _
                                                 ByVal targetCol As Long, _
                                                 ByVal ratioAddress As String) As String
    Dim sourceCol As Long
    If isDayColumn Then
        sourceCol = VENDOR_UNIT_PRICE_REF_UNIT_COL
    Else
        sourceCol = VENDOR_UNIT_PRICE_REF_WIDTH_COL
    End If

    Dim sourceOffset As Long
    sourceOffset = sourceCol - targetCol

    BuildVendorUnitPriceFormulaR1C1 = "=ROUND(RC[" & sourceOffset & "]*(" & ratioAddress & ")," & _
        "-INT(LOG10(RC[" & sourceOffset & "]*(" & ratioAddress & ")))+2)"
End Function

Public Function BuildVendorUnitPriceHeaderText(ByVal wsInfo As Worksheet) As String
    BuildVendorUnitPriceHeaderText = Trim$(CStr(wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value)) & _
                                     CommonExtractYear4Digits(CStr(wsInfo.Range(BASIC_INFO_YEAR_CELL).value)) & _
                                     VendorUnitPriceOutsourceLabelText()
End Function

Public Function BuildVendorUnitPriceJrHeaderText(ByVal wsInfo As Worksheet) As String
    BuildVendorUnitPriceJrHeaderText = Trim$(CStr(wsInfo.Range(BASIC_INFO_YEAR_CELL).value)) & _
                                       VendorJrUnitPriceLabelText() & _
                                       Trim$(CStr(wsInfo.Range(BASIC_INFO_PRICE_KIND_CELL).value))
End Function

Public Function ChangedRangeIncludesVendorWorkTypeRow(ByVal wsInfo As Worksheet, _
                                                       ByVal changedRange As Range) As Boolean
    Dim vendorCount As Long
    Dim i As Long

    If wsInfo Is Nothing Then Exit Function
    If changedRange Is Nothing Then Exit Function

    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)
    For i = 1 To vendorCount
        If Not Intersect(changedRange, wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, mod_VendorBlockLayout.VendorValueColumnByIndex(i))) Is Nothing Then
            ChangedRangeIncludesVendorWorkTypeRow = True
            Exit Function
        End If
    Next i
End Function

Public Function CollectMonitorChangedValueColumns(ByVal wsInfo As Worksheet, _
                                                   ByVal changedRange As Range, _
                                                   ByVal includeOutsourceRatioRow As Boolean, _
                                                   ByVal includeWeldingRatioRow As Boolean) As Collection
    Dim result As Collection
    Dim vendorCount As Long
    Dim i As Long
    Set result = New Collection

    If wsInfo Is Nothing Then
        Set CollectMonitorChangedValueColumns = result
        Exit Function
    End If
    If changedRange Is Nothing Then
        Set CollectMonitorChangedValueColumns = result
        Exit Function
    End If

    If Not Intersect(changedRange, wsInfo.Range(BASIC_INFO_YEAR_CELL & "," & BASIC_INFO_BILLING_COUNT_CELL)) Is Nothing Then
        vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)
        For i = 1 To vendorCount
            AddUniqueLongToCollection result, mod_VendorBlockLayout.VendorValueColumnByIndex(i)
        Next i
        Set CollectMonitorChangedValueColumns = result
        Exit Function
    End If

    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)
    For i = 1 To vendorCount
        Dim valueCol As Long
        Dim monitorCell As Range
        valueCol = mod_VendorBlockLayout.VendorValueColumnByIndex(i)

        Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueCol)
        If Not Intersect(changedRange, monitorCell) Is Nothing Then
            AddUniqueLongToCollection result, valueCol
            GoTo ContinueNextVendorColumn
        End If

        Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol)
        If Not Intersect(changedRange, monitorCell) Is Nothing Then
            AddUniqueLongToCollection result, valueCol
            GoTo ContinueNextVendorColumn
        End If

        If includeOutsourceRatioRow Then
            Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueCol)
            If Not Intersect(changedRange, monitorCell) Is Nothing Then
                AddUniqueLongToCollection result, valueCol
                GoTo ContinueNextVendorColumn
            End If
        End If

        If includeWeldingRatioRow Then
            Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_WELDING_RATIO_ROW, valueCol)
            If Not Intersect(changedRange, monitorCell) Is Nothing Then
                AddUniqueLongToCollection result, valueCol
                GoTo ContinueNextVendorColumn
            End If

            Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_RAIL_PATTERN_ROW, valueCol)
            If Not Intersect(changedRange, monitorCell) Is Nothing Then
                AddUniqueLongToCollection result, valueCol
            End If
        End If

ContinueNextVendorColumn:
    Next i

    Set CollectMonitorChangedValueColumns = result
End Function

Public Function CollectOutsourceRatioOnlyChangedValueColumns(ByVal wsInfo As Worksheet, _
                                                              ByVal changedRange As Range, _
                                                              ByVal excludeColumns As Collection) As Collection
    Dim result As Collection
    Set result = New Collection

    If wsInfo Is Nothing Then
        Set CollectOutsourceRatioOnlyChangedValueColumns = result
        Exit Function
    End If
    If changedRange Is Nothing Then
        Set CollectOutsourceRatioOnlyChangedValueColumns = result
        Exit Function
    End If

    Dim vendorCount As Long
    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)

    Dim i As Long
    For i = 1 To vendorCount
        Dim valueCol As Long
        valueCol = mod_VendorBlockLayout.VendorValueColumnByIndex(i)

        If Not IsLongInCollection(excludeColumns, valueCol) Then
            Dim monitorCell As Range
            Set monitorCell = wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueCol)
            If Not Intersect(changedRange, monitorCell) Is Nothing Then
                AddUniqueLongToCollection result, valueCol
            End If
        End If
    Next i

    Set CollectOutsourceRatioOnlyChangedValueColumns = result
End Function

Public Function GetPreferredWeldingRatioColumnFromChange(ByVal wsInfo As Worksheet, _
                                                          ByVal changedRange As Range) As Long
    Dim vendorCount As Long
    Dim i As Long
    Dim ratioCell As Range

    GetPreferredWeldingRatioColumnFromChange = 0
    If wsInfo Is Nothing Then Exit Function
    If changedRange Is Nothing Then Exit Function

    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)
    For i = 1 To vendorCount
        Set ratioCell = wsInfo.Cells(BASIC_INFO_VENDOR_WELDING_RATIO_ROW, mod_VendorBlockLayout.VendorValueColumnByIndex(i))
        If Not Intersect(changedRange, ratioCell) Is Nothing Then
            GetPreferredWeldingRatioColumnFromChange = ratioCell.Column
            Exit Function
        End If
    Next i
End Function

Public Function GetVendorIndexFromValueColumn(ByVal valueColumn As Long) As Long
    Dim i As Long
    For i = 1 To MAX_VENDOR_BLOCK_COUNT
        If mod_VendorBlockLayout.VendorValueColumnByIndex(i) = valueColumn Then
            GetVendorIndexFromValueColumn = i
            Exit Function
        End If
    Next i
End Function

Public Function GetVendorOutsourceRatioAddress(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As String
    GetVendorOutsourceRatioAddress = "'" & Replace$(wsInfo.Name, "'", "''") & "'!" & _
                                     wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueColumn).Address(True, True)
End Function

Public Function GetVendorOutsourceRatioNumericValue(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Variant
    Dim ratioCell As Range
    Set ratioCell = wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueColumn)

    Dim cellValue As Variant
    cellValue = ratioCell.Value2
    If IsNumeric(cellValue) Then
        GetVendorOutsourceRatioNumericValue = CDbl(cellValue)
        Exit Function
    End If

    Dim textValue As String
    textValue = Trim$(CStr(ratioCell.value))
    If Len(textValue) = 0 Then Exit Function

    If Right$(textValue, 1) = ChrW$(&HFF05) Then
        textValue = Trim$(Left$(textValue, Len(textValue) - 1))
        If IsNumeric(textValue) Then GetVendorOutsourceRatioNumericValue = CDbl(textValue) / 100#
        Exit Function
    End If

    If IsNumeric(textValue) Then GetVendorOutsourceRatioNumericValue = CDbl(textValue)
End Function

Public Function GetVendorOutsourceRatioPercentValue(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Variant
    Dim ratioValue As Variant
    ratioValue = GetVendorOutsourceRatioNumericValue(wsInfo, valueColumn)
    If Not IsNumeric(ratioValue) Then Exit Function

    Dim normalizedValue As Double
    normalizedValue = CDbl(ratioValue)
    If normalizedValue > 1# And normalizedValue <= 100# Then normalizedValue = normalizedValue / 100#
    GetVendorOutsourceRatioPercentValue = normalizedValue
End Function

Public Function GetVendorUnitPriceInitialFillLastColumn(ByVal wsInfo As Worksheet) As Long
    Dim lastFillCol As Long
    lastFillCol = VENDOR_UNIT_PRICE_INITIAL_FILL_LAST_COL

    Dim vendorCount As Long
    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = mod_VendorBlockLayout.VendorValueColumnByIndex(vendorIndex)
        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            Dim nightCol As Long
            nightCol = VendorUnitPriceDayColumnByValueColumn(valueColumn) + 1
            If nightCol > lastFillCol Then lastFillCol = nightCol
        End If
    Next vendorIndex

    GetVendorUnitPriceInitialFillLastColumn = lastFillCol
End Function

Public Function GetVendorUnitPriceLastDataRow(ByVal wsUnitPrice As Worksheet) As Long
    GetVendorUnitPriceLastDataRow = wsUnitPrice.Cells(wsUnitPrice.rows.Count, VENDOR_UNIT_PRICE_LAST_ROW_COL).End(xlUp).Row
End Function

Public Function HasNumericVendorUnitPriceSource(ByVal sourceCell As Range) As Boolean
    If sourceCell Is Nothing Then Exit Function
    If IsError(sourceCell.Value2) Then Exit Function
    If Len(Trim$(CStr(sourceCell.Value2))) = 0 Then Exit Function
    HasNumericVendorUnitPriceSource = IsNumeric(sourceCell.Value2)
End Function

Public Function HasVendorName(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Boolean
    HasVendorName = (Len(Trim$(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value))) > 0)
End Function

Public Function HasVendorOutsourceRatio(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Boolean
    HasVendorOutsourceRatio = (Len(Trim$(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueColumn).value))) > 0)
End Function

Public Function IsBlankSourceValue(ByVal srcValue As Variant) As Boolean
    If IsError(srcValue) Then Exit Function
    IsBlankSourceValue = (Len(Trim$(CStr(CommonNzText(srcValue)))) = 0)
End Function

' E列/F列のみ: 空欄グレー塗りと数値桁区切りを配列読み取りで一括適用(業者列は触らない)。

Public Function IsLongInCollection(ByVal target As Collection, ByVal value As Long) As Boolean
    If target Is Nothing Then Exit Function

    Dim existing As Variant
    For Each existing In target
        If CLng(existing) = value Then
            IsLongInCollection = True
            Exit Function
        End If
    Next existing
End Function

' 外注比率(29行目)のみが変わった業者列を集める。工事種別/業者名が変わった列(excludeColumns)は
' 全展開側(RefreshVendorUnitPriceForValueColumn)が比率表示も含めて処理するため、ここでは除外する。

Public Function IsNumericSourceValue(ByVal srcValue As Variant) As Boolean
    If IsError(srcValue) Then Exit Function
    If Len(Trim$(CStr(CommonNzText(srcValue)))) = 0 Then Exit Function
    IsNumericSourceValue = IsNumeric(srcValue)
End Function

' IsVendorUnitPriceSourceCellBlank(セル版)と同一判定を、読み取り済み値に対して行う。

Public Function IsRailConstructionVendorBlock(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Boolean
    IsRailConstructionVendorBlock = _
        (StrComp(CommonNormalizeText(CStr(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueColumn).value)), _
                 VendorRailConstructionText(), vbTextCompare) = 0)
End Function

Public Function IsVendorUnitPriceBlockAlreadyBuilt(ByVal wsUnitPrice As Worksheet, _
                                                    ByVal dayCol As Long, _
                                                    ByVal nightCol As Long) As Boolean
    On Error GoTo ExitHandler

    Dim headerCell As Range
    Set headerCell = wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol)

    IsVendorUnitPriceBlockAlreadyBuilt = _
        headerCell.MergeCells And _
        (Len(Trim$(CStr(headerCell.MergeArea.Cells(1, 1).value))) > 0)
    Exit Function

ExitHandler:
    IsVendorUnitPriceBlockAlreadyBuilt = False
End Function

Public Function IsVendorUnitPriceSourceBlank(ByVal wsUnitPrice As Worksheet, _
                                              ByVal rowIndex As Long, _
                                              ByVal sourceCol As Long) As Boolean
    IsVendorUnitPriceSourceBlank = _
        (Len(Trim$(CStr(wsUnitPrice.Cells(rowIndex, sourceCol).value))) = 0)
End Function

Public Function IsVendorUnitPriceSourceCellBlank(ByVal sourceCell As Range) As Boolean
    If sourceCell Is Nothing Then
        IsVendorUnitPriceSourceCellBlank = True
        Exit Function
    End If
    If IsError(sourceCell.Value2) Then Exit Function
    IsVendorUnitPriceSourceCellBlank = (Len(Trim$(CStr(sourceCell.Value2))) = 0)
End Function

Public Function ResolveVendorUnitPriceName(ByVal vendorUnitPriceNameMap As Object, _
                                            ByVal basicInfoVendorName As String) As String
    Dim vendorNameKey As String
    vendorNameKey = CommonNormalizeText(basicInfoVendorName)

    If Not vendorUnitPriceNameMap Is Nothing Then
        If vendorUnitPriceNameMap.Exists(vendorNameKey) Then
            ResolveVendorUnitPriceName = CStr(vendorUnitPriceNameMap(vendorNameKey))
            Exit Function
        End If
    End If

    ResolveVendorUnitPriceName = basicInfoVendorName
    If vendorNameKey <> "" Then
        mod_DebugLog.Log "[VendorMaster] Unit price vendor name not found: " & basicInfoVendorName
    End If
End Function

Public Function RowNeedsFormulaFromArrays(ByVal srcArr As Variant, _
                                           ByVal workArr As Variant, _
                                           ByVal r As Long, _
                                           ByVal wasteKeyword As String) As Boolean
    If Len(Trim$(CStr(CommonNzText(srcArr(r, 1))))) = 0 Then Exit Function

    Dim workTypeName As String
    workTypeName = CommonNormalizeText(CStr(CommonNzText(workArr(r, 1))))
    If workTypeName <> "" Then
        If InStr(1, workTypeName, wasteKeyword, vbTextCompare) > 0 Then Exit Function
    End If

    RowNeedsFormulaFromArrays = True
End Function

' 集約したグレー塗り範囲へ、ApplyVendorUnitPriceGreyFill と同一の書式をまとめて適用する。

Public Function ShouldApplyVendorUnitPriceBlock(ByVal wsInfo As Worksheet, ByVal valueColumn As Long) As Boolean
    ShouldApplyVendorUnitPriceBlock = _
        IsRailConstructionVendorBlock(wsInfo, valueColumn) And _
        HasVendorName(wsInfo, valueColumn) And _
        HasVendorOutsourceRatio(wsInfo, valueColumn)
End Function

Public Function VendorJrUnitPriceLabelText() As String
    VendorJrUnitPriceLabelText = "JR"
End Function

Public Function VendorRailConstructionText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H5DE5) & ChrW$(&H4E8B)
    End If
    VendorRailConstructionText = cached
End Function

Public Function VendorUnitPriceDayColumnByValueColumn(ByVal valueColumn As Long) As Long
    Dim vendorIndex As Long
    vendorIndex = GetVendorIndexFromValueColumn(valueColumn)
    If vendorIndex < 1 Then vendorIndex = 1
    VendorUnitPriceDayColumnByValueColumn = VENDOR_UNIT_PRICE_FIRST_DAY_COL + ((vendorIndex - 1) * 2)
End Function

Public Function VendorUnitPriceDayLabelText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H663C) & ChrW$(&H9593)
    VendorUnitPriceDayLabelText = cached
End Function

Public Function VendorUnitPriceFontNameText() As String
    Static cached As String
    If cached = "" Then
        cached = "BIZ UD" & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
    End If
    VendorUnitPriceFontNameText = cached
End Function

Public Function VendorUnitPriceNightLabelText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H591C) & ChrW$(&H9593)
    VendorUnitPriceNightLabelText = cached
End Function

Public Function VendorUnitPriceOutsourceLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H5358) & ChrW$(&H4FA1)
    End If
    VendorUnitPriceOutsourceLabelText = cached
End Function

Public Function VendorUnitPriceOutsourceRatioLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H6BD4) & ChrW$(&H7387) & ChrW$(&HFF1D)
    End If
    VendorUnitPriceOutsourceRatioLabelText = cached
End Function

Public Function VendorUnitPriceRowNeedsFormulaForSource(ByVal wsUnitPrice As Worksheet, _
                                                         ByVal rowIndex As Long, _
                                                         ByVal sourceCol As Long, _
                                                         ByVal wasteKeyword As String) As Boolean
    If IsVendorUnitPriceSourceBlank(wsUnitPrice, rowIndex, sourceCol) Then Exit Function

    Dim workTypeName As String
    workTypeName = CommonNormalizeText(CStr(wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_WORK_TYPE_COL).value))
    If workTypeName <> "" Then
        If InStr(1, workTypeName, wasteKeyword, vbTextCompare) > 0 Then Exit Function
    End If

    VendorUnitPriceRowNeedsFormulaForSource = True
End Function

Public Function VendorWasteDisposalKeywordText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H7523) & ChrW$(&H5EC3) & ChrW$(&H51E6) & ChrW$(&H7406)
    VendorWasteDisposalKeywordText = cached
End Function

Public Sub AddUniqueLongToCollection(ByVal target As Collection, ByVal value As Long)
    Dim existing As Variant
    For Each existing In target
        If CLng(existing) = value Then Exit Sub
    Next existing
    target.Add value
End Sub

Public Sub ApplyConstructionUnitPriceImportedRowDecorationsFast(ByVal wsUnitPrice As Worksheet, _
                                                                ByVal wsInfo As Worksheet, _
                                                                ByVal firstRow As Long, _
                                                                ByVal lastRow As Long, _
                                                                ByVal skipVendorColumnRefresh As Boolean)
    Dim rowCount As Long
    rowCount = lastRow - firstRow + 1
    If rowCount <= 0 Then Exit Sub

    ' B列(最終行判定列)を一括読み取り
    Dim bArr As Variant
    ReadVendorUnitPriceColumnValues wsUnitPrice, firstRow, lastRow, VENDOR_UNIT_PRICE_LAST_ROW_COL, bArr

    ' 「B列に値がある行」の連続セグメントごとに、罫線(A～F)と塗り(E～F)を範囲一括適用。
    ' 業者列(G列以降)の塗り/数式は ApplyVendorUnitPriceBlockToSheet 側で設定済みのため、
    ' E～lastFillCol へ塗ると業者列の書式を上書きしてしまう(E～F のみに限定する)。
    Dim segStart As Long
    segStart = 0

    Dim r As Long
    For r = 1 To rowCount + 1
        Dim rowIndex As Long
        rowIndex = firstRow + r - 1

        Dim hasValue As Boolean
        hasValue = False
        If r <= rowCount Then
            hasValue = (Len(Trim$(CStr(CommonNzText(bArr(r, 1))))) > 0)
        End If

        If hasValue Then
            If segStart = 0 Then segStart = rowIndex
        Else
            If segStart > 0 Then
                ' 罫線: A列～F列
                With wsUnitPrice.Range(wsUnitPrice.Cells(segStart, 1), _
                                       wsUnitPrice.Cells(rowIndex - 1, VENDOR_UNIT_PRICE_REF_WIDTH_COL)).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                ' 塗り: E列～F列(JR単価元)
                With wsUnitPrice.Range(wsUnitPrice.Cells(segStart, VENDOR_UNIT_PRICE_REF_UNIT_COL), _
                                       wsUnitPrice.Cells(rowIndex - 1, VENDOR_UNIT_PRICE_REF_WIDTH_COL)).Interior
                    .Color = RGB(VENDOR_UNIT_PRICE_FILL_COLOR_R, _
                                 VENDOR_UNIT_PRICE_FILL_COLOR_G, _
                                 VENDOR_UNIT_PRICE_FILL_COLOR_B)
                End With
                segStart = 0
            End If
        End If
    Next r

    If skipVendorColumnRefresh Then
        ApplyVendorUnitPriceSourceEfDecorationsFast wsUnitPrice, firstRow, lastRow, bArr
        RefreshVendorUnitPriceBordersForSheet wsUnitPrice, wsInfo
        Exit Sub
    End If

    ' シート全体の外枠罫線再適用(従来 ApplyVendorUnitPriceBaseRowBorders 内で呼んでいた処理)
    RefreshVendorUnitPriceBordersForSheet wsUnitPrice, wsInfo

    ' 単価元セル(E/F)の書式(空欄グレー塗り/数値桁区切り)と業者列の再計算
    ApplyVendorUnitPriceSourceRowsForRangeFast wsUnitPrice, wsInfo, firstRow, lastRow, bArr

    ' 単価元列(E/F)の桁区切り書式をまとめて適用
    ApplyVendorUnitPriceSourceColumnsNumberFormatFast wsUnitPrice, firstRow, lastRow, bArr
End Sub

' 工事単価シートのデータ行(7行目以降)へ罫線・塗りつぶし・桁区切りを一括適用する。

Public Sub ApplyVendorUnitPriceAddedBlocksToSheetSafely(ByVal wsUnitPrice As Worksheet, _
                                                         ByVal wsInfo As Worksheet, _
                                                         ByVal previousCount As Long, _
                                                         ByVal vendorCount As Long, _
                                                         ByVal vendorUnitPriceNameMap As Object)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim valueColumn As Long
    Dim dayCol As Long
    Dim nightCol As Long
    For i = previousCount + 1 To vendorCount
        valueColumn = mod_VendorBlockLayout.VendorValueColumnByIndex(i)
        dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
        nightCol = dayCol + 1
        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn, vendorUnitPriceNameMap
        Else
            ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
        End If
    Next i
    Exit Sub

ErrorHandler:
    mod_DebugLog.Log "[VendorMaster] ApplyVendorUnitPriceAddedBlocksToSheetSafely failed sheet=[" & _
                     wsUnitPrice.Name & "] Err " & Err.Number & ": " & Err.Description
End Sub

' F9減少時: 不要になった業者ブロックを1シートからクリアする。エラーはログして続行。

Public Sub ApplyVendorUnitPriceBaseRowBorders(ByVal wsUnitPrice As Worksheet, _
                                                ByVal wsInfo As Worksheet, _
                                                ByVal changedCells As Range)
    Dim changedCell As Range
    For Each changedCell In changedCells.Cells
        If changedCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) > 0 Then
                With wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, 1), _
                                            wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_REF_WIDTH_COL)).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            End If
        End If
    Next changedCell

    RefreshVendorUnitPriceBordersForSheet wsUnitPrice, wsInfo
End Sub

Public Sub ApplyVendorUnitPriceBlockToSheet(ByVal wsUnitPrice As Worksheet, _
                                              ByVal wsInfo As Worksheet, _
                                              ByVal valueColumn As Long, _
                                              ByVal vendorUnitPriceNameMap As Object, _
                                              Optional ByVal cachedLastRow As Long = 0, _
                                              Optional ByVal usePreloadedArrays As Boolean = False, _
                                              Optional ByRef preloadedDaySrcArr As Variant, _
                                              Optional ByRef preloadedNightSrcArr As Variant, _
                                              Optional ByRef preloadedWorkArr As Variant)
    Dim dayCol As Long
    Dim nightCol As Long
    Dim lastRow As Long
    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
    nightCol = dayCol + 1
    If cachedLastRow >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
        lastRow = cachedLastRow
    Else
        lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    End If

    ApplyVendorUnitPriceColumnWidths wsUnitPrice, dayCol, nightCol
    ApplyVendorUnitPriceOutsourceRatioRow wsUnitPrice, wsInfo, valueColumn, dayCol, nightCol
    ApplyVendorUnitPriceMergedHeader wsUnitPrice, dayCol, nightCol, BuildVendorUnitPriceHeaderText(wsInfo)
    ApplyVendorUnitPriceMergedVendorName wsUnitPrice, dayCol, nightCol, _
        ResolveVendorUnitPriceName(vendorUnitPriceNameMap, _
                                   CStr(wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value))

    With wsUnitPrice
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, dayCol).value = VendorUnitPriceDayLabelText()
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, nightCol).value = VendorUnitPriceNightLabelText()
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, dayCol).HorizontalAlignment = xlCenter
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, dayCol).VerticalAlignment = xlCenter
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, nightCol).HorizontalAlignment = xlCenter
        .Cells(VENDOR_UNIT_PRICE_LABEL_ROW, nightCol).VerticalAlignment = xlCenter
    End With

    ApplyVendorUnitPriceDataRows wsUnitPrice, wsInfo, valueColumn, dayCol, nightCol, lastRow, _
        usePreloadedArrays, preloadedDaySrcArr, preloadedNightSrcArr, preloadedWorkArr
    ApplyVendorUnitPriceFont wsUnitPrice, dayCol, nightCol, lastRow
    ApplyVendorUnitPriceBorders wsUnitPrice, dayCol, nightCol, lastRow
End Sub

Public Sub ApplyVendorUnitPriceBorders(ByVal wsUnitPrice As Worksheet, _
                                        ByVal dayCol As Long, _
                                        ByVal nightCol As Long, _
                                        Optional ByVal cachedLastRow As Long = 0)
    Dim lastRow As Long
    If cachedLastRow >= VENDOR_UNIT_PRICE_HEADER_ROW Then
        lastRow = cachedLastRow
    Else
        lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    End If
    If lastRow < VENDOR_UNIT_PRICE_HEADER_ROW Then Exit Sub

    Dim borderRange As Range
    Set borderRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol), _
                                        wsUnitPrice.Cells(lastRow, nightCol))

    With borderRange
        .Borders.LineStyle = xlNone
        With .Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

Public Sub ApplyVendorUnitPriceCell(ByVal targetCell As Range, _
                                       ByVal isDayColumn As Boolean, _
                                       ByVal wsUnitPrice As Worksheet, _
                                       ByVal rowIndex As Long, _
                                       ByVal ratioAddress As String)
    Dim sourceCol As Long
    Dim workTypeName As String
    If isDayColumn Then
        sourceCol = VENDOR_UNIT_PRICE_REF_UNIT_COL
    Else
        sourceCol = VENDOR_UNIT_PRICE_REF_WIDTH_COL
    End If

    workTypeName = CommonNormalizeText(CStr(wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_WORK_TYPE_COL).value))

    With targetCell
        .ShrinkToFit = False
        .Interior.ColorIndex = xlColorIndexNone

        If IsVendorUnitPriceSourceBlank(wsUnitPrice, rowIndex, sourceCol) Then
            ApplyVendorUnitPriceGreyFill targetCell
            Exit Sub
        End If

        If InStr(1, workTypeName, VendorWasteDisposalKeywordText(), vbTextCompare) > 0 Then
            ApplyVendorUnitPriceGreyFill targetCell
            Exit Sub
        End If

        .Formula = BuildVendorUnitPriceFormula(wsUnitPrice, rowIndex, sourceCol, ratioAddress)
        .NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
    End With
End Sub

Public Sub ApplyVendorUnitPriceCellsForSourceRow(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal wsInfo As Worksheet, _
                                                   ByVal rowIndex As Long, _
                                                   ByVal isDayColumn As Boolean)
    Dim vendorCount As Long
    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = mod_VendorBlockLayout.VendorValueColumnByIndex(vendorIndex)
        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            Dim targetCol As Long
            targetCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
            If Not isDayColumn Then targetCol = targetCol + 1

            ApplyVendorUnitPriceCell wsUnitPrice.Cells(rowIndex, targetCol), _
                                     isDayColumn, wsUnitPrice, rowIndex, _
                                     GetVendorOutsourceRatioAddress(wsInfo, valueColumn)
        End If
    Next vendorIndex
End Sub

Public Sub ApplyVendorUnitPriceColumnWidths(ByVal wsUnitPrice As Worksheet, _
                                               ByVal dayCol As Long, _
                                               ByVal nightCol As Long)
    wsUnitPrice.Columns(dayCol).ColumnWidth = wsUnitPrice.Columns(VENDOR_UNIT_PRICE_REF_UNIT_COL).ColumnWidth
    wsUnitPrice.Columns(nightCol).ColumnWidth = wsUnitPrice.Columns(VENDOR_UNIT_PRICE_REF_WIDTH_COL).ColumnWidth
End Sub

Public Sub ApplyVendorUnitPriceDataColumn(ByVal wsUnitPrice As Worksheet, _
                                           ByVal wsInfo As Worksheet, _
                                           ByVal valueColumn As Long, _
                                           ByVal targetCol As Long, _
                                           ByVal sourceCol As Long, _
                                           ByVal formulaR1C1 As String, _
                                           ByVal wasteKeyword As String, _
                                           ByVal lastRow As Long, _
                                           ByVal isDayColumn As Boolean, _
                                           Optional ByVal usePreloadedArrays As Boolean = False, _
                                           Optional ByRef preloadedSrcArr As Variant, _
                                           Optional ByRef preloadedWorkArr As Variant)
    ' 行ごとにセルを個別読み取りすると単価シート(数千行)×列×社数でCOM往復が
    ' 爆発するため、判定に必要な列(単価元 sourceCol / 工種名 C列)を配列で一括読み取りし、
    ' メモリ上で「数式を入れる行/グレー塗りにする行」を判定する。書き込みは従来通り
    ' 連続行をまとめてセグメント単位で適用する(結果は行単位処理と完全に同一)。
    If lastRow < VENDOR_UNIT_PRICE_DATA_START_ROW Then Exit Sub

    Dim firstRow As Long
    firstRow = VENDOR_UNIT_PRICE_DATA_START_ROW
    Dim rowCount As Long
    rowCount = lastRow - firstRow + 1

    Dim srcArr As Variant
    Dim workArr As Variant
    If usePreloadedArrays Then
        srcArr = preloadedSrcArr
        workArr = preloadedWorkArr
    Else
        ReadVendorUnitPriceColumnValues wsUnitPrice, firstRow, lastRow, sourceCol, srcArr
        ReadVendorUnitPriceColumnValues wsUnitPrice, firstRow, lastRow, VENDOR_UNIT_PRICE_WORK_TYPE_COL, workArr
    End If

    Dim greySegStart As Long
    Dim formulaSegStart As Long
    greySegStart = 0
    formulaSegStart = 0

    Dim r As Long
    For r = 1 To rowCount + 1
        Dim rowIndex As Long
        rowIndex = firstRow + r - 1

        Dim needsFormula As Boolean
        needsFormula = False
        If r <= rowCount Then
            needsFormula = RowNeedsFormulaFromArrays(srcArr, workArr, r, wasteKeyword)
        End If

        If needsFormula Then
            ' 数式行に切り替わる直前に、溜まっていたグレー塗りセグメントを一括適用
            If greySegStart > 0 Then
                ApplyVendorUnitPriceGreyFillRange wsUnitPrice.Range( _
                    wsUnitPrice.Cells(greySegStart, targetCol), _
                    wsUnitPrice.Cells(rowIndex - 1, targetCol))
                greySegStart = 0
            End If
            If formulaSegStart = 0 Then formulaSegStart = rowIndex
        Else
            ' グレー塗り行に切り替わる直前に、溜まっていた数式セグメントを一括適用
            If formulaSegStart > 0 Then
                ApplyVendorUnitPriceFormulaSegment wsUnitPrice, wsInfo, valueColumn, formulaSegStart, rowIndex - 1, _
                    targetCol, formulaR1C1, isDayColumn
                formulaSegStart = 0
            End If
            If r <= rowCount Then
                If greySegStart = 0 Then greySegStart = rowIndex
            Else
                ' 最終番兵行: 残ったグレー塗りセグメントを適用
                If greySegStart > 0 Then
                    ApplyVendorUnitPriceGreyFillRange wsUnitPrice.Range( _
                        wsUnitPrice.Cells(greySegStart, targetCol), _
                        wsUnitPrice.Cells(rowIndex - 1, targetCol))
                    greySegStart = 0
                End If
            End If
        End If
    Next r
End Sub

' 指定列の firstRow～lastRow を1回のCOM呼び出しで配列取得する。
' 常に (行数,1) の2次元配列(1始まり)へ正規化し、単一セルでも同じ形で扱えるようにする。

Public Sub ApplyVendorUnitPriceDataRows(ByVal wsUnitPrice As Worksheet, _
                                           ByVal wsInfo As Worksheet, _
                                           ByVal valueColumn As Long, _
                                           ByVal dayCol As Long, _
                                           ByVal nightCol As Long, _
                                           Optional ByVal cachedLastRow As Long = 0, _
                                           Optional ByVal usePreloadedArrays As Boolean = False, _
                                           Optional ByRef preloadedDaySrcArr As Variant, _
                                           Optional ByRef preloadedNightSrcArr As Variant, _
                                           Optional ByRef preloadedWorkArr As Variant)
    Dim lastRow As Long
    If cachedLastRow >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
        lastRow = cachedLastRow
    Else
        lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    End If
    If lastRow < VENDOR_UNIT_PRICE_DATA_START_ROW Then Exit Sub

    Dim ratioAddress As String
    ratioAddress = GetVendorOutsourceRatioAddress(wsInfo, valueColumn)

    Dim dayFormulaR1C1 As String
    Dim nightFormulaR1C1 As String
    dayFormulaR1C1 = BuildVendorUnitPriceFormulaR1C1(True, dayCol, ratioAddress)
    nightFormulaR1C1 = BuildVendorUnitPriceFormulaR1C1(False, nightCol, ratioAddress)

    Dim wasteKeyword As String
    wasteKeyword = VendorWasteDisposalKeywordText()

    ApplyVendorUnitPriceDataColumn wsUnitPrice, wsInfo, valueColumn, dayCol, VENDOR_UNIT_PRICE_REF_UNIT_COL, _
        dayFormulaR1C1, wasteKeyword, lastRow, True, usePreloadedArrays, preloadedDaySrcArr, preloadedWorkArr
    ApplyVendorUnitPriceDataColumn wsUnitPrice, wsInfo, valueColumn, nightCol, VENDOR_UNIT_PRICE_REF_WIDTH_COL, _
        nightFormulaR1C1, wasteKeyword, lastRow, False, usePreloadedArrays, preloadedNightSrcArr, preloadedWorkArr
End Sub

Public Sub ApplyVendorUnitPriceFont(ByVal wsUnitPrice As Worksheet, _
                                     ByVal dayCol As Long, _
                                     ByVal nightCol As Long, _
                                     Optional ByVal cachedLastRow As Long = 0)
    Dim lastRow As Long
    If cachedLastRow >= VENDOR_UNIT_PRICE_HEADER_ROW Then
        lastRow = cachedLastRow
    Else
        lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    End If
    If lastRow < VENDOR_UNIT_PRICE_HEADER_ROW Then lastRow = VENDOR_UNIT_PRICE_DATA_START_ROW

    Dim fontRange As Range
    Set fontRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol), _
                                      wsUnitPrice.Cells(lastRow, nightCol))

    With fontRange.Font
        .Name = VendorUnitPriceFontNameText()
        On Error Resume Next
        .NameFarEast = VendorUnitPriceFontNameText()
        On Error GoTo 0
    End With
End Sub

Public Sub ApplyVendorUnitPriceFormulaSegment(ByVal wsUnitPrice As Worksheet, _
                                               ByVal wsInfo As Worksheet, _
                                               ByVal valueColumn As Long, _
                                               ByVal segStart As Long, _
                                               ByVal segEnd As Long, _
                                               ByVal targetCol As Long, _
                                               ByVal formulaR1C1 As String, _
                                               ByVal isDayColumn As Boolean)
    If segStart > segEnd Then Exit Sub

    On Error GoTo RowFallback
    With wsUnitPrice.Range(wsUnitPrice.Cells(segStart, targetCol), wsUnitPrice.Cells(segEnd, targetCol))
        .FormulaR1C1 = formulaR1C1
        .NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
        .Interior.ColorIndex = xlColorIndexNone
        .ShrinkToFit = False
    End With
    Exit Sub

RowFallback:
    Err.Clear
    Dim ratioAddress As String
    Dim rowIndex As Long
    ratioAddress = GetVendorOutsourceRatioAddress(wsInfo, valueColumn)
    For rowIndex = segStart To segEnd
        ApplyVendorUnitPriceCell wsUnitPrice.Cells(rowIndex, targetCol), isDayColumn, _
                                 wsUnitPrice, rowIndex, ratioAddress
    Next rowIndex
End Sub

Public Sub ApplyVendorUnitPriceGreyFill(ByVal targetCell As Range)
    With targetCell
        .ClearContents
        .NumberFormat = "General"
        .Interior.Color = RGB(VENDOR_UNIT_PRICE_FILL_COLOR_R, _
                              VENDOR_UNIT_PRICE_FILL_COLOR_G, _
                              VENDOR_UNIT_PRICE_FILL_COLOR_B)
    End With
End Sub

Public Sub ApplyVendorUnitPriceGreyFillRange(ByVal targetRange As Range)
    With targetRange
        .ClearContents
        .NumberFormat = "General"
        .Interior.Color = RGB(VENDOR_UNIT_PRICE_FILL_COLOR_R, _
                              VENDOR_UNIT_PRICE_FILL_COLOR_G, _
                              VENDOR_UNIT_PRICE_FILL_COLOR_B)
    End With
End Sub

Public Sub ApplyVendorUnitPriceJrHeader(ByVal wsUnitPrice As Worksheet, ByVal wsInfo As Worksheet)
    Dim headerRange As Range
    Set headerRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_JR_HEADER_ROW, VENDOR_UNIT_PRICE_JR_HEADER_COL_START), _
                                        wsUnitPrice.Cells(VENDOR_UNIT_PRICE_JR_HEADER_ROW, VENDOR_UNIT_PRICE_JR_HEADER_COL_END))

    mod_VendorBlockLayout.SafeUnmergeRange headerRange
    headerRange.Merge
    With headerRange
        .value = BuildVendorUnitPriceJrHeaderText(wsInfo)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .WrapText = False
        .Font.Name = VendorUnitPriceFontNameText()
        On Error Resume Next
        .Font.NameFarEast = VendorUnitPriceFontNameText()
        On Error GoTo 0
    End With

    With headerRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Public Sub ApplyVendorUnitPriceMergedHeader(ByVal wsUnitPrice As Worksheet, _
                                               ByVal dayCol As Long, _
                                               ByVal nightCol As Long, _
                                               ByVal headerText As String)
    Dim headerRange As Range
    Set headerRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol), _
                                        wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, nightCol))

    mod_VendorBlockLayout.SafeUnmergeRange headerRange
    headerRange.Merge
    With headerRange
        .value = headerText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = False
        .WrapText = False
    End With
End Sub

Public Sub ApplyVendorUnitPriceMergedVendorName(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal dayCol As Long, _
                                                   ByVal nightCol As Long, _
                                                   ByVal vendorName As String)
    Dim nameRange As Range
    Set nameRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_NAME_ROW, dayCol), _
                                      wsUnitPrice.Cells(VENDOR_UNIT_PRICE_NAME_ROW, nightCol))

    mod_VendorBlockLayout.SafeUnmergeRange nameRange
    nameRange.Merge
    With nameRange
        .value = vendorName
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .WrapText = False
    End With
End Sub

Public Sub ApplyVendorUnitPriceNewRowFill(ByVal wsUnitPrice As Worksheet, _
                                            ByVal wsInfo As Worksheet, _
                                            ByVal changedCells As Range)
    Dim lastFillCol As Long
    lastFillCol = GetVendorUnitPriceInitialFillLastColumn(wsInfo)

    Dim changedCell As Range
    For Each changedCell In changedCells.Cells
        If changedCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(changedCell.value))) > 0 Then
                With wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_REF_UNIT_COL), _
                                            wsUnitPrice.Cells(changedCell.Row, lastFillCol)).Interior
                    .Color = RGB(VENDOR_UNIT_PRICE_FILL_COLOR_R, _
                                 VENDOR_UNIT_PRICE_FILL_COLOR_G, _
                                 VENDOR_UNIT_PRICE_FILL_COLOR_B)
                End With
            End If
        End If
    Next changedCell
End Sub

Public Sub ApplyVendorUnitPriceOutsourceRatioRow(ByVal wsUnitPrice As Worksheet, _
                                                  ByVal wsInfo As Worksheet, _
                                                  ByVal valueColumn As Long, _
                                                  ByVal dayCol As Long, _
                                                  ByVal nightCol As Long)
    With wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, dayCol)
        .Formula = ""
        .value = VendorUnitPriceOutsourceRatioLabelText()
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    With wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, nightCol)
        .Formula = ""
        .value = GetVendorOutsourceRatioPercentValue(wsInfo, valueColumn)
        .NumberFormat = VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_NUMBER_FORMAT
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ApplyVendorUnitPriceOutsourceRatioRowFont wsUnitPrice, dayCol, nightCol
End Sub

Public Sub ApplyVendorUnitPriceOutsourceRatioRowFont(ByVal wsUnitPrice As Worksheet, _
                                                      ByVal dayCol As Long, _
                                                      ByVal nightCol As Long)
    Dim fontRange As Range
    Set fontRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, dayCol), _
                                      wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, nightCol))

    With fontRange.Font
        .Name = VendorUnitPriceFontNameText()
        .Size = VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_FONT_SIZE
        On Error Resume Next
        .NameFarEast = VendorUnitPriceFontNameText()
        On Error GoTo 0
    End With
End Sub

Public Sub ApplyVendorUnitPriceSourceCellDecoration(ByVal sourceCell As Range, ByVal srcValue As Variant)
    If IsNumericSourceValue(srcValue) Then
        sourceCell.Interior.ColorIndex = xlColorIndexNone
        sourceCell.NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
    ElseIf IsBlankSourceValue(srcValue) Then
        ApplyVendorUnitPriceSourceGreyFill sourceCell
    End If
End Sub

' 全展開経路用の高速版。E列/F列の桁区切り書式を、B列に値のある連続行セグメント単位で
' 範囲一括適用する(結果は行単位の NumberFormat 設定と同一)。

Public Sub ApplyVendorUnitPriceSourceColumnsNumberFormat(ByVal wsUnitPrice As Worksheet, _
                                                          ByVal firstRow As Long, _
                                                          ByVal lastRow As Long)
    If wsUnitPrice Is Nothing Then Exit Sub
    If lastRow < firstRow Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = firstRow To lastRow
        If Len(Trim$(CStr(wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) > 0 Then
            wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_REF_UNIT_COL).NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
            wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_REF_WIDTH_COL).NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
        End If
    Next rowIndex
End Sub

Public Sub ApplyVendorUnitPriceSourceColumnsNumberFormatFast(ByVal wsUnitPrice As Worksheet, _
                                                              ByVal firstRow As Long, _
                                                              ByVal lastRow As Long, _
                                                              ByVal bArr As Variant)
    Dim rowCount As Long
    rowCount = lastRow - firstRow + 1
    If rowCount <= 0 Then Exit Sub

    Dim segStart As Long
    segStart = 0

    Dim r As Long
    For r = 1 To rowCount + 1
        Dim rowIndex As Long
        rowIndex = firstRow + r - 1

        Dim hasValue As Boolean
        hasValue = False
        If r <= rowCount Then hasValue = (Len(Trim$(CStr(CommonNzText(bArr(r, 1))))) > 0)

        If hasValue Then
            If segStart = 0 Then segStart = rowIndex
        Else
            If segStart > 0 Then
                wsUnitPrice.Range(wsUnitPrice.Cells(segStart, VENDOR_UNIT_PRICE_REF_UNIT_COL), _
                                  wsUnitPrice.Cells(rowIndex - 1, VENDOR_UNIT_PRICE_REF_WIDTH_COL)).NumberFormat = _
                    VENDOR_UNIT_PRICE_NUMBER_FORMAT
                segStart = 0
            End If
        End If
    Next r
End Sub

Public Sub ApplyVendorUnitPriceSourceEfDecorationsFast(ByVal wsUnitPrice As Worksheet, _
                                                        ByVal firstRow As Long, _
                                                        ByVal lastRow As Long, _
                                                        ByVal bArr As Variant)
    Dim rowCount As Long
    rowCount = lastRow - firstRow + 1
    If rowCount <= 0 Then Exit Sub

    Dim eArr As Variant
    Dim fArr As Variant
    ReadVendorUnitPriceColumnValues wsUnitPrice, firstRow, lastRow, VENDOR_UNIT_PRICE_REF_UNIT_COL, eArr
    ReadVendorUnitPriceColumnValues wsUnitPrice, firstRow, lastRow, VENDOR_UNIT_PRICE_REF_WIDTH_COL, fArr

    Dim r As Long
    For r = 1 To rowCount
        If Len(Trim$(CStr(CommonNzText(bArr(r, 1))))) = 0 Then GoTo ContinueRow

        Dim rowIndex As Long
        rowIndex = firstRow + r - 1
        ApplyVendorUnitPriceSourceCellDecoration wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_REF_UNIT_COL), eArr(r, 1)
        ApplyVendorUnitPriceSourceCellDecoration wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_REF_WIDTH_COL), fArr(r, 1)
ContinueRow:
    Next r
End Sub

Public Sub ApplyVendorUnitPriceSourceGreyFill(ByVal sourceCell As Range)
    sourceCell.Interior.Color = RGB(VENDOR_UNIT_PRICE_FILL_COLOR_R, _
                                    VENDOR_UNIT_PRICE_FILL_COLOR_G, _
                                    VENDOR_UNIT_PRICE_FILL_COLOR_B)
End Sub

Public Sub ApplyVendorUnitPriceSourceRowIfNeeded(ByVal wsUnitPrice As Worksheet, _
                                                  ByVal wsInfo As Worksheet, _
                                                  ByVal rowIndex As Long, _
                                                  ByVal sourceCol As Long, _
                                                  ByVal isDayColumn As Boolean)
    Dim sourceCell As Range
    Set sourceCell = wsUnitPrice.Cells(rowIndex, sourceCol)

    If HasNumericVendorUnitPriceSource(sourceCell) Then
        sourceCell.Interior.ColorIndex = xlColorIndexNone
        sourceCell.NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
        ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, rowIndex, isDayColumn
    ElseIf IsVendorUnitPriceSourceCellBlank(sourceCell) Then
        ApplyVendorUnitPriceSourceGreyFill sourceCell
        ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, rowIndex, isDayColumn
    End If
End Sub

Public Sub ApplyVendorUnitPriceSourceRowIfNeededFromValue(ByVal wsUnitPrice As Worksheet, _
                                                           ByVal wsInfo As Worksheet, _
                                                           ByVal rowIndex As Long, _
                                                           ByVal sourceCol As Long, _
                                                           ByVal isDayColumn As Boolean, _
                                                           ByVal srcValue As Variant)
    Dim sourceCell As Range
    Set sourceCell = wsUnitPrice.Cells(rowIndex, sourceCol)

    If IsNumericSourceValue(srcValue) Then
        sourceCell.Interior.ColorIndex = xlColorIndexNone
        sourceCell.NumberFormat = VENDOR_UNIT_PRICE_NUMBER_FORMAT
        ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, rowIndex, isDayColumn
    ElseIf IsBlankSourceValue(srcValue) Then
        ApplyVendorUnitPriceSourceGreyFill sourceCell
        ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, rowIndex, isDayColumn
    End If
End Sub

' HasNumericVendorUnitPriceSource(セル版)と同一判定を、読み取り済み値に対して行う。

Public Sub ApplyVendorUnitPriceSourceRowsForRange(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal wsInfo As Worksheet, _
                                                   ByVal firstRow As Long, _
                                                   ByVal lastRow As Long)
    Dim rowIndex As Long
    For rowIndex = firstRow To lastRow
        If Len(Trim$(CStr(wsUnitPrice.Cells(rowIndex, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) = 0 Then GoTo ContinueRow

        ApplyVendorUnitPriceSourceRowIfNeeded wsUnitPrice, wsInfo, rowIndex, _
            VENDOR_UNIT_PRICE_REF_UNIT_COL, True
        ApplyVendorUnitPriceSourceRowIfNeeded wsUnitPrice, wsInfo, rowIndex, _
            VENDOR_UNIT_PRICE_REF_WIDTH_COL, False
ContinueRow:
    Next rowIndex
End Sub

' 全展開経路用の高速版。B列(判定用)/E列/F列を配列で一括読み取りし、行判定をメモリで行う。
' 各行の書式適用(グレー塗り/桁区切り)と業者列への数式反映は既存関数を呼ぶため結果は同一。
' bArr は呼び出し元で読み取り済みのB列配列(1始まり, firstRow基準)。

Public Sub ApplyVendorUnitPriceSourceRowsForRangeFast(ByVal wsUnitPrice As Worksheet, _
                                                       ByVal wsInfo As Worksheet, _
                                                       ByVal firstRow As Long, _
                                                       ByVal lastRow As Long, _
                                                       ByVal bArr As Variant)
    Dim rowCount As Long
    rowCount = lastRow - firstRow + 1
    If rowCount <= 0 Then Exit Sub

    ' E列(昼単価元)/F列(夜単価元)を一括読み取り
    Dim eArr As Variant
    Dim fArr As Variant
    ReadVendorUnitPriceColumnValues wsUnitPrice, firstRow, lastRow, VENDOR_UNIT_PRICE_REF_UNIT_COL, eArr
    ReadVendorUnitPriceColumnValues wsUnitPrice, firstRow, lastRow, VENDOR_UNIT_PRICE_REF_WIDTH_COL, fArr

    Dim r As Long
    For r = 1 To rowCount
        If Len(Trim$(CStr(CommonNzText(bArr(r, 1))))) = 0 Then GoTo ContinueRow

        Dim rowIndex As Long
        rowIndex = firstRow + r - 1

        ApplyVendorUnitPriceSourceRowIfNeededFromValue wsUnitPrice, wsInfo, rowIndex, _
            VENDOR_UNIT_PRICE_REF_UNIT_COL, True, eArr(r, 1)
        ApplyVendorUnitPriceSourceRowIfNeededFromValue wsUnitPrice, wsInfo, rowIndex, _
            VENDOR_UNIT_PRICE_REF_WIDTH_COL, False, fArr(r, 1)
ContinueRow:
    Next r
End Sub

' ApplyVendorUnitPriceSourceRowIfNeeded と同一処理だが、単価元セルの値は
' 事前読み取り済みの srcValue を使い、セル値の再読み取りを避ける。

Public Sub ClearVendorUnitPriceBlockOnSheet(ByVal wsUnitPrice As Worksheet, _
                                             ByVal dayCol As Long, _
                                             ByVal nightCol As Long, _
                                             Optional ByVal cachedLastRow As Long = 0)
    If wsUnitPrice Is Nothing Then Exit Sub

    ClearVendorUnitPriceOutsourceRatioRow wsUnitPrice, dayCol, nightCol

    Dim lastRow As Long
    If cachedLastRow >= VENDOR_UNIT_PRICE_HEADER_ROW Then
        lastRow = cachedLastRow
    Else
        lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    End If
    If lastRow < VENDOR_UNIT_PRICE_HEADER_ROW Then lastRow = VENDOR_UNIT_PRICE_DATA_START_ROW + 200

    Dim clearRange As Range
    Set clearRange = wsUnitPrice.Range(wsUnitPrice.Cells(VENDOR_UNIT_PRICE_HEADER_ROW, dayCol), _
                                       wsUnitPrice.Cells(lastRow, nightCol))

    mod_VendorBlockLayout.SafeUnmergeRange clearRange
    clearRange.ClearContents
    clearRange.Interior.ColorIndex = xlColorIndexNone
    clearRange.ShrinkToFit = False
    clearRange.Borders.LineStyle = xlNone
End Sub

Public Sub ClearVendorUnitPriceOutsourceRatioRow(ByVal wsUnitPrice As Worksheet, _
                                                  ByVal dayCol As Long, _
                                                  ByVal nightCol As Long)
    With wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, dayCol)
        .ClearContents
        .HorizontalAlignment = xlGeneral
    End With

    With wsUnitPrice.Cells(VENDOR_UNIT_PRICE_OUTSOURCE_RATIO_ROW, nightCol)
        .ClearContents
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
    End With
End Sub

Public Sub ClearVendorUnitPriceRemovedBlocksOnSheetSafely(ByVal wsUnitPrice As Worksheet, _
                                                           ByVal previousCount As Long, _
                                                           ByVal vendorCount As Long)
    On Error GoTo ErrorHandler

    Dim clearLastIndex As Long
    clearLastIndex = previousCount
    If clearLastIndex > MAX_VENDOR_BLOCK_COUNT Then clearLastIndex = MAX_VENDOR_BLOCK_COUNT

    Dim blockIndex As Long
    Dim dayCol As Long
    Dim nightCol As Long
    For blockIndex = vendorCount + 1 To clearLastIndex
        dayCol = VendorUnitPriceDayColumnByValueColumn(mod_VendorBlockLayout.VendorValueColumnByIndex(blockIndex))
        nightCol = dayCol + 1
        ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
    Next blockIndex
    Exit Sub

ErrorHandler:
    mod_DebugLog.Log "[VendorMaster] ClearVendorUnitPriceRemovedBlocksOnSheetSafely failed sheet=[" & _
                     wsUnitPrice.Name & "] Err " & Err.Number & ": " & Err.Description
End Sub

Public Sub EnsureApplicationCalculationAutomatic()
    ' 計算モードはイベント/呼び出し側で一括管理する。
    ' 一括書込み中に自動計算へ強制すると、セル書込みごとに再計算が走り低速化するため、
    ' ここでは計算モードを変更しない(各処理終端の .Calculate と呼び出し側の復元で再計算)。
End Sub

Public Sub EnsureVendorUnitPriceNewRowFillForSourceRows(ByVal wsUnitPrice As Worksheet, _
                                                         ByVal wsInfo As Worksheet, _
                                                         ByVal sourceChangedCells As Range)
    If wsUnitPrice Is Nothing Then Exit Sub
    If wsInfo Is Nothing Then Exit Sub
    If sourceChangedCells Is Nothing Then Exit Sub

    Dim bColRange As Range
    Dim sourceCell As Range
    For Each sourceCell In sourceChangedCells.Cells
        If sourceCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(wsUnitPrice.Cells(sourceCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) > 0 Then
                If bColRange Is Nothing Then
                    Set bColRange = wsUnitPrice.Cells(sourceCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL)
                Else
                    Set bColRange = Union(bColRange, wsUnitPrice.Cells(sourceCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL))
                End If
            End If
        End If
    Next sourceCell

    If Not bColRange Is Nothing Then
        ApplyVendorUnitPriceNewRowFill wsUnitPrice, wsInfo, bColRange
    End If
End Sub

Public Sub HandleVendorUnitPriceSourceChanges(ByVal wsUnitPrice As Worksheet, _
                                                ByVal wsInfo As Worksheet, _
                                                ByVal changedCells As Range, _
                                                ByVal isDayColumn As Boolean)
    Dim sourceCell As Range
    For Each sourceCell In changedCells.Cells
        If sourceCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(wsUnitPrice.Cells(sourceCell.Row, VENDOR_UNIT_PRICE_LAST_ROW_COL).value))) = 0 Then
                GoTo ContinueNextSourceCell
            End If

            If HasNumericVendorUnitPriceSource(sourceCell) Then
                sourceCell.Interior.ColorIndex = xlColorIndexNone
                ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, sourceCell.Row, isDayColumn
            ElseIf IsVendorUnitPriceSourceCellBlank(sourceCell) Then
                ApplyVendorUnitPriceSourceGreyFill sourceCell
                ApplyVendorUnitPriceCellsForSourceRow wsUnitPrice, wsInfo, sourceCell.Row, isDayColumn
            End If
        End If
ContinueNextSourceCell:
    Next sourceCell
End Sub

' 独自工種などプログラム追記行は B 列の Worksheet_Change を経由しないため、
' E/F 入力時に手入力追記と同じ新規行初期化を補完する。

Public Sub ReadVendorUnitPriceColumnValues(ByVal wsUnitPrice As Worksheet, _
                                            ByVal firstRow As Long, _
                                            ByVal lastRow As Long, _
                                            ByVal targetCol As Long, _
                                            ByRef outArr As Variant)
    Dim rowCount As Long
    rowCount = lastRow - firstRow + 1
    If rowCount <= 0 Then
        outArr = Empty
        Exit Sub
    End If

    If rowCount = 1 Then
        ReDim outArr(1 To 1, 1 To 1)
        outArr(1, 1) = wsUnitPrice.Cells(firstRow, targetCol).Value
    Else
        outArr = wsUnitPrice.Range(wsUnitPrice.Cells(firstRow, targetCol), _
                                   wsUnitPrice.Cells(lastRow, targetCol)).Value
    End If
End Sub

' 配列(1始まり)の r 行目について、単価元が空でなく、工種名に産廃キーワードを
' 含まない場合に True。VendorUnitPriceRowNeedsFormulaForSource と同一判定。

Public Sub RefreshVendorUnitPriceBlocksOnSheet(ByVal wsUnitPrice As Worksheet, _
                                                 ByVal wsInfo As Worksheet, _
                                                 ByVal vendorCount As Long, _
                                                 ByVal vendorUnitPriceNameMap As Object)
    Dim sheetLastRow As Long
    sheetLastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)

    Dim usePreloadedArrays As Boolean
    Dim preloadedDaySrcArr As Variant
    Dim preloadedNightSrcArr As Variant
    Dim preloadedWorkArr As Variant
    usePreloadedArrays = False
    If sheetLastRow >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
        ReadVendorUnitPriceColumnValues wsUnitPrice, VENDOR_UNIT_PRICE_DATA_START_ROW, sheetLastRow, _
            VENDOR_UNIT_PRICE_REF_UNIT_COL, preloadedDaySrcArr
        ReadVendorUnitPriceColumnValues wsUnitPrice, VENDOR_UNIT_PRICE_DATA_START_ROW, sheetLastRow, _
            VENDOR_UNIT_PRICE_REF_WIDTH_COL, preloadedNightSrcArr
        ReadVendorUnitPriceColumnValues wsUnitPrice, VENDOR_UNIT_PRICE_DATA_START_ROW, sheetLastRow, _
            VENDOR_UNIT_PRICE_WORK_TYPE_COL, preloadedWorkArr
        usePreloadedArrays = True
    End If

    Dim i As Long
    For i = 1 To vendorCount
        Dim valueColumn As Long
        Dim dayCol As Long
        Dim nightCol As Long
        valueColumn = mod_VendorBlockLayout.VendorValueColumnByIndex(i)
        dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
        nightCol = dayCol + 1

        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn, vendorUnitPriceNameMap, _
                sheetLastRow, usePreloadedArrays, preloadedDaySrcArr, preloadedNightSrcArr, preloadedWorkArr
        Else
            ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol, sheetLastRow
        End If
    Next i

    For i = vendorCount + 1 To MAX_VENDOR_BLOCK_COUNT
        valueColumn = mod_VendorBlockLayout.VendorValueColumnByIndex(i)
        dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
        nightCol = dayCol + 1
        ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol, sheetLastRow
    Next i
End Sub

Public Sub RefreshVendorUnitPriceBordersForSheet(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal wsInfo As Worksheet)
    Dim vendorCount As Long
    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)

    Dim vendorIndex As Long
    For vendorIndex = 1 To vendorCount
        Dim valueColumn As Long
        valueColumn = mod_VendorBlockLayout.VendorValueColumnByIndex(vendorIndex)
        If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
            Dim dayCol As Long
            dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
            ApplyVendorUnitPriceBorders wsUnitPrice, dayCol, dayCol + 1
        End If
    Next vendorIndex
End Sub

Public Sub RefreshVendorUnitPriceForValueColumn(ByVal wsInfo As Worksheet, ByVal valueColumn As Long)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim dayCol As Long
    Dim nightCol As Long
    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
    nightCol = dayCol + 1

    Dim vendorUnitPriceNameMap As Object
    Set vendorUnitPriceNameMap = BuildVendorUnitPriceNameMap(wsInfo)

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) And mod_MaterialPriceImport.IsCurrentImportBatchUnitPriceSheet(wsUnitPrice) Then
            If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
                ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn, vendorUnitPriceNameMap
            ElseIf IsRailConstructionVendorBlock(wsInfo, valueColumn) And _
                   HasVendorName(wsInfo, valueColumn) Then
                ' 11行目のみ入力済み。29行目入力待ちの間は既存列を消さない
            Else
                ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
            End If
        End If
    Next wsUnitPrice
End Sub

' 外注比率(29行目)のみが変わった場合の軽量更新。
' 単価行の数式(day/night列)は基本情報シートの比率セルを直接参照(絶対参照)しているため、
' 比率の値が変わっても数式自体を書き直す必要はなく、Excelの再計算で自動的に反映される。
' そのため、ブロックが構築済みの単価シートに対しては比率表示欄(ラベル・値の2セル)だけを
' 更新し、見出し結合・罫線・フォント・数式の全再構築(ApplyVendorUnitPriceBlockToSheet)は行わない。
' 業者名は入力済みだが比率が今回初めて入力された(ブロック未構築)単価シートに対しては、
' 従来通り全展開(ApplyVendorUnitPriceBlockToSheet)にフォールバックする。

Public Sub RefreshVendorUnitPriceOutsourceRatioOnlyForValueColumn(ByVal wsInfo As Worksheet, ByVal valueColumn As Long)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim dayCol As Long
    Dim nightCol As Long
    dayCol = VendorUnitPriceDayColumnByValueColumn(valueColumn)
    nightCol = dayCol + 1

    Dim vendorUnitPriceNameMap As Object
    Dim nameMapLoaded As Boolean
    nameMapLoaded = False

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) And mod_MaterialPriceImport.IsCurrentImportBatchUnitPriceSheet(wsUnitPrice) Then
            If ShouldApplyVendorUnitPriceBlock(wsInfo, valueColumn) Then
                If IsVendorUnitPriceBlockAlreadyBuilt(wsUnitPrice, dayCol, nightCol) Then
                    ApplyVendorUnitPriceOutsourceRatioRow wsUnitPrice, wsInfo, valueColumn, dayCol, nightCol
                Else
                    If Not nameMapLoaded Then
                        Set vendorUnitPriceNameMap = BuildVendorUnitPriceNameMap(wsInfo)
                        nameMapLoaded = True
                    End If
                    ApplyVendorUnitPriceBlockToSheet wsUnitPrice, wsInfo, valueColumn, vendorUnitPriceNameMap
                End If
            ElseIf IsRailConstructionVendorBlock(wsInfo, valueColumn) And _
                   HasVendorName(wsInfo, valueColumn) Then
                ' 11行目のみ入力済み。29行目入力待ちの間は既存列を消さない
            Else
                ClearVendorUnitPriceBlockOnSheet wsUnitPrice, dayCol, nightCol
            End If
        End If
    Next wsUnitPrice
End Sub

' 単価シート側に指定列の業者ブロック(見出し結合セル)が既に構築済みかどうかを判定する。
' 未構築(初回)の場合は全展開が必要なため、呼び出し元でフォールバックする。

Public Sub RefreshVendorUnitPriceSheetForSyncSafely(ByVal wsUnitPrice As Worksheet, _
                                                     ByVal wsInfo As Worksheet, _
                                                     ByVal vendorCount As Long, _
                                                     ByVal vendorUnitPriceNameMap As Object)
    On Error GoTo ErrorHandler
    RefreshVendorUnitPriceBlocksOnSheet wsUnitPrice, wsInfo, vendorCount, vendorUnitPriceNameMap
    ' 業者列は直前の RefreshVendorUnitPriceBlocksOnSheet で展開済みのため、
    ' JR列(A～F)の装飾とE/F桁区切りだけ適用し業者列の二重走査を避ける。
    RefreshConstructionUnitPriceSheetDataDecorations wsUnitPrice, wsInfo, True
    Exit Sub

ErrorHandler:
    mod_DebugLog.Log "[VendorMaster] RefreshVendorUnitPriceSheetForSyncSafely failed sheet=[" & _
                     wsUnitPrice.Name & "] Err " & Err.Number & ": " & Err.Description
End Sub

' F9増加時: 追加分の業者ブロックだけを1シートへ適用する。エラーはログして続行。

Public Sub ResetClearedVendorUnitPriceRows(ByVal wsUnitPrice As Worksheet, _
                                             ByVal wsInfo As Worksheet, _
                                             ByVal changedCells As Range)
    Dim lastResetCol As Long
    lastResetCol = GetVendorUnitPriceInitialFillLastColumn(wsInfo)

    Dim changedCell As Range
    For Each changedCell In changedCells.Cells
        If changedCell.Row >= VENDOR_UNIT_PRICE_DATA_START_ROW Then
            If Len(Trim$(CStr(changedCell.value))) = 0 Then
                Dim resetRange As Range
                Set resetRange = wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, 1), _
                                                   wsUnitPrice.Cells(changedCell.Row, lastResetCol))
                resetRange.Borders.LineStyle = xlNone

                wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_REF_UNIT_COL), _
                                  wsUnitPrice.Cells(changedCell.Row, lastResetCol)).Interior.ColorIndex = xlColorIndexNone

                With wsUnitPrice.Range(wsUnitPrice.Cells(changedCell.Row, VENDOR_UNIT_PRICE_FIRST_DAY_COL), _
                                       wsUnitPrice.Cells(changedCell.Row, lastResetCol))
                    .ClearContents
                    .NumberFormat = "General"
                    .ShrinkToFit = False
                End With
            End If
        End If
    Next changedCell
End Sub

Public Sub SyncVendorUnitPriceBlocksAfterCountChange(ByVal wsInfo As Worksheet, _
                                                      ByVal vendorCount As Long, _
                                                      ByVal previousCount As Long, _
                                                      ByVal deferCalculation As Boolean)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Dim vendorUnitPriceNameMap As Object
    Dim wsUnitPrice As Worksheet
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    Set targetBook = wsInfo.Parent
    previousScreenUpdating = Application.screenUpdating
    previousCalculation = Application.Calculation
    Application.screenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo SyncCleanup

    If previousCount <= 0 Or vendorCount = previousCount Then
        Set vendorUnitPriceNameMap = BuildVendorUnitPriceNameMap(wsInfo)

        For Each wsUnitPrice In targetBook.worksheets
            If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) And mod_MaterialPriceImport.IsCurrentImportBatchUnitPriceSheet(wsUnitPrice) Then
                RefreshVendorUnitPriceSheetForSyncSafely wsUnitPrice, wsInfo, vendorCount, vendorUnitPriceNameMap
            End If
        Next wsUnitPrice

        If Not deferCalculation Then
            On Error Resume Next
            Application.Calculate
            On Error GoTo 0
        End If
        GoTo SyncCleanup
    End If

    If vendorCount > previousCount Then
        Set vendorUnitPriceNameMap = BuildVendorUnitPriceNameMap(wsInfo)
        For Each wsUnitPrice In targetBook.worksheets
            If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) And mod_MaterialPriceImport.IsCurrentImportBatchUnitPriceSheet(wsUnitPrice) Then
                ApplyVendorUnitPriceAddedBlocksToSheetSafely wsUnitPrice, wsInfo, previousCount, vendorCount, vendorUnitPriceNameMap
            End If
        Next wsUnitPrice
        GoTo SyncCleanup
    End If

    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) And mod_MaterialPriceImport.IsCurrentImportBatchUnitPriceSheet(wsUnitPrice) Then
            ClearVendorUnitPriceRemovedBlocksOnSheetSafely wsUnitPrice, previousCount, vendorCount
        End If
    Next wsUnitPrice

SyncCleanup:
    If Err.Number <> 0 Then
        mod_DebugLog.Log "[VendorMaster] SyncVendorUnitPriceBlocksAfterCountChange Err " & _
                         Err.Number & ": " & Err.Description
        Err.Clear
    End If
    Application.screenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
End Sub

' 単価シート1枚分の全社ブロック再展開＋装飾(罫線・桁区切り・塗り)。
' 1シートでエラーが出てもログに記録して続行し、
' 以降のシートが未装飾のまま取り残されるのを防ぐ。

Public Function BuildVendorUnitPriceNameMap(ByVal wsInfo As Worksheet) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim BranchName As String
    BranchName = CommonNormalizeText(CStr(wsInfo.Range("B6").value))
    If BranchName = "" Then
        Set BuildVendorUnitPriceNameMap = result
        Exit Function
    End If

    Dim vendorRows As Collection
    Set vendorRows = mod_VendorMaster.LoadVendorRows(BranchName)
    If Not mod_VendorMaster.HasVendorRows(vendorRows) Then
        Set BuildVendorUnitPriceNameMap = result
        Exit Function
    End If

    Dim rowData As Variant
    For Each rowData In vendorRows
        Dim vendorNameKey As String
        Dim unitPriceVendorName As String
        vendorNameKey = CommonNormalizeText(CStr(rowData(VENDOR_ROW_NAME_INDEX)))
        unitPriceVendorName = CommonNzText(rowData(VENDOR_ROW_UNIT_PRICE_NAME_INDEX))

        If vendorNameKey <> "" And unitPriceVendorName <> "" Then
            If Not result.Exists(vendorNameKey) Then
                result.Add vendorNameKey, unitPriceVendorName
            End If
        End If
    Next rowData

    Set BuildVendorUnitPriceNameMap = result
End Function

Public Function GetVendorUnitPriceMonitorRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function

    Dim vendorCount As Long
    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)

    Dim result As Range
    Dim i As Long
    For i = 1 To vendorCount
        Dim valueCol As Long
        valueCol = mod_VendorBlockLayout.VendorValueColumnByIndex(i)

        Dim blockRange As Range
        Set blockRange = Union(wsInfo.Cells(BASIC_INFO_VENDOR_BLOCK_TOP_ROW, valueCol), _
                               wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueCol), _
                               wsInfo.Cells(BASIC_INFO_VENDOR_OUTSOURCE_RATIO_ROW, valueCol), _
                               wsInfo.Cells(BASIC_INFO_VENDOR_RAIL_PATTERN_ROW, valueCol), _
                               wsInfo.Cells(BASIC_INFO_VENDOR_WELDING_RATIO_ROW, valueCol))

        If result Is Nothing Then
            Set result = blockRange
        Else
            Set result = Union(result, blockRange)
        End If
    Next i

    Dim yearBillingRange As Range
    Set yearBillingRange = wsInfo.Range(BASIC_INFO_YEAR_CELL & "," & BASIC_INFO_BILLING_COUNT_CELL)
    If result Is Nothing Then
        Set result = yearBillingRange
    Else
        Set result = Union(result, yearBillingRange)
    End If
    Set GetVendorUnitPriceMonitorRange = result
End Function

Public Sub ApplyConstructionUnitPriceImportedRowDecorations(ByVal wsUnitPrice As Worksheet, _
                                                            ByVal firstRow As Long, _
                                                            ByVal lastRow As Long)
    If wsUnitPrice Is Nothing Then Exit Sub
    If lastRow < firstRow Then Exit Sub
    If firstRow < VENDOR_UNIT_PRICE_DATA_START_ROW Then Exit Sub

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(wsUnitPrice.Parent)
    If wsInfo Is Nothing Then Exit Sub

    ' 全展開(firstRow:lastRowの連続範囲)では、B列(最終行判定用)を1回で一括読み取りし、
    ' 「B列に値がある行」を連続セグメントにまとめて罫線・塗り・桁区切りを範囲一括適用する。
    ' 従来のセル単位ループ(数千行×COM往復)を排除する。手動編集時の飛び飛び範囲は
    ' 従来通り ApplyVendorUnitPriceNewRowFill 等の汎用関数が処理する(そちらは温存)。
    ApplyConstructionUnitPriceImportedRowDecorationsFast wsUnitPrice, wsInfo, firstRow, lastRow, False
End Sub

Public Sub ApplyImportedUnitPriceJrHeadersForBasicInfo(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Dim targetBook As Workbook
    Set targetBook = wsInfo.Parent

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In targetBook.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) And mod_MaterialPriceImport.IsCurrentImportBatchUnitPriceSheet(wsUnitPrice) Then
            ApplyVendorUnitPriceJrHeader wsUnitPrice, wsInfo
        End If
    Next wsUnitPrice
End Sub

Public Sub HandleConstructionUnitPriceSheetChange(ByVal wsUnitPrice As Worksheet, _
                                                   ByVal changedRange As Range)
    If wsUnitPrice Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(wsUnitPrice.Parent)
    If wsInfo Is Nothing Then Exit Sub

    EnsureApplicationCalculationAutomatic

    Dim changedB As Range
    Dim changedE As Range
    Dim changedF As Range
    Set changedB = Intersect(changedRange, wsUnitPrice.Columns(VENDOR_UNIT_PRICE_LAST_ROW_COL))
    Set changedE = Intersect(changedRange, wsUnitPrice.Columns(VENDOR_UNIT_PRICE_REF_UNIT_COL))
    Set changedF = Intersect(changedRange, wsUnitPrice.Columns(VENDOR_UNIT_PRICE_REF_WIDTH_COL))

    If Not changedB Is Nothing Then
        ApplyVendorUnitPriceNewRowFill wsUnitPrice, wsInfo, changedB
        ApplyVendorUnitPriceBaseRowBorders wsUnitPrice, wsInfo, changedB
    End If
    If Not changedE Is Nothing Then
        EnsureVendorUnitPriceNewRowFillForSourceRows wsUnitPrice, wsInfo, changedE
        HandleVendorUnitPriceSourceChanges wsUnitPrice, wsInfo, changedE, True
    End If
    If Not changedF Is Nothing Then
        EnsureVendorUnitPriceNewRowFillForSourceRows wsUnitPrice, wsInfo, changedF
        HandleVendorUnitPriceSourceChanges wsUnitPrice, wsInfo, changedF, False
    End If
    If Not changedB Is Nothing Then
        ResetClearedVendorUnitPriceRows wsUnitPrice, wsInfo, changedB
    End If

    mod_Construction_Order_Import.RefreshConstructionReferencePricesForUnitPriceChange _
        wsUnitPrice, changedRange

    On Error Resume Next
    wsUnitPrice.Calculate
    On Error GoTo 0
End Sub

Public Sub HandleVendorUnitPriceMonitorChange(ByVal wsInfo As Worksheet, ByVal changedRange As Range)
    If wsInfo Is Nothing Then Exit Sub
    If changedRange Is Nothing Then Exit Sub
    If mSyncVendorBlocksInProgress Then Exit Sub
    If Intersect(changedRange, GetVendorUnitPriceMonitorRange(wsInfo)) Is Nothing Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    On Error GoTo ExitHandler
    Application.EnableEvents = False

    ' 工事種別(10行目)・業者名(11行目)が変わった列はヘッダー/結合/罫線/数式を含めた全展開が必要
    Dim structuralCols As Collection
    Set structuralCols = CollectMonitorChangedValueColumns(wsInfo, changedRange, False, False)

    ' 外注比率(29行目)のみが変わった列は、単価数式が比率セルを直接参照しているため
    ' 表示欄(ラベル・値)だけ更新すれば良く、全展開は不要(速度改善)
    Dim ratioOnlyCols As Collection
    Set ratioOnlyCols = CollectOutsourceRatioOnlyChangedValueColumns(wsInfo, changedRange, structuralCols)

    Dim weldingCols As Collection
    Set weldingCols = CollectMonitorChangedValueColumns(wsInfo, changedRange, False, True)

    Dim preferredRatioColumn As Long
    preferredRatioColumn = GetPreferredWeldingRatioColumnFromChange(wsInfo, changedRange)

    Dim col As Variant
    For Each col In structuralCols
        mod_DebugLog.Log "[VendorMaster] HandleVendorUnitPriceMonitorChange: full refresh col=" & CLng(col)
        RefreshVendorUnitPriceForValueColumn wsInfo, CLng(col)
    Next col

    For Each col In ratioOnlyCols
        mod_DebugLog.Log "[VendorMaster] HandleVendorUnitPriceMonitorChange: ratio-only refresh col=" & CLng(col)
        RefreshVendorUnitPriceOutsourceRatioOnlyForValueColumn wsInfo, CLng(col)
    Next col

    If weldingCols.Count > 0 Then
        ' 10行目(工事種別)変更は溶接/軌道列の配置が変わるため全展開する
        If ChangedRangeIncludesVendorWorkTypeRow(wsInfo, changedRange) Then
            mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfo wsInfo, False, preferredRatioColumn
        Else
            mod_WeldingUnitPrice.ApplyWeldingVendorUnitPricesForBasicInfoColumns _
                wsInfo, weldingCols, preferredRatioColumn
        End If
    End If

ExitHandler:
    Application.EnableEvents = prevEvents
End Sub

Public Sub RefreshAllConstructionUnitPriceSheetDataDecorations(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Dim wsUnitPrice As Worksheet
    For Each wsUnitPrice In wsInfo.Parent.worksheets
        If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(wsUnitPrice) And mod_MaterialPriceImport.IsCurrentImportBatchUnitPriceSheet(wsUnitPrice) Then
            RefreshConstructionUnitPriceSheetDataDecorations wsUnitPrice, wsInfo
        End If
    Next wsUnitPrice
End Sub

Public Sub RefreshAllVendorUnitPricesForBasicInfo(Optional ByVal wsInfo As Worksheet, Optional ByVal deferCalculation As Boolean = False)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    EnsureApplicationCalculationAutomatic

    Dim vendorCount As Long
    vendorCount = mod_VendorBlockLayout.GetVendorBlockCount(wsInfo)
    SyncVendorUnitPriceBlocksAfterCountChange wsInfo, vendorCount, 0, deferCalculation
End Sub

' F9(施工会社数)変更時: 増減した列だけ工事単価シートへ反映し、全社・全シート再展開を避ける。

Public Sub RefreshConstructionUnitPriceSheetDataDecorations(ByVal wsUnitPrice As Worksheet, _
                                                            ByVal wsInfo As Worksheet, _
                                                            Optional ByVal skipVendorColumnRefresh As Boolean = False)
    If wsUnitPrice Is Nothing Then Exit Sub
    If wsInfo Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = GetVendorUnitPriceLastDataRow(wsUnitPrice)
    If lastRow < VENDOR_UNIT_PRICE_DATA_START_ROW Then Exit Sub

    ApplyConstructionUnitPriceImportedRowDecorationsFast wsUnitPrice, wsInfo, _
        VENDOR_UNIT_PRICE_DATA_START_ROW, lastRow, skipVendorColumnRefresh
End Sub
