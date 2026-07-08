Option Explicit

' 内訳明細シートの明細部転記エンジン。
' 施工指示書(工事)/施工通知書(工事)・施工指示書(溶接)/施工通知書(溶接)の取込済みシートから
' 施工会社(業者マスタA列 業者名)で抽出し、契約線区名_管理室のセクション単位で転記する。
' 工事区分が軌道工事の場合: 工事シート → 2行空け → 溶接シート(軌道手元会社で抽出)。
' 工事区分が溶接工事の場合: 溶接シート(溶接会社で抽出)のみ。
' 改修履歴: CHANGELOG.md 参照

Private Const SRC_FIELD_SEIRI As Long = 1
Private Const SRC_FIELD_TYPE As Long = 2
Private Const SRC_FIELD_DAYNIGHT As Long = 3
Private Const SRC_FIELD_UNIT As Long = 4
Private Const SRC_FIELD_QTY As Long = 5
Private Const SRC_FIELD_PRICE As Long = 6

Private Const FMT_GROUP_GENERAL As Long = 0
Private Const FMT_GROUP_INTEGER As Long = 1
Private Const FMT_GROUP_DECIMAL As Long = 2

' 内訳明細シートの明細部(11行目～小計行の直前)へ転記する
Public Sub ApplyBreakdownDetails(ByVal wsBreakdown As Worksheet, _
                                 ByVal vendorName As String, _
                                 ByVal officialName As String, _
                                 ByVal branchName As String, _
                                 ByVal workTypeText As String)
    If wsBreakdown Is Nothing Then Exit Sub
    If CommonNormalizeText(vendorName) = "" And CommonNormalizeText(officialName) = "" Then Exit Sub

    On Error GoTo ErrorHandler

    ' 業者マスタの名寄せ(正規名/略称のどちらの表記でも一致させる)
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

    ' 抽出: 軌道工事は工事シート(A列 施工業者)+溶接シート(B列 軌道手元会社)、
    '       溶接工事は溶接シート(A列 溶接会社)のみ
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

    Dim worksLineCount As Long
    Dim weldLineCount As Long
    worksLineCount = CountBlockLines(worksSections)
    weldLineCount = CountBlockLines(weldSections)

    Dim totalLines As Long
    totalLines = worksLineCount + weldLineCount
    If worksLineCount > 0 And weldLineCount > 0 Then totalLines = totalLines + 2
    If totalLines = 0 Then
        mod_OrderTpl_Shared.OrderTplLog "no detail rows: " & wsBreakdown.Name & " vendor=" & vendorName
        Exit Sub
    End If

    EnsureDetailCapacity wsBreakdown, subtotalRow, totalLines

    ' 転記値の組み立て
    Dim valuesAF() As Variant
    Dim valuesNO() As Variant
    ReDim valuesAF(1 To totalLines, 1 To 6)
    ReDim valuesNO(1 To totalLines, 1 To 2)

    Dim headerLineRows As Collection
    Dim integerLineRows As Collection
    Dim decimalLineRows As Collection
    Set headerLineRows = New Collection
    Set integerLineRows = New Collection
    Set decimalLineRows = New Collection

    Dim lineCursor As Long
    lineCursor = 0
    WriteBlockLines worksSections, False, lineCursor, valuesAF, valuesNO, _
                    headerLineRows, integerLineRows, decimalLineRows
    If worksLineCount > 0 And weldLineCount > 0 Then lineCursor = lineCursor + 2
    WriteBlockLines weldSections, True, lineCursor, valuesAF, valuesNO, _
                    headerLineRows, integerLineRows, decimalLineRows

    ' 一括書き込み
    Dim startRow As Long
    Dim endRow As Long
    startRow = ORDER_TPL_DETAIL_START_ROW
    endRow = startRow + totalLines - 1

    wsBreakdown.Range(wsBreakdown.Cells(startRow, 1), wsBreakdown.Cells(endRow, 6)).value = valuesAF
    wsBreakdown.Range(wsBreakdown.Cells(startRow, 14), wsBreakdown.Cells(endRow, 15)).value = valuesNO

    ApplyDetailFormats wsBreakdown, startRow, endRow, headerLineRows, integerLineRows, decimalLineRows

    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownDetails done: " & wsBreakdown.Name & _
        " rows=" & totalLines & " works=" & worksLineCount & " weld=" & weldLineCount
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownDetails error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' 明細部を初期状態(テンプレート行数)へ戻し、値をクリアする
Private Sub ResetDetailArea(ByVal wsBreakdown As Worksheet, ByRef subtotalRow As Long)
    subtotalRow = FindSubtotalRow(wsBreakdown)
    If subtotalRow = 0 Then Exit Sub

    Dim defaultLastRow As Long
    defaultLastRow = ORDER_TPL_DETAIL_START_ROW + ORDER_TPL_DETAIL_DEFAULT_ROWS - 1

    If subtotalRow - 1 > defaultLastRow Then
        wsBreakdown.Rows((defaultLastRow + 1) & ":" & (subtotalRow - 1)).Delete
        subtotalRow = defaultLastRow + 1
    End If

    Dim clearLastRow As Long
    clearLastRow = subtotalRow - 1

    With wsBreakdown
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 1), .Cells(clearLastRow, 6)).ClearContents
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 14), .Cells(clearLastRow, 15)).ClearContents
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 5), .Cells(clearLastRow, 6)).NumberFormat = "General"
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 14), .Cells(clearLastRow, 15)).NumberFormat = "General"
    End With
End Sub

' A列の「小計」行を探す
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

' 必要行数を確保する(不足分は小計行の直前へ行挿入し、数式・書式を引き継ぐ)
Private Sub EnsureDetailCapacity(ByVal wsBreakdown As Worksheet, _
                                 ByRef subtotalRow As Long, _
                                 ByVal neededLines As Long)
    Dim availableRows As Long
    availableRows = subtotalRow - ORDER_TPL_DETAIL_START_ROW

    If neededLines <= availableRows - 1 Then Exit Sub

    Dim insertCount As Long
    insertCount = neededLines - availableRows + 1

    wsBreakdown.Rows(subtotalRow & ":" & (subtotalRow + insertCount - 1)).Insert _
        Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' 数式列(G:P)をクリーンなテンプレート行(12行目)からコピーして引き継ぐ
    Dim sourceRow As Range
    Set sourceRow = wsBreakdown.Range(wsBreakdown.Cells(ORDER_TPL_DETAIL_START_ROW + 1, 7), _
                                      wsBreakdown.Cells(ORDER_TPL_DETAIL_START_ROW + 1, 16))
    sourceRow.Copy Destination:=wsBreakdown.Range( _
        wsBreakdown.Cells(subtotalRow, 7), _
        wsBreakdown.Cells(subtotalRow + insertCount - 1, 16))
    Application.CutCopyMode = False

    subtotalRow = subtotalRow + insertCount
End Sub

' 取込済みシートから施工会社で抽出し、契約線区名_管理室のセクション一覧を作る
' 戻り値: Collection of Array(セクション見出し, 行Collection)。行 = Array(整理番号, 工事種類, 昼夜別, 単位, 数量, JR単価)
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

    Dim sectionKeys As Collection
    Set sectionKeys = New Collection

    Dim sectionMap As Object
    Set sectionMap = CreateObject("Scripting.Dictionary")
    sectionMap.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To UBound(sourceValues, 1)
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

            sectionRows.Add Array( _
                sourceValues(i, COL_SEIRI + columnOffset), _
                sourceValues(i, COL_TYPE + columnOffset), _
                sourceValues(i, COL_DAYNIGHT + columnOffset), _
                sourceValues(i, COL_UNIT + columnOffset), _
                sourceValues(i, COL_QTY + columnOffset), _
                sourceValues(i, COL_JR_PRICE + columnOffset))
        End If
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

' セクション見出し: 契約線区名(接尾辞除去)[_レール溶接]_管理室名
Private Function BuildSectionLabel(ByVal lineText As String, _
                                   ByVal managerRoomText As String, _
                                   ByVal isWeldingSource As Boolean) As String
    Dim label As String
    label = mod_OrderTpl_Shared.OrderTplStripLineSuffix(lineText, isWeldingSource)
    If isWeldingSource Then
        label = label & "_" & mod_OrderTpl_Shared.OrderTplRailWeldingLabelText()
    End If
    BuildSectionLabel = label & "_" & CommonNormalizeText(managerRoomText)
End Function

' セクション内ソート: 第1キー 昼夜別(昼が先)、第2キー 整理番号(昇順)
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
        If total > 0 Then total = total + 1   ' セクション間の空白行
        total = total + 1                     ' セクション見出し行
        total = total + sectionItem(1).Count  ' データ行
    Next sectionItem
    CountBlockLines = total
End Function

' セクション一覧を転記用配列へ展開する
Private Sub WriteBlockLines(ByVal sections As Collection, _
                            ByVal isWeldingSource As Boolean, _
                            ByRef lineCursor As Long, _
                            ByRef valuesAF() As Variant, _
                            ByRef valuesNO() As Variant, _
                            ByVal headerLineRows As Collection, _
                            ByVal integerLineRows As Collection, _
                            ByVal decimalLineRows As Collection)
    If sections Is Nothing Then Exit Sub

    Dim isFirstSection As Boolean
    isFirstSection = True

    Dim sectionItem As Variant
    For Each sectionItem In sections
        If Not isFirstSection Then lineCursor = lineCursor + 1   ' セクション間の空白行
        isFirstSection = False

        lineCursor = lineCursor + 1
        valuesAF(lineCursor, 1) = CStr(sectionItem(0))
        headerLineRows.Add lineCursor

        Dim dataRow As Variant
        For Each dataRow In sectionItem(1)
            lineCursor = lineCursor + 1
            valuesAF(lineCursor, 1) = dataRow(SRC_FIELD_SEIRI - 1)
            valuesAF(lineCursor, 2) = dataRow(SRC_FIELD_TYPE - 1)
            valuesAF(lineCursor, 3) = dataRow(SRC_FIELD_DAYNIGHT - 1)
            valuesAF(lineCursor, 4) = dataRow(SRC_FIELD_UNIT - 1)
            valuesAF(lineCursor, 5) = NumericOrValue(dataRow(SRC_FIELD_QTY - 1))
            valuesAF(lineCursor, 6) = NumericOrValue(dataRow(SRC_FIELD_PRICE - 1))
            valuesNO(lineCursor, 1) = valuesAF(lineCursor, 5)
            valuesNO(lineCursor, 2) = valuesAF(lineCursor, 6)

            Select Case QuantityFormatGroup(CommonNzText(dataRow(SRC_FIELD_UNIT - 1)), isWeldingSource)
                Case FMT_GROUP_INTEGER
                    integerLineRows.Add lineCursor
                Case FMT_GROUP_DECIMAL
                    decimalLineRows.Add lineCursor
            End Select
        Next dataRow
    Next sectionItem
End Sub

' 数量の表示形式: 溶接は桁切りなし(General)、工事は単位に応じて整数(0)/小数3桁(0.000)
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

' 書式適用: フォント・配置・縮小表示・表示形式
Private Sub ApplyDetailFormats(ByVal wsBreakdown As Worksheet, _
                               ByVal startRow As Long, _
                               ByVal endRow As Long, _
                               ByVal headerLineRows As Collection, _
                               ByVal integerLineRows As Collection, _
                               ByVal decimalLineRows As Collection)
    With wsBreakdown
        ' 入力範囲全体を BIZ UDゴシックに
        .Range(.Cells(startRow, 1), .Cells(endRow, 16)).Font.Name = BASIC_INFO_REF_FONT_NAME

        ' A列: 整理番号は左右中央揃え
        .Range(.Cells(startRow, 1), .Cells(endRow, 1)).HorizontalAlignment = xlCenter

        ' B列: 工事種類は左詰め・縮小して全体を表示
        With .Range(.Cells(startRow, 2), .Cells(endRow, 2))
            .HorizontalAlignment = xlLeft
            .WrapText = False
            .ShrinkToFit = True
        End With

        ' C列: 昼夜別は左右中央揃え
        .Range(.Cells(startRow, 3), .Cells(endRow, 3)).HorizontalAlignment = xlCenter

        ' F列/O列: JR単価は小数点なし
        .Range(.Cells(startRow, 6), .Cells(endRow, 6)).NumberFormat = "0"
        .Range(.Cells(startRow, 15), .Cells(endRow, 15)).NumberFormat = "0"

        ' G列/P列: 金額は小数が出る場合のみ小数表示(General)
        .Range(.Cells(startRow, 7), .Cells(endRow, 7)).NumberFormat = "General"
        .Range(.Cells(startRow, 16), .Cells(endRow, 16)).NumberFormat = "General"
    End With

    ' セクション見出し行: 左詰め
    Dim lineIndex As Variant
    For Each lineIndex In headerLineRows
        wsBreakdown.Cells(startRow + CLng(lineIndex) - 1, 1).HorizontalAlignment = xlLeft
    Next lineIndex

    ' E列/N列: 数量の表示形式(整数/小数3桁)
    ApplyNumberFormatToLineRows wsBreakdown, startRow, 5, integerLineRows, "0"
    ApplyNumberFormatToLineRows wsBreakdown, startRow, 14, integerLineRows, "0"
    ApplyNumberFormatToLineRows wsBreakdown, startRow, 5, decimalLineRows, "0.000"
    ApplyNumberFormatToLineRows wsBreakdown, startRow, 14, decimalLineRows, "0.000"
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

Private Sub AddTargetKey(ByVal targetKeys As Object, _
                         ByVal nameText As String, _
                         ByVal aliasMap As Object)
    Dim keyText As String
    keyText = mod_Construction_BasicTotals.ResolveVendorCanonicalKey(nameText, aliasMap)
    If keyText = "" Then Exit Sub
    If Not targetKeys.Exists(keyText) Then targetKeys.Add keyText, True
End Sub
