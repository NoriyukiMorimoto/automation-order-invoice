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
Private Const SUMMARY_EXTRA_ROWS As Long = 5   ' 値引/計/消費税/合計/罫線用空白
Private Const SUMMARY_NUMBER_FORMAT As String = "#,##0"
Private Const DETAIL_AMOUNT_NUMBER_FORMAT As String = "#,##0;-#,##0;"            ' 桁区切り・ゼロ非表示
Private Const DETAIL_QTY_DECIMAL_NUMBER_FORMAT As String = "#,##0.000;-#,##0.000;" ' 小数3桁・ゼロ非表示
Private Const DETAIL_FONT_SIZE As Double = 10#

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

    PositionSubtotalRow wsBreakdown, subtotalRow, totalLines

    ' 転記値の組み立て
    Dim valuesAF() As Variant
    Dim valuesNO() As Variant
    ReDim valuesAF(1 To totalLines, 1 To 7)
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

    wsBreakdown.Range(wsBreakdown.Cells(startRow, 1), wsBreakdown.Cells(endRow, 7)).value = valuesAF
    wsBreakdown.Range(wsBreakdown.Cells(startRow, 15), wsBreakdown.Cells(endRow, 16)).value = valuesNO

    ' 金額列(H/K/N/Q)へ数式(数量×単価)を設定する(小計行の直前まで)
    Dim formulaColumn As Variant
    For Each formulaColumn In Array("H", "K", "N", "Q")
        wsBreakdown.Range(formulaColumn & startRow & ":" & formulaColumn & (subtotalRow - 1)).FormulaR1C1 = _
            "=RC[-2]*RC[-1]"
    Next formulaColumn

    ApplyDetailFormats wsBreakdown, startRow, endRow, headerLineRows, integerLineRows, decimalLineRows

    ' B:C列の結合(11行目～小計行の直前。セクション見出し行はA:D結合のため除く)
    MergeDetailNameColumns wsBreakdown, startRow, subtotalRow - 1, headerLineRows

    BuildSummaryBlock wsBreakdown, subtotalRow, totalLines

    ' 11行目以降(集計ブロック含む)のフォントサイズを10ポイントへ統一する
    wsBreakdown.Range(wsBreakdown.Cells(startRow, 1), _
                      wsBreakdown.Cells(subtotalRow + SUMMARY_EXTRA_ROWS, 17)).Font.Size = DETAIL_FONT_SIZE

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

    ' 以前に生成した集計ブロック(値引/計/消費税/合計/罫線用空白行)を削除する
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

    ' 前回転記のセクション見出し(A:D結合)を解除してから値をクリアする
    mod_VendorBlockLayout.SafeUnmergeRange wsBreakdown.Range( _
        wsBreakdown.Cells(ORDER_TPL_DETAIL_START_ROW, 1), wsBreakdown.Cells(clearLastRow, 4))

    With wsBreakdown
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 1), .Cells(clearLastRow, 7)).ClearContents
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 15), .Cells(clearLastRow, 16)).ClearContents
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 6), .Cells(clearLastRow, 7)).NumberFormat = "General"
        .Range(.Cells(ORDER_TPL_DETAIL_START_ROW, 15), .Cells(clearLastRow, 16)).NumberFormat = "General"
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

' 小計行を「最終データ行から2行空けた位置」へ移動する(不足は行挿入、余剰は行削除)
Private Sub PositionSubtotalRow(ByVal wsBreakdown As Worksheet, _
                                ByRef subtotalRow As Long, _
                                ByVal totalLines As Long)
    Dim desiredRow As Long
    desiredRow = ORDER_TPL_DETAIL_START_ROW + totalLines + 2

    If desiredRow > subtotalRow Then
        Dim insertCount As Long
        insertCount = desiredRow - subtotalRow

        ' 書式は上方行から継承。金額数式(H/K/N/Q)は転記後に全行へ設定する
        wsBreakdown.Rows(subtotalRow & ":" & (subtotalRow + insertCount - 1)).Insert _
            Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ElseIf desiredRow < subtotalRow Then
        wsBreakdown.Rows(desiredRow & ":" & (subtotalRow - 1)).Delete
    End If

    subtotalRow = desiredRow
End Sub

' 集計ブロックを構築する: 小計/値引/計/消費税/合計 + 罫線用空白行
Private Sub BuildSummaryBlock(ByVal wsBreakdown As Worksheet, _
                              ByVal subtotalRow As Long, _
                              ByVal totalLines As Long)
    Dim lastDataRow As Long
    lastDataRow = ORDER_TPL_DETAIL_START_ROW + totalLines - 1

    ' 小計行の下へ4行(値引/計/消費税/合計)+罫線用空白行を挿入する
    wsBreakdown.Rows((subtotalRow + 1) & ":" & (subtotalRow + SUMMARY_EXTRA_ROWS)).Insert _
        Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' 消費税率(基本情報B34「消費税(10%)：」のカッコ内)を取得する
    Dim taxRateText As String
    taxRateText = Trim$(Str$(mod_Construction_BasicTotals.ResolveBasicInfoTaxRate( _
        CommonGetBasicInfoWorksheet())))

    ' ラベル(A:C結合・中央揃え)
    WriteSummaryLabel wsBreakdown, subtotalRow, mod_OrderTpl_Shared.OrderTplSubtotalLabelText()
    WriteSummaryLabel wsBreakdown, subtotalRow + 1, DiscountLabelText()
    WriteSummaryLabel wsBreakdown, subtotalRow + 2, NetTotalLabelText()
    WriteSummaryLabel wsBreakdown, subtotalRow + 3, TaxLabelText()
    WriteSummaryLabel wsBreakdown, subtotalRow + 4, GrandTotalLabelText()

    ' H/K/N/Q 列: 小計=SUM(11:最終データ行)、値引=-MOD(小計,1000)、計=小計+値引、
    ' 消費税=ROUNDDOWN(計*税率,0)、合計=計+消費税
    Dim summaryColumns As Variant
    summaryColumns = Array("H", "K", "N", "Q")

    Dim i As Long
    For i = LBound(summaryColumns) To UBound(summaryColumns)
        Dim colLetter As String
        colLetter = CStr(summaryColumns(i))

        wsBreakdown.Range(colLetter & subtotalRow).Formula = _
            "=SUM(" & colLetter & ORDER_TPL_DETAIL_START_ROW & ":" & colLetter & lastDataRow & ")"
        wsBreakdown.Range(colLetter & (subtotalRow + 1)).Formula = _
            "=-MOD(" & colLetter & subtotalRow & ",1000)"
        wsBreakdown.Range(colLetter & (subtotalRow + 2)).Formula = _
            "=" & colLetter & subtotalRow & "+" & colLetter & (subtotalRow + 1)
        wsBreakdown.Range(colLetter & (subtotalRow + 3)).Formula = _
            "=ROUNDDOWN(" & colLetter & (subtotalRow + 2) & "*" & taxRateText & ",0)"
        wsBreakdown.Range(colLetter & (subtotalRow + 4)).Formula = _
            "=" & colLetter & (subtotalRow + 2) & "+" & colLetter & (subtotalRow + 3)

        wsBreakdown.Range(colLetter & subtotalRow & ":" & colLetter & (subtotalRow + 4)).NumberFormat = _
            SUMMARY_NUMBER_FORMAT
    Next i

    ' フォント
    wsBreakdown.Range(wsBreakdown.Cells(subtotalRow, 1), _
                      wsBreakdown.Cells(subtotalRow + 4, 17)).Font.Name = BASIC_INFO_REF_FONT_NAME

    ApplySummaryBorders wsBreakdown, subtotalRow
End Sub

' ラベルセル(A:C結合・上下左右中央揃え)への書き込み
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

' 集計ブロックの罫線: 小計行の上罫線=二重線、ブロック内横罫線=細線、
' 罫線用空白行の下罫線=中線(A:P)。縦罫線は挿入時に上方セルから継承済み
Private Sub ApplySummaryBorders(ByVal wsBreakdown As Worksheet, _
                                ByVal subtotalRow As Long)
    ' 明細部(11行目～小計行の直前)の横罫線をすべて細線で引き直す(転記の最後に実行)
    With wsBreakdown.Range(wsBreakdown.Cells(ORDER_TPL_DETAIL_START_ROW, 1), _
                           wsBreakdown.Cells(subtotalRow - 1, 17))
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With

    Dim blockRange As Range
    Set blockRange = wsBreakdown.Range(wsBreakdown.Cells(subtotalRow, 1), _
                                       wsBreakdown.Cells(subtotalRow + SUMMARY_EXTRA_ROWS, 17))

    With blockRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    ' 縦罫線の再設定: 集計ブロック内の既存縦罫線をいったん消去する
    blockRange.Borders(xlInsideVertical).LineStyle = xlNone

    ' D/E列間: 11行目～小計行の直前まで。小計行以下は消去
    With wsBreakdown.Range(wsBreakdown.Cells(ORDER_TPL_DETAIL_START_ROW, 4), _
                           wsBreakdown.Cells(subtotalRow - 1, 4)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    wsBreakdown.Range(wsBreakdown.Cells(subtotalRow, 4), _
                      wsBreakdown.Cells(subtotalRow + SUMMARY_EXTRA_ROWS, 4)).Borders(xlEdgeRight).LineStyle = xlNone

    ' G/H・J/K・M/N・P/Q列間: 11行目～合計行の1行下まで
    Dim verticalRightColumns As Variant
    verticalRightColumns = Array(7, 10, 13, 16)

    Dim c As Variant
    For Each c In verticalRightColumns
        With wsBreakdown.Range(wsBreakdown.Cells(ORDER_TPL_DETAIL_START_ROW, CLng(c)), _
                               wsBreakdown.Cells(subtotalRow + SUMMARY_EXTRA_ROWS, CLng(c))).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    Next c

    With wsBreakdown.Range(wsBreakdown.Cells(subtotalRow, 1), _
                           wsBreakdown.Cells(subtotalRow, 17)).Borders(xlEdgeTop)
        .LineStyle = xlDouble
    End With

    With wsBreakdown.Range(wsBreakdown.Cells(subtotalRow + SUMMARY_EXTRA_ROWS, 1), _
                           wsBreakdown.Cells(subtotalRow + SUMMARY_EXTRA_ROWS, 17)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
End Sub

' A列セルの表示文字列が指定ラベルと一致するか
Private Function IsSummaryLabelRow(ByVal wsBreakdown As Worksheet, _
                                   ByVal rowIndex As Long, _
                                   ByVal labelText As String) As Boolean
    IsSummaryLabelRow = (StrComp( _
        CommonRemoveAllSpaces(CommonNormalizeText(CommonNzText(wsBreakdown.Cells(rowIndex, 1).value))), _
        labelText, vbTextCompare) = 0)
End Function

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

    ' 施工会社別単価列(ヘッダー「会社名+単価」)を名寄せで特定する
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

            ' 単価は施工会社別単価列(会社名+単価)から取得する(未展開・未割当時は空欄)
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

    ' 施行通知書には管理室名が無いため、空の場合は「_」を付けない
    Dim managerRoomName As String
    managerRoomName = CommonNormalizeText(managerRoomText)
    If Len(managerRoomName) > 0 Then label = label & "_" & managerRoomName

    BuildSectionLabel = label
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
            valuesAF(lineCursor, 4) = dataRow(SRC_FIELD_DAYNIGHT - 1)
            valuesAF(lineCursor, 5) = dataRow(SRC_FIELD_UNIT - 1)
            valuesAF(lineCursor, 6) = NumericOrValue(dataRow(SRC_FIELD_QTY - 1))
            valuesAF(lineCursor, 7) = NumericOrValue(dataRow(SRC_FIELD_PRICE - 1))
            valuesNO(lineCursor, 1) = valuesAF(lineCursor, 6)
            valuesNO(lineCursor, 2) = valuesAF(lineCursor, 7)

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
        .Range(.Cells(startRow, 1), .Cells(endRow, 17)).Font.Name = BASIC_INFO_REF_FONT_NAME

        ' A列: 整理番号は左右中央揃え
        .Range(.Cells(startRow, 1), .Cells(endRow, 1)).HorizontalAlignment = xlCenter

        ' B列: 工事種類は左詰め・縮小して全体を表示
        With .Range(.Cells(startRow, 2), .Cells(endRow, 2))
            .HorizontalAlignment = xlLeft
            .WrapText = False
            .ShrinkToFit = True
        End With

        ' D列: 昼夜別は左右中央揃え
        .Range(.Cells(startRow, 4), .Cells(endRow, 4)).HorizontalAlignment = xlCenter

        ' F～Q列: 桁区切り・ゼロ値非表示。数量列(F/I/L/O)の既定=整数、小数3桁の行は後で上書き
        .Range(.Cells(startRow, 6), .Cells(endRow, 17)).NumberFormat = DETAIL_AMOUNT_NUMBER_FORMAT
    End With

    ' セクション見出し行(契約線区名・管理室名): A:D結合・左詰め・縮小してセル内に収める
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

    ' 数量列(F/I/L/O): 単位が小数3桁対象(m/m3/M/㎡/t)の行を小数3桁表示へ上書き
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

' 取込元シートの1行目から「会社名+単価」ヘッダー列を名寄せで探す
Private Function FindVendorUnitPriceColumn(ByVal wsSource As Worksheet, _
                                           ByVal targetKeys As Object, _
                                           ByVal aliasMap As Object) As Long
    If wsSource Is Nothing Then Exit Function
    If targetKeys Is Nothing Then Exit Function

    Dim lastColumn As Long
    lastColumn = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    Dim suffixText As String
    suffixText = mod_OrderTpl_Shared.OrderTplUnitPriceHeaderSuffixText()

    Dim c As Long
    For c = 1 To lastColumn
        Dim headerText As String
        headerText = CommonRemoveAllSpaces(CommonNormalizeText(CommonNzText(wsSource.Cells(1, c).value)))
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

' B:C列を結合する(明細部のみ。セクション見出し行はA:D結合のため除外)
Private Sub MergeDetailNameColumns(ByVal wsBreakdown As Worksheet, _
                                   ByVal firstRow As Long, _
                                   ByVal lastRow As Long, _
                                   ByVal headerLineRows As Collection)
    Dim headerRowMap As Object
    Set headerRowMap = CreateObject("Scripting.Dictionary")

    Dim lineIndex As Variant
    If Not headerLineRows Is Nothing Then
        For Each lineIndex In headerLineRows
            headerRowMap(firstRow + CLng(lineIndex) - 1) = True
        Next lineIndex
    End If

    Dim r As Long
    For r = firstRow To lastRow
        If Not headerRowMap.Exists(r) Then
            With wsBreakdown.Range(wsBreakdown.Cells(r, 2), wsBreakdown.Cells(r, 3))
                If Not .MergeCells Then .Merge
                .WrapText = False
                .ShrinkToFit = True
                .HorizontalAlignment = xlLeft
            End With
        End If
    Next r
End Sub

' 内訳明細シートの前回迄(I列)/今回(L列)への数量入力を監視し、単価列(J/M)を自動入力する。
' ThisWorkbook.Workbook_SheetChange から呼ばれる
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

    Dim hitRange As Range
    Set hitRange = Intersect(target, Union(ws.Columns(9), ws.Columns(12)))
    If hitRange Is Nothing Then Exit Sub

    Dim subtotalRow As Long
    subtotalRow = FindSubtotalRow(ws)
    If subtotalRow = 0 Then Exit Sub

    Set hitRange = Intersect(hitRange, _
        ws.Range(ws.Cells(ORDER_TPL_DETAIL_START_ROW, 9), ws.Cells(subtotalRow - 1, 12)))
    If hitRange Is Nothing Then Exit Sub

    Dim prevEnableEvents As Boolean
    prevEnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim changedCell As Range
    For Each changedCell In hitRange.Cells
        If changedCell.Column = 9 Or changedCell.Column = 12 Then
            Dim priceCell As Range
            Set priceCell = ws.Cells(changedCell.Row, changedCell.Column + 1)

            Dim quantityValue As Variant
            quantityValue = changedCell.value
            If IsNumeric(quantityValue) And Len(Trim$(CStr(quantityValue))) > 0 Then
                ' 単価は当初単価(G列=会社別単価)と同値を入力する
                priceCell.value = ws.Cells(changedCell.Row, 7).value
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

Private Sub AddTargetKey(ByVal targetKeys As Object, _
                         ByVal nameText As String, _
                         ByVal aliasMap As Object)
    Dim keyText As String
    keyText = mod_Construction_BasicTotals.ResolveVendorCanonicalKey(nameText, aliasMap)
    If keyText = "" Then Exit Sub
    If Not targetKeys.Exists(keyText) Then targetKeys.Add keyText, True
End Sub
