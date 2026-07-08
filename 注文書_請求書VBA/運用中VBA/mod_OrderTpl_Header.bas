Option Explicit

' 内訳明細シートのヘッダー部(部店コード・注文番号・工事番号・工事名称・工期・作成日・所長・外注会社)転記。
' 入力文字はすべて BIZ UDゴシック(BASIC_INFO_REF_FONT_NAME)で入力する。
' 改修履歴: CHANGELOG.md 参照

' 内訳明細ヘッダー部へ基本情報シートの内容を転記する
Public Sub ApplyBreakdownHeader(ByVal wsInfo As Worksheet, _
                                ByVal wsBreakdown As Worksheet, _
                                ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsBreakdown Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' C2: 部店コード(出張所別_単価適用線区の単価適用線区シート B/C列照合 → G列)
    Dim branchOfficeCode As String
    branchOfficeCode = mod_OrderTpl_Shared.OrderTplResolveBranchOfficeCode( _
        CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value), _
        CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    WriteHeaderText wsBreakdown.Range("C2"), branchOfficeCode, False

    ' C3: 注文番号(施工会社ブロック27行目)
    WriteHeaderValue wsBreakdown.Range("C3"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, False

    ' B6: 工事番号(基本情報C9)
    WriteHeaderValue wsBreakdown.Range("B6"), wsInfo.Range("C9").value, False

    ' C6: 工事名称(基本情報C10、結合セルのため中央揃え)
    WriteHeaderValue wsBreakdown.Range("C6"), wsInfo.Range("C10").value, True

    ' K5/K6: 工期 自/至(基本情報C15/C16、和暦表示)
    WriteHeaderDate wsBreakdown.Range("K5"), wsInfo.Range("C15").value
    WriteHeaderDate wsBreakdown.Range("K6"), wsInfo.Range("C16").value

    ' N2: 作成日(基本情報C2、和暦表示)
    WriteHeaderDate wsBreakdown.Range("N2"), wsInfo.Range("C2").value

    ' N3: 所長名(基本情報F6)
    WriteHeaderValue wsBreakdown.Range("N3"), wsInfo.Range("F6").value, True

    ' O5: 外注会社名(施工会社ブロック11行目)
    WriteHeaderValue wsBreakdown.Range("O5"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' O6: 業者コード(施工会社ブロック16行目)
    WriteHeaderValue wsBreakdown.Range("O6"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' 複数ページ印刷用にヘッダー行(7:10行目)をタイトル行に設定する
    On Error Resume Next
    wsBreakdown.PageSetup.PrintTitleRows = ORDER_TPL_PRINT_TITLE_ROWS
    On Error GoTo ErrorHandler

    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownHeader done: " & wsBreakdown.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyBreakdownHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' 結合セル対応の値転記(フォント適用、必要に応じて上下左右中央揃え)
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

' 文字列としての転記(部店コード等、日付誤変換を防ぐ)
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

' 日付の転記(和暦表示形式、結合セルのため中央揃え)
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
