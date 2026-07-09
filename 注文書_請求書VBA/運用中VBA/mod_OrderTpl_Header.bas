Option Explicit

' 注文書テンプレート各シートへの基本情報ヘッダー転記。
' 生成時に加え、基本情報シートの転記元セル変更時にも再転記される(ライブ反映)。
' 転記対象シートの追加は ApplyVendorSheetHeaders のディスパッチへ実装を差し込む。
' 改修履歴: CHANGELOG.md 参照

' 基本情報のヘッダー転記元セル(全社共通分)。ブロック列(16/27行目)は動的に組み立てる
Private Const HEADER_SOURCE_COMMON_CELLS As String = "B6,C6,C2,C9,C10,C15:C16,F6"

' 指定ブロックの施工会社に対応するテンプレート5シートへヘッダーを転記する(ディスパッチャ)
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
End Sub

' 全確定会社のテンプレートシートへヘッダーを再転記する
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
            End If
        End If
    Next vendorIndex
    Exit Sub

Quiet:
    mod_OrderTpl_Shared.OrderTplLog "RefreshAllVendorSheetHeaders error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' Sheet1(基本情報)のWorksheet_Changeから呼ばれる入口。
' ヘッダー転記元セル(B6/C6/C2/C9/C10/C15:C16/F6、各ブロックの16/27行目)の変更を各社シートへ反映する
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

' ヘッダー転記元セルの監視範囲を返す公開ラッパー(Sheet1の変更ゲート構築用)
Public Function GetBasicInfoHeaderSourceMonitorRange(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function
    Set GetBasicInfoHeaderSourceMonitorRange = BuildHeaderSourceRange(wsInfo)
End Function

' ヘッダー転記元セルの監視範囲(共通セル + 各ブロックの業者コード16行目/注文番号27行目)
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
                           wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn), _
                           wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn))
    Next vendorIndex

    Set BuildHeaderSourceRange = result
End Function

' シート見出し(タブ)色: 工事区分(基本情報10行目)セルの塗りつぶし色を適用する
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

' 受注者用シートへの転記(転記仕様が確定したらここへ実装する。実装後は自動でライブ反映される)
Private Sub ApplyContractorHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    ' 転記仕様 未指定(コピーのみ)
End Sub

' 注文請書シートへの転記(転記仕様が確定したらここへ実装する)
Private Sub ApplyAcceptanceHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    ' 転記仕様 未指定(コピーのみ)
End Sub

' 支店控シートへの転記(転記仕様が確定したらここへ実装する)
Private Sub ApplyBranchCopyHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    ' 転記仕様 未指定(コピーのみ)
End Sub

' 別紙Ⅲシートへの転記(転記仕様が確定したらここへ実装する)
Private Sub ApplyAttachment3Header(ByVal wsInfo As Worksheet, _
                                   ByVal wsTarget As Worksheet, _
                                   ByVal vendorIndex As Long)
    ' 転記仕様 未指定(コピーのみ)
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
