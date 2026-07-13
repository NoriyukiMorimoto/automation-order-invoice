Option Explicit

' 注文書テンプレート各シートへの基本情報ヘッダー転記。
' 生成時に加え、基本情報シートの転記元セル変更時にも再転記される(ライブ反映)。
' 転記対象シートの追加は ApplyVendorSheetHeaders のディスパッチへ実装を差し込む。
' 改修履歴: CHANGELOG.md 参照

' 基本情報のヘッダー転記元セル(全社共通分)。ブロック列(16/27行目)は動的に組み立てる
Private Const HEADER_SOURCE_COMMON_CELLS As String = "B6,C6,C2,C9,C10,C13,C15:C16,F6"
Private Const HEADER_DATE_FONT_SIZE As Double = 14#
' 受注者用シート転記で参照する施工会社ブロック行(基本情報)
Private Const CONTRACTOR_CONTRACT_AMOUNT_ROW As Long = 33
Private Const CONTRACTOR_CONSUMPTION_TAX_ROW As Long = 34
Private Const CONTRACTOR_CONTRACT_TOTAL_ROW As Long = 35

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

    ' D6: 工事名称(基本情報C10、結合セルのため中央揃え)
    WriteHeaderValue wsBreakdown.Range("D6"), wsInfo.Range("C10").value, True

    ' L5/L6: 工期 自/至(基本情報C15/C16、和暦表示)
    WriteHeaderDate wsBreakdown.Range("L5"), wsInfo.Range("C15").value
    WriteHeaderDate wsBreakdown.Range("L6"), wsInfo.Range("C16").value
    wsBreakdown.Range("L5:L6").Font.Size = HEADER_DATE_FONT_SIZE

    ' O2: 作成日(基本情報C2、和暦表示)
    WriteHeaderDate wsBreakdown.Range("O2"), wsInfo.Range("C2").value
    wsBreakdown.Range("O2").MergeArea.Font.Size = HEADER_DATE_FONT_SIZE

    ' O3: 所長名(基本情報F6)
    WriteHeaderValue wsBreakdown.Range("O3"), wsInfo.Range("F6").value, True

    ' P5: 外注会社名(施工会社ブロック11行目)
    WriteHeaderValue wsBreakdown.Range("P5"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' P6: 業者コード(施工会社ブロック16行目)
    WriteHeaderValue wsBreakdown.Range("P6"), _
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
' 受注者用シートへの転記(注文書テンプレート 受注者用シートの様式に合わせる)
'   S1:注文番号(27) Q2:作成日C2(西暦) E9:業者コード(16) A13:会社名(11)
'   E20:工事名C10 E22:都道府県C13 G24:工期自C15(西暦) G26:工期至C16(西暦) …中央
'   Q22:税込(35) Q23:税抜(33) Q24:消費税(34) …右詰
'   M10:発注者住所 M11:大鉄工業株式会社+基幹出張所 M12:役職氏名(出張所長リスト参照)
Private Sub ApplyContractorHeader(ByVal wsInfo As Worksheet, _
                                  ByVal wsTarget As Worksheet, _
                                  ByVal vendorIndex As Long)
    If wsInfo Is Nothing Then Exit Sub
    If wsTarget Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)

    ' S1: 注文番号(施工会社ブロック27行), 中央
    WriteHeaderValue wsTarget.Range("S1"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_ORDER_NO_ROW, valueColumn).value, True

    ' Q2: 作成日(基本情報C2, 西暦), 中央
    WriteHeaderDateGregorian wsTarget.Range("Q2"), wsInfo.Range("C2").value

    ' E9: 業者コード(施工会社ブロック16行), 中央
    WriteHeaderValue wsTarget.Range("E9"), _
                     wsInfo.Cells(ORDER_TPL_BLOCK_VENDOR_CODE_ROW, valueColumn).value, True

    ' A13: 会社名(施工会社ブロック11行), 中央
    WriteHeaderValue wsTarget.Range("A13"), _
                     wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value, True

    ' E20: 工事名(基本情報C10), 中央
    WriteHeaderValue wsTarget.Range("E20"), wsInfo.Range("C10").value, True

    ' E22: 都道府県(基本情報C13), 中央
    WriteHeaderValue wsTarget.Range("E22"), wsInfo.Range("C13").value, True

    ' G24: 工期 自(基本情報C15, 西暦), 中央
    WriteHeaderDateGregorian wsTarget.Range("G24"), wsInfo.Range("C15").value

    ' G26: 工期 至(基本情報C16, 西暦), 中央
    WriteHeaderDateGregorian wsTarget.Range("G26"), wsInfo.Range("C16").value

    ' Q22: 税込金額(施工会社ブロック35行), 右詰
    WriteHeaderValueRight wsTarget.Range("Q22"), _
                          wsInfo.Cells(CONTRACTOR_CONTRACT_TOTAL_ROW, valueColumn).value
    ' Q23: 契約金額 税抜(施工会社ブロック33行), 右詰
    WriteHeaderValueRight wsTarget.Range("Q23"), _
                          wsInfo.Cells(CONTRACTOR_CONTRACT_AMOUNT_ROW, valueColumn).value
    ' Q24: 消費税(施工会社ブロック34行), 右詰
    WriteHeaderValueRight wsTarget.Range("Q24"), _
                          wsInfo.Cells(CONTRACTOR_CONSUMPTION_TAX_ROW, valueColumn).value

    ' M10/M11/M12: 出張所長リスト参照(発注者 住所/名称/役職氏名)
    ApplyOfficeChiefBlock wsInfo, wsTarget

    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader done: " & wsTarget.Name
    Exit Sub

ErrorHandler:
    mod_OrderTpl_Shared.OrderTplLog "ApplyContractorHeader error: " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

' M10:住所 / M11:大鉄工業株式会社+全角空白+基幹出張所 / M12:(条件で)役職氏名 を転記
Private Sub ApplyOfficeChiefBlock(ByVal wsInfo As Worksheet, ByVal wsTarget As Worksheet)
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

    ' M10: 住所
    WriteHeaderValue wsTarget.Range("M10"), addr, False

    ' M11: 大鉄工業株式会社 + 全角空白 + 基幹出張所
    WriteHeaderValue wsTarget.Range("M11"), CommonCompanyNameText() & fw & coreOffice, False

    ' M12: 出張所=基幹出張所なら「役職 氏名」、不一致なら「出張所 役職 氏名」
    Dim m12 As String
    If StrComp(CommonNormalizeText(matchedOffice), CommonNormalizeText(coreOffice), vbTextCompare) = 0 Then
        m12 = title & fw & chiefName
    Else
        m12 = matchedOffice & fw & title & fw & chiefName
    End If
    WriteHeaderValue wsTarget.Range("M12"), m12, False
End Sub

' 右詰の値転記(BizUDゴシック適用)
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
    target.MergeArea.Font.Name = BASIC_INFO_REF_FONT_NAME
    target.MergeArea.HorizontalAlignment = xlRight
    target.MergeArea.VerticalAlignment = xlCenter
End Sub

' 西暦日付の転記(yyyy年m月d日・中央)
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
