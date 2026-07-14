Option Explicit

' 注文書テンプレート取込(施工会社確定時のシート挿入・内訳明細転記)の共通定数・照合ヘルパー。
' 改修履歴: CHANGELOG.md 参照

Public Const ORDER_TPL_DETAIL_START_ROW As Long = 11      ' 内訳明細の明細開始行(セクション見出し行)
Public Const ORDER_TPL_DETAIL_DEFAULT_ROWS As Long = 22   ' テンプレート初期の明細行数(11:32行)
Public Const ORDER_TPL_PRINT_TITLE_ROWS As String = "$7:$10"
Public Const ORDER_TPL_BLOCK_VENDOR_CODE_ROW As Long = 16 ' 基本情報: 業者コード行
Public Const ORDER_TPL_BLOCK_ORDER_NO_ROW As Long = 27    ' 基本情報: 注文番号行
Public Const ORDER_TPL_VENDOR_ADO_NAME_FIELD As Long = 0      ' 業者マスタ A列 業者名
Public Const ORDER_TPL_VENDOR_ADO_OFFICIAL_FIELD As Long = 1  ' 業者マスタ B列 請求者氏名
Public Const ORDER_TPL_VENDOR_ADO_WORK_FIELD As Long = 14     ' 業者マスタ O列 担当工事
Public Const ORDER_TPL_VENDOR_ADO_ALIAS_FIELD As Long = 15    ' 業者マスタ P列 略称

Private mVendorInfoCache As Object       ' 支店|会社名 → Array(業者名, 略称, 担当工事)
Private mBranchOfficeCodeCache As Object ' 支店|出張所 → 部店コード

' テンプレートファイル名
Public Function OrderTplTemplateFileNameText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30C6, &H30F3, &H30D7, &H30EC, &H30FC) & _
                 CommonTextFromChars(&H30C8) & _
                 ".xlsx"
    End If
    OrderTplTemplateFileNameText = cached
End Function

Public Function OrderTplBaseNameBreakdownText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5185, &H8A33, &H660E, &H7D30)
    End If
    OrderTplBaseNameBreakdownText = cached
End Function

Public Function OrderTplBaseNameContractorText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H53D7, &H6CE8, &H8005, &H7528)
    End If
    OrderTplBaseNameContractorText = cached
End Function

Public Function OrderTplBaseNameAcceptanceText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H8ACB, &H66F8)
    End If
    OrderTplBaseNameAcceptanceText = cached
End Function

Public Function OrderTplBaseNameBranchCopyText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H652F, &H5E97, &H63A7)
    End If
    OrderTplBaseNameBranchCopyText = cached
End Function

Public Function OrderTplBaseNameAttachment3Text() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5225, &H7D19, &H2162)
    End If
    OrderTplBaseNameAttachment3Text = cached
End Function

Public Function OrderTplRailWeldingLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H30EC, &H30FC, &H30EB, &H6EB6, &H63A5)
    End If
    OrderTplRailWeldingLabelText = cached
End Function

' 施工会社別単価列ヘッダーの接尾辞
Public Function OrderTplUnitPriceHeaderSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5358, &H4FA1)
    End If
    OrderTplUnitPriceHeaderSuffixText = cached
End Function

Public Function OrderTplDayFirstText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H663C)
    End If
    OrderTplDayFirstText = cached
End Function

Public Function OrderTplSubtotalLabelText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H5C0F, &H8A08)
    End If
    OrderTplSubtotalLabelText = cached
End Function

' 桁切りなし整数の単位一覧
Private Function OrderTplIntegerUnitListText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H53E3) & _
                 "|" & _
                 CommonTextFromChars(&H7A74) & _
                 "|" & _
                 CommonTextFromChars(&H56DE) & _
                 "|" & _
                 CommonTextFromChars(&H500B) & _
                 "|" & _
                 CommonTextFromChars(&H7B87, &H6240) & _
                 "|" & _
                 CommonTextFromChars(&H672C) & _
                 "|" & _
                 CommonTextFromChars(&H7D44) & _
                 "|" & _
                 CommonTextFromChars(&H5F0F) & _
                 "|" & _
                 CommonTextFromChars(&H679A)
    End If
    OrderTplIntegerUnitListText = cached
End Function

' 小数3桁の単位一覧
Private Function OrderTplDecimalUnitListText() As String
    Static cached As String
    If cached = "" Then
        cached = "m|m3|M|m2|m" & _
                 CommonTextFromChars(&HB2) & _
                 "|" & _
                 CommonTextFromChars(&H33A1) & _
                 "|m" & _
                 CommonTextFromChars(&HB3) & _
                 "|M3|t"
    End If
    OrderTplDecimalUnitListText = cached
End Function

' 和暦表示形式
Public Function OrderTplEraDateNumberFormatText() As String
    Static cached As String
    If cached = "" Then
        cached = "[$-411]ggge""" & _
                 CommonTextFromChars(&H5E74) & _
                 """m""" & _
                 CommonTextFromChars(&H6708) & _
                 """d""" & _
                 CommonTextFromChars(&H65E5) & _
                 """"
    End If
    OrderTplEraDateNumberFormatText = cached
End Function

Public Function OrderTplVendorNotFoundMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H696D, &H8005, &H30DE, &H30B9, &H30BF) & _
                 "(" & _
                 CommonTextFromChars(&H5168, &H793E, &H7248) & _
                 ")" & _
                 CommonTextFromChars(&H306B, &H8A72, &H5F53, &H3059, &H308B, &H4F1A, &H793E, &H304C) & _
                 CommonTextFromChars(&H898B, &H3064, &H304B, &H3089, &H306A, &H3044, &H305F, &H3081) & _
                 CommonTextFromChars(&H3001, &H6CE8, &H6587, &H66F8, &H30B7, &H30FC, &H30C8, &H3092) & _
                 CommonTextFromChars(&H4F5C, &H6210, &H3067, &H304D, &H307E, &H305B, &H3093, &H3002)
    End If
    OrderTplVendorNotFoundMessageText = cached
End Function

Public Function OrderTplAliasEmptyMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H696D, &H8005, &H30DE, &H30B9, &H30BF) & _
                 "(" & _
                 CommonTextFromChars(&H5168, &H793E, &H7248) & _
                 ")" & _
                 CommonTextFromChars(&H306E) & _
                 "P" & _
                 CommonTextFromChars(&H5217) & _
                 "(" & _
                 CommonTextFromChars(&H7565, &H79F0) & _
                 ")" & _
                 CommonTextFromChars(&H304C, &H672A, &H5165, &H529B, &H306E, &H305F, &H3081, &H3001) & _
                 CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30B7, &H30FC, &H30C8, &H3092, &H4F5C) & _
                 CommonTextFromChars(&H6210, &H3067, &H304D, &H307E, &H305B, &H3093, &H3002)
    End If
    OrderTplAliasEmptyMessageText = cached
End Function

Public Function OrderTplTemplateNotFoundMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30C6, &H30F3, &H30D7, &H30EC, &H30FC) & _
                 CommonTextFromChars(&H30C8) & _
                 ".xlsx " & _
                 CommonTextFromChars(&H304C, &H898B, &H3064, &H304B, &H308A, &H307E, &H305B, &H3093) & _
                 CommonTextFromChars(&H3002)
    End If
    OrderTplTemplateNotFoundMessageText = cached
End Function

Public Function OrderTplDuplicateAliasMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H540C, &H3058, &H65BD, &H5DE5, &H4F1A, &H793E, &H304C, &H8907) & _
                 CommonTextFromChars(&H6570, &H306E, &H696D, &H8005, &H60C5, &H5831, &H30D6, &H30ED) & _
                 CommonTextFromChars(&H30C3, &H30AF, &H306B, &H5165, &H529B, &H3055, &H308C, &H3066) & _
                 CommonTextFromChars(&H3044, &H308B, &H305F, &H3081, &H3001, &H6CE8, &H6587, &H66F8) & _
                 CommonTextFromChars(&H30B7, &H30FC, &H30C8, &H306E, &H4F5C, &H6210, &H3092, &H4E2D) & _
                 CommonTextFromChars(&H6B62, &H3057, &H307E, &H3059, &H3002)
    End If
    OrderTplDuplicateAliasMessageText = cached
End Function

Public Function OrderTplRefreshDoneMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = CommonTextFromChars(&H6CE8, &H6587, &H66F8, &H30B7, &H30FC, &H30C8, &H3078, &H306E) & _
                 CommonTextFromChars(&H518D, &H8EE2, &H8A18, &H304C, &H5B8C, &H4E86, &H3057, &H307E) & _
                 CommonTextFromChars(&H3057, &H305F, &H3002)
    End If
    OrderTplRefreshDoneMessageText = cached
End Function

Public Sub OrderTplClearCaches()
    Set mVendorInfoCache = Nothing
    Set mBranchOfficeCodeCache = Nothing
End Sub

Public Function OrderTplTemplateSheetBaseNames() As Variant
    OrderTplTemplateSheetBaseNames = Array(OrderTplBaseNameBreakdownText(), _
                                           OrderTplBaseNameContractorText(), _
                                           OrderTplBaseNameAcceptanceText(), _
                                           OrderTplBaseNameBranchCopyText(), _
                                           OrderTplBaseNameAttachment3Text())
End Function

' マスタデータフォルダ内ファイルのパス解決(共有フォルダ → ブック親フォルダ → ブック直下)
Public Function OrderTplMasterDataFilePath(ByVal fileName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")

    Dim candidatePath As String
    If Len(Trim$(userProfilePath)) > 0 Then
        candidatePath = userProfilePath & Chr$(92) & CommonCompanyNameText() & Chr$(92) & _
                        CommonOrderInvoiceDocumentFolderText() & Chr$(92) & _
                        CommonMasterDataFolderText() & Chr$(92) & fileName
        If Len(Dir(candidatePath, vbNormal)) > 0 Then
            OrderTplMasterDataFilePath = candidatePath
            Exit Function
        End If
    End If

    If Len(ThisWorkbook.Path) > 0 Then
        candidatePath = fso.GetParentFolderName(ThisWorkbook.Path) & Chr$(92) & _
                        CommonMasterDataFolderText() & Chr$(92) & fileName
        If Len(Dir(candidatePath, vbNormal)) > 0 Then
            OrderTplMasterDataFilePath = candidatePath
            Exit Function
        End If

        candidatePath = ThisWorkbook.Path & Chr$(92) & CommonMasterDataFolderText() & Chr$(92) & fileName
        If Len(Dir(candidatePath, vbNormal)) > 0 Then OrderTplMasterDataFilePath = candidatePath
    End If
End Function

' 基本情報シートの施工会社セル(11行目)の値を取得する
Public Function OrderTplGetVendorCompanyName(ByVal wsInfo As Worksheet, ByVal vendorIndex As Long) As String
    If wsInfo Is Nothing Then Exit Function
    Dim valueColumn As Long
    valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(vendorIndex)
    OrderTplGetVendorCompanyName = CommonNormalizeText(CommonNzText( _
        wsInfo.Cells(BASIC_INFO_VENDOR_NAME_ROW, valueColumn).value))
End Function

' 業者マスタ(全社版)から会社名(B列 請求者氏名)で照合し、A列 業者名・P列 略称・O列 担当工事を返す
Public Function OrderTplResolveVendorMasterInfo(ByVal branchName As String, _
                                                ByVal officialName As String, _
                                                ByRef vendorName As String, _
                                                ByRef aliasText As String, _
                                                ByRef workText As String) As Boolean
    vendorName = ""
    aliasText = ""
    workText = ""

    Dim normalizedBranch As String
    Dim normalizedOfficial As String
    normalizedBranch = CommonNormalizeText(branchName)
    normalizedOfficial = CommonNormalizeText(officialName)
    If normalizedBranch = "" Or normalizedOfficial = "" Then Exit Function

    If mVendorInfoCache Is Nothing Then
        Set mVendorInfoCache = CreateObject("Scripting.Dictionary")
        mVendorInfoCache.CompareMode = vbTextCompare
    End If

    Dim cacheKey As String
    cacheKey = normalizedBranch & "|" & normalizedOfficial
    If mVendorInfoCache.Exists(cacheKey) Then
        Dim cachedInfo As Variant
        cachedInfo = mVendorInfoCache(cacheKey)
        vendorName = CStr(cachedInfo(0))
        aliasText = CStr(cachedInfo(1))
        workText = CStr(cachedInfo(2))
        OrderTplResolveVendorMasterInfo = (Len(vendorName) > 0)
        Exit Function
    End If

    Dim sourceFilePath As String
    sourceFilePath = OrderTplMasterDataFilePath(VENDOR_MASTER_FILE)
    If sourceFilePath = "" Then Exit Function

    Dim connection As Object
    Set connection = CommonOpenExcelAdoConnection(sourceFilePath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim sheetName As String
    sheetName = ResolveAdoWorksheetName(connection, normalizedBranch)
    If sheetName = "" Then GoTo Cleanup

    Dim recordset As Object
    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT * FROM [" & Replace$(sheetName, "]", "]]") & "$A2:P500]", connection, 0, 1, 1

    Do Until recordset.EOF
        Dim rowOfficial As String
        rowOfficial = CommonNormalizeText(CommonNzText( _
            CommonGetAdoFieldValue(recordset, ORDER_TPL_VENDOR_ADO_OFFICIAL_FIELD)))
        If StrComp(rowOfficial, normalizedOfficial, vbTextCompare) = 0 Then
            vendorName = CommonNormalizeText(CommonNzText( _
                CommonGetAdoFieldValue(recordset, ORDER_TPL_VENDOR_ADO_NAME_FIELD)))
            aliasText = CommonNormalizeText(CommonNzText( _
                CommonGetAdoFieldValue(recordset, ORDER_TPL_VENDOR_ADO_ALIAS_FIELD)))
            workText = CommonNormalizeText(CommonNzText( _
                CommonGetAdoFieldValue(recordset, ORDER_TPL_VENDOR_ADO_WORK_FIELD)))
            Exit Do
        End If
        recordset.MoveNext
    Loop
    CommonCloseAdoRecordset recordset

    mVendorInfoCache.Add cacheKey, Array(vendorName, aliasText, workText)
    OrderTplResolveVendorMasterInfo = (Len(vendorName) > 0)

Cleanup:
    CommonCloseAdoConnection connection
End Function

' 出張所別_単価適用線区.xlsx の単価適用線区シートから部店コード(G列)を取得する
Public Function OrderTplResolveBranchOfficeCode(ByVal branchName As String, _
                                                ByVal officeName As String) As String
    Dim normalizedBranch As String
    Dim normalizedOffice As String
    normalizedBranch = CommonNormalizeText(branchName)
    normalizedOffice = CommonNormalizeText(officeName)
    If normalizedBranch = "" Or normalizedOffice = "" Then Exit Function

    If mBranchOfficeCodeCache Is Nothing Then
        Set mBranchOfficeCodeCache = CreateObject("Scripting.Dictionary")
        mBranchOfficeCodeCache.CompareMode = vbTextCompare
    End If

    Dim cacheKey As String
    cacheKey = normalizedBranch & "|" & normalizedOffice
    If mBranchOfficeCodeCache.Exists(cacheKey) Then
        OrderTplResolveBranchOfficeCode = CStr(mBranchOfficeCodeCache(cacheKey))
        Exit Function
    End If

    Dim sourceFilePath As String
    sourceFilePath = OrderTplMasterDataFilePath(UNIT_PRICE_LINE_MASTER_FILE)
    If sourceFilePath = "" Then Exit Function

    Dim connection As Object
    Set connection = CommonOpenExcelAdoConnection(sourceFilePath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim resultCode As String
    Dim recordset As Object
    Set recordset = CreateObject("ADODB.Recordset")
    recordset.Open "SELECT * FROM [" & PRICE_LINE_SHEET & "$A2:G500]", connection, 0, 1, 1

    Do Until recordset.EOF
        If StrComp(CommonNormalizeText(CommonNzText(CommonGetAdoFieldValue(recordset, 1))), _
                   normalizedBranch, vbTextCompare) = 0 Then
            If StrComp(CommonNormalizeText(CommonNzText(CommonGetAdoFieldValue(recordset, 2))), _
                       normalizedOffice, vbTextCompare) = 0 Then
                resultCode = CommonNormalizeText(CommonNzText(CommonGetAdoFieldValue(recordset, 6)))
                Exit Do
            End If
        End If
        recordset.MoveNext
    Loop
    CommonCloseAdoRecordset recordset

    mBranchOfficeCodeCache.Add cacheKey, resultCode
    OrderTplResolveBranchOfficeCode = resultCode

Cleanup:
    CommonCloseAdoConnection connection
End Function

Private Function ResolveAdoWorksheetName(ByVal connection As Object, ByVal targetSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(connection)
    If sheetNames Is Nothing Then Exit Function

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If StrComp(CommonNormalizeText(CStr(sheetName)), targetSheetName, vbTextCompare) = 0 Then
            ResolveAdoWorksheetName = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

' テンプレートシート名 + 略称(P列)からシート名を組み立てる
Public Function OrderTplBuildSheetName(ByVal baseName As String, ByVal aliasText As String) As String
    OrderTplBuildSheetName = mod_Construction_OutputLayout.SanitizeSheetName(baseName & aliasText)
End Function

' 生成済みテンプレートシートか判定し、基本名と略称を返す
Public Function OrderTplIsGeneratedSheet(ByVal ws As Worksheet, _
                                         ByRef baseName As String, _
                                         ByRef aliasText As String) As Boolean
    baseName = ""
    aliasText = ""
    If ws Is Nothing Then Exit Function

    Dim normalizedName As String
    normalizedName = mod_Construction_Import_Load.NormalizeSheetNameParentheses(CommonNormalizeText(ws.Name))

    Dim baseNames As Variant
    baseNames = OrderTplTemplateSheetBaseNames()

    Dim i As Long
    For i = LBound(baseNames) To UBound(baseNames)
        Dim candidateBase As String
        candidateBase = CStr(baseNames(i))
        If Len(normalizedName) > Len(candidateBase) + 1 Then
            If StrComp(Left$(normalizedName, Len(candidateBase)), candidateBase, vbTextCompare) = 0 Then
                Dim suffixText As String
                suffixText = Mid$(normalizedName, Len(candidateBase) + 1)
                If Left$(suffixText, 1) = "(" And Right$(suffixText, 1) = ")" Then
                    baseName = candidateBase
                    aliasText = suffixText
                    OrderTplIsGeneratedSheet = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Function OrderTplSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    OrderTplSheetExists = Not ws Is Nothing
End Function

' 施工指示書(工事)/施工通知書(工事)の取込済みシートを探す(存在する方を自動使用)
Public Function OrderTplFindWorksSourceSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If mod_Construction_Import_Load.IsManagedImportOutputSheet(ws) Then
            If Not mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then
                If mod_Construction_Import_Load.SheetNameEndsWithSuffixText(ws.Name, CONSTRUCTION_SHEET_SUFFIX_WORKS) Then
                    If Not mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then
                        Set OrderTplFindWorksSourceSheet = ws
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ws
End Function

' 施工指示書(溶接)/施工通知書(溶接)の取込済みシートを探す(存在する方を自動使用)
Public Function OrderTplFindWeldingSourceSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If mod_Construction_Import_Load.IsManagedImportOutputSheet(ws) Then
            If Not mod_Construction_BasicTotals.IsPurchaseOutputSheet(ws) Then
                If mod_Construction_Import_Load.SheetNameEndsWithSuffixText(ws.Name, CONSTRUCTION_SHEET_SUFFIX_WELDING) Then
                    If mod_Construction_OutputLayout.IsWeldingOutputSheetCore(ws) Then
                        Set OrderTplFindWeldingSourceSheet = ws
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ws
End Function

' 契約線区名から接尾辞((軌道)/(溶接指示書用)/(溶接通知書用))を除去する
' 既存のマーカー除去関数へ委譲し、指示書・通知書の両表記に対応する
Public Function OrderTplStripLineSuffix(ByVal lineText As String, ByVal isWeldingSource As Boolean) As String
    Dim t As String
    t = mod_Construction_Import_Load.NormalizeSheetNameParentheses(CommonNormalizeText(lineText))

    If isWeldingSource Then
        t = mod_Construction_LineMapping.RemoveWeldingInstructionMarker(t)
    Else
        t = mod_Construction_LineMapping.RemoveTrackDesignationMarker(t)
    End If
    OrderTplStripLineSuffix = Trim$(t)
End Function

' 数量の単位が整数(桁切りなし)対象か判定する
Public Function OrderTplIsIntegerUnit(ByVal unitText As String) As Boolean
    OrderTplIsIntegerUnit = UnitListContains(OrderTplIntegerUnitListText(), unitText)
End Function

' 数量の単位が小数3桁対象か判定する
Public Function OrderTplIsDecimalUnit(ByVal unitText As String) As Boolean
    OrderTplIsDecimalUnit = UnitListContains(OrderTplDecimalUnitListText(), unitText)
End Function

Private Function UnitListContains(ByVal listText As String, ByVal unitText As String) As Boolean
    Dim normalizedUnit As String
    normalizedUnit = CommonRemoveAllSpaces(CommonNormalizeText(unitText))
    If normalizedUnit = "" Then Exit Function

    Dim items As Variant
    items = Split(listText, "|")

    Dim i As Long
    For i = LBound(items) To UBound(items)
        If StrComp(normalizedUnit, CStr(items(i)), vbBinaryCompare) = 0 Then
            UnitListContains = True
            Exit Function
        End If
    Next i
End Function

' 生成済みテンプレート(支店控/受注者用/注文請書)に残るプレースホルダー数式を除去する。
' テンプレート xlsx の自己参照(例: ='支店控(略称)'!E20)や #REF!、支店控へのミラー数式が
' ブック再計算時に循環参照ダイアログを出すため、VBA 転記前に値セルへ戻す。
Public Sub OrderTplRepairAllGeneratedPlaceholderFormulas()
    Dim ws As Worksheet
    Dim baseName As String
    Dim aliasText As String

    For Each ws In ThisWorkbook.Worksheets
        If OrderTplIsGeneratedSheet(ws, baseName, aliasText) Then
            OrderTplSanitizePlaceholderFormulas ws
        End If
    Next ws
End Sub

Public Sub OrderTplSanitizePlaceholderFormulas(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim baseName As String
    Dim aliasText As String
    If Not OrderTplIsGeneratedSheet(ws, baseName, aliasText) Then Exit Sub

    If baseName <> OrderTplBaseNameBranchCopyText() And _
       baseName <> OrderTplBaseNameContractorText() And _
       baseName <> OrderTplBaseNameAcceptanceText() Then Exit Sub

    Dim formulaRange As Range
    On Error Resume Next
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formulaRange Is Nothing Then Exit Sub

    Dim processed As Object
    Set processed = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In formulaRange.Cells
        Dim repCell As Range
        Set repCell = cell.MergeArea.Cells(1, 1)

        Dim repKey As String
        repKey = repCell.Address(True, True)
        If processed.Exists(repKey) Then GoTo ContinueCell
        processed.Add repKey, True

        If OrderTplShouldClearPlaceholderFormula(ws, baseName, repCell) Then
            OrderTplClearFormulaCell repCell
        End If
ContinueCell:
    Next cell
End Sub

Private Function OrderTplGetFormulaText(ByVal cell As Range) As String
    On Error Resume Next
    OrderTplGetFormulaText = CStr(cell.MergeArea.Cells(1, 1).Formula)
    On Error GoTo 0
End Function

Private Function OrderTplShouldClearPlaceholderFormula(ByVal ws As Worksheet, _
                                                       ByVal baseName As String, _
                                                       ByVal cell As Range) As Boolean
    Dim formulaText As String
    formulaText = OrderTplGetFormulaText(cell)
    If Len(formulaText) = 0 Then Exit Function
    If Left$(formulaText, 1) <> "=" Then Exit Function

    If InStr(1, formulaText, "#REF!", vbTextCompare) > 0 Then
        OrderTplShouldClearPlaceholderFormula = True
        Exit Function
    End If

    If OrderTplFormulaReferencesSameCell(ws, cell, formulaText) Then
        OrderTplShouldClearPlaceholderFormula = True
        Exit Function
    End If

    If baseName = OrderTplBaseNameContractorText() Or _
       baseName = OrderTplBaseNameAcceptanceText() Then
        If OrderTplFormulaReferencesBranchCopySheet(formulaText) Then
            OrderTplShouldClearPlaceholderFormula = True
        End If
    End If
End Function

Private Function OrderTplFormulaReferencesSameCell(ByVal ws As Worksheet, _
                                                 ByVal cell As Range, _
                                                 ByVal formulaText As String) As Boolean
    Dim normalizedFormula As String
    normalizedFormula = UCase$(Replace$(formulaText, " ", ""))

    Dim quotedSheet As String
    quotedSheet = "'" & Replace$(ws.Name, "'", "''") & "'!"

    Dim addr As Variant
    For Each addr In Array(cell.Address(True, False), cell.Address(False, False), _
                           cell.Address(True, True), cell.Address(False, True))
        If InStr(1, normalizedFormula, UCase$(quotedSheet & CStr(addr)), vbBinaryCompare) > 0 Then
            OrderTplFormulaReferencesSameCell = True
            Exit Function
        End If
    Next addr

    If InStr(1, normalizedFormula, "!", vbBinaryCompare) = 0 Then
        For Each addr In Array(cell.Address(True, False), cell.Address(False, False), _
                               cell.Address(True, True), cell.Address(False, True))
            If normalizedFormula = "=" & UCase$(CStr(addr)) Then
                OrderTplFormulaReferencesSameCell = True
                Exit Function
            End If
        Next addr
    End If
End Function

Private Function OrderTplFormulaReferencesBranchCopySheet(ByVal formulaText As String) As Boolean
    Dim branchBase As String
    branchBase = OrderTplBaseNameBranchCopyText()
    If InStr(1, formulaText, "'" & branchBase, vbTextCompare) > 0 Then
        OrderTplFormulaReferencesBranchCopySheet = True
    End If
End Function

Private Sub OrderTplClearFormulaCell(ByVal cell As Range)
    Dim writeCell As Range
    On Error Resume Next
    Set writeCell = cell.MergeArea.Cells(1, 1)
    If writeCell Is Nothing Then Set writeCell = cell
    On Error Resume Next
    writeCell.ClearContents
    On Error GoTo 0
End Sub

Public Sub OrderTplLog(ByVal msg As String)
    mod_DebugLog.Log "[OrderTpl] " & msg
End Sub
