Option Explicit

' ????: CHANGELOG.md ??
' mod_Construction_OutputLayout (split from mod_Construction_Order_Import)

Public Function IsManagedConstructionImportOutputSheet(ByVal ws As Worksheet) As Boolean
    IsManagedConstructionImportOutputSheet = mod_Construction_Import_Load.IsManagedImportOutputSheet(ws)
End Function

Public Function IsWeldingOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    IsWeldingOutputSheet = (mod_Construction_BasicTotals.FindHeaderColumn(ws, WELDING_VENDOR_HEADER) > 0)
End Function

Public Function OutputSheetCol(ByVal ws As Worksheet, ByVal baseCol As Long) As Long
    If ws Is Nothing Then
        OutputSheetCol = baseCol + WELDING_OUTPUT_COL_OFFSET
    ElseIf IsWeldingOutputSheet(ws) Then
        OutputSheetCol = baseCol + WELDING_OUTPUT_COL_OFFSET
    Else
        OutputSheetCol = baseCol
    End If
End Function

Public Function OutputSheetSeiriColumn(ByVal ws As Worksheet) As Long
    OutputSheetSeiriColumn = OutputSheetCol(ws, COL_SEIRI)
End Function

Public Function OutputSheetSubconPriceFirstCol(ByVal ws As Worksheet) As Long
    Dim jrAmountCol As Long
    jrAmountCol = mod_Construction_BasicTotals.FindHeaderColumn(ws, "JR金額")
    If jrAmountCol > 0 Then
        OutputSheetSubconPriceFirstCol = jrAmountCol + 1
        Exit Function
    End If

    If IsWeldingOutputSheet(ws) Then
        OutputSheetSubconPriceFirstCol = WELDING_SUBCON_PRICE_FIRST_COL
    Else
        OutputSheetSubconPriceFirstCol = SUBCON_PRICE_FIRST_COL
    End If
End Function

Public Function OutputSheetVendorColumns(ByVal ws As Worksheet) As Collection
    Dim result As New Collection
    If IsWeldingOutputSheet(ws) Then
        result.Add WELD_COL_WELDING_VENDOR
        result.Add WELD_COL_TRACK_VENDOR
    Else
        result.Add COL_VENDOR
    End If
    Set OutputSheetVendorColumns = result
End Function

Public Function WeldingRowArrayIndex(ByVal baseCol As Long) As Long
    If baseCol = COL_VENDOR Then
        WeldingRowArrayIndex = WELD_COL_WELDING_VENDOR - 1
    Else
        WeldingRowArrayIndex = baseCol + WELDING_OUTPUT_COL_OFFSET - 1
    End If
End Function

Public Function IsConstructionVendorOutputSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If mod_Construction_BasicTotals.FindHeaderColumn(ws, "施工業者") > 0 Then
        IsConstructionVendorOutputSheet = True
        Exit Function
    End If
    IsConstructionVendorOutputSheet = IsWeldingOutputSheet(ws)
End Function

Public Function JoinKeys(ByVal d As Object) As String
    Dim s As String
    If Not d Is Nothing Then
        Dim k As Variant
        For Each k In d.Keys
            s = s & IIf(s = "", "", ", ") & CStr(k)
        Next k
    End If
    If s = "" Then s = "(なし)"
    JoinKeys = s
End Function

Public Function BuildConstructionSheetName(ByVal sourceA3Text As String, _
                                            ByVal isWelding As Boolean) As String
    Dim suffix As String
    If isWelding Then
        suffix = CONSTRUCTION_SHEET_SUFFIX_WELDING
    Else
        suffix = CONSTRUCTION_SHEET_SUFFIX_WORKS
    End If

    Dim baseName As String
    baseName = SanitizeSheetName(sourceA3Text)
    If baseName = "" Then Exit Function

    Dim maxBaseLen As Long
    maxBaseLen = 31 - Len(suffix)
    If maxBaseLen < 1 Then maxBaseLen = 1
    If Len(baseName) > maxBaseLen Then baseName = Left$(baseName, maxBaseLen)

    BuildConstructionSheetName = baseName & suffix
End Function

Public Function SanitizeSheetName(ByVal s As String) As String
    Dim t As String
    t = CommonNormalizeText(s)
    Dim bad As Variant, ch As Variant
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In bad
        t = Replace$(t, CStr(ch), "_")
    Next ch
    t = Trim$(t)
    If Len(t) > 31 Then t = Left$(t, 31)
    SanitizeSheetName = t
End Function


Public Sub FillPurchaseUnitPrices(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, PURCHASE_NOTICE_SEIRI_COL).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim priceSheetName As String
    priceSheetName = ResolvePurchasePriceSheetName()
    If priceSheetName = "" Then
        LogCI "購入充当単価: シート名の解決に失敗(基本情報B6/C6 または 単価適用線区マスタ未一致)"
        Exit Sub
    End If

    Dim priceSheet As Worksheet
    On Error Resume Next
    Set priceSheet = ThisWorkbook.worksheets(priceSheetName)
    On Error GoTo 0
    If priceSheet Is Nothing Then
        LogCI "購入充当単価: シート「" & priceSheetName & "」が見つかりません"
        MsgBox "単価シート「" & priceSheetName & "」が見つかりませんでした。" & vbCrLf & _
               "購入充当単価の取込みをスキップします。", vbExclamation
        Exit Sub
    End If

    Dim priceMap As Object
    Set priceMap = CreateObject("Scripting.Dictionary")
    priceMap.CompareMode = vbTextCompare

    Dim pLast As Long, pr As Long, key As String
    pLast = priceSheet.Cells(priceSheet.rows.Count, PURCHASE_PRICE_KEY_COL).End(xlUp).Row
    For pr = PURCHASE_PRICE_DATA_START_ROW To pLast
        key = CommonRemoveAllSpaces(CommonNzText(priceSheet.Cells(pr, PURCHASE_PRICE_KEY_COL).value))
        If key <> "" And Not priceMap.Exists(key) Then
            priceMap.Add key, priceSheet.Cells(pr, PURCHASE_PRICE_VALUE_COL).value
        End If
    Next pr

    Dim matched As Long, lookupCount As Long, unmatched As Long
    Dim r As Long, lookupKey As String
    For r = 2 To lastRow
        lookupKey = mod_Construction_LineMapping.NormalizeRecordKey(ws.Cells(r, PURCHASE_PRICE_LOOKUP_COL).value)
        If lookupKey <> "" Then
            lookupCount = lookupCount + 1
            If priceMap.Exists(lookupKey) Then
                ws.Cells(r, PURCHASE_NOTICE_AUTO_PRICE_COL).value = priceMap(lookupKey)
                matched = matched + 1
            Else
                unmatched = unmatched + 1
            End If
        End If
    Next r

    ws.Range(ws.Cells(2, PURCHASE_NOTICE_AUTO_AMOUNT_COL), _
             ws.Cells(lastRow, PURCHASE_NOTICE_AUTO_AMOUNT_COL)).FormulaR1C1 = _
        "=RC[" & (PURCHASE_NOTICE_AUTO_PRICE_COL - PURCHASE_NOTICE_AUTO_AMOUNT_COL) & _
        "]*RC[" & (PURCHASE_NOTICE_QTY_COL - PURCHASE_NOTICE_AUTO_AMOUNT_COL) & "]"

    For r = 2 To lastRow
        WritePriceComparisonAtColumns _
            ws, r, priceSheetName, PURCHASE_NOTICE_AUTO_PRICE_COL, _
            PURCHASE_NOTICE_JR_PRICE_COL, PURCHASE_NOTICE_PRICE_COMPARE_COL, _
            PURCHASE_NOTICE_PRICE_GUIDANCE_COL, False
    Next r

    With ws.Range(ws.Cells(2, PURCHASE_NOTICE_AUTO_PRICE_COL), _
                  ws.Cells(lastRow, PURCHASE_NOTICE_AUTO_AMOUNT_COL))
        .NumberFormatLocal = "#,##0;[赤]-#,##0"
    End With
    ws.Range(ws.Cells(1, PURCHASE_NOTICE_PRICE_COMPARE_COL), _
             ws.Cells(lastRow, PURCHASE_NOTICE_PRICE_COMPARE_COL)).HorizontalAlignment = xlCenter

    LogCI "購入充当単価転記: " & matched & " 件 / 照合対象=" & lookupCount & _
          " / 未一致=" & unmatched & " (単価シート=" & priceSheetName & ")"
End Sub

Public Function ResolvePurchasePriceSheetName() As String
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Function

    Dim BranchName As String, officeName As String
    BranchName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))
    officeName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    If BranchName = "" Or officeName = "" Then Exit Function

    Dim masterPath As String
    Dim connection As Object
    Set connection = OpenUnitPriceMasterAdoConnection(masterPath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim actualSheetName As String
    actualSheetName = FindAdoWorksheetName(connection, PRICE_LINE_SHEET)

    Dim resultName As String
    If actualSheetName = "" Then
        LogCI "購入充当単価: マスタに「" & PRICE_LINE_SHEET & "」シートがありません"
    Else
        Dim recordset As Object
        Set recordset = CreateObject("ADODB.Recordset")
        recordset.Open "SELECT [F2], [F3], [F5] FROM " & _
                       BuildAdoSheetTableName(actualSheetName), connection, 0, 1, 1

        Dim b As String, c As String, nameText As String
        Do Until recordset.EOF
            b = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(0).value))
            c = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(1).value))
            If b = BranchName And c = officeName Then
                nameText = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(2).value))
                If nameText <> "" Then
                    resultName = nameText & PURCHASE_PRICE_SHEET_SUFFIX
                    Exit Do
                End If
            End If
            recordset.MoveNext
        Loop
    End If

Cleanup:
    If Err.Number <> 0 Then
        LogCI "購入充当単価マスタADO読込エラー Err " & Err.Number & ": " & Err.Description
    End If
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
    ResolvePurchasePriceSheetName = resultName
End Function

Public Function ResolveWeldingPriceSheetName() As String
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then Exit Function

    Dim BranchName As String, officeName As String
    BranchName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_BRANCH_CELL).value))
    officeName = CommonRemoveAllSpaces(CommonNzText(wsInfo.Range(BASIC_INFO_OFFICE_CELL).value))
    If BranchName = "" Or officeName = "" Then Exit Function

    Dim masterPath As String
    Dim connection As Object
    Set connection = OpenUnitPriceMasterAdoConnection(masterPath)
    If connection Is Nothing Then Exit Function

    On Error GoTo Cleanup

    Dim actualSheetName As String
    Dim recordset As Object
    actualSheetName = FindAdoWorksheetName(connection, PRICE_LINE_SHEET)

    Dim resultName As String
    If actualSheetName = "" Then
        LogCI "レール溶接単価: マスタに「" & PRICE_LINE_SHEET & "」シートがありません"
    Else
        Set recordset = CreateObject("ADODB.Recordset")
        recordset.Open "SELECT [F2], [F3], [F5] FROM " & _
                       BuildAdoSheetTableName(actualSheetName), connection, 0, 1, 1

        Dim b As String, c As String, nameText As String
        Do Until recordset.EOF
            b = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(0).value))
            c = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(1).value))
            If b = BranchName And c = officeName Then
                nameText = CommonRemoveAllSpaces(CommonNzText(recordset.Fields(2).value))
                If nameText <> "" Then
                    resultName = nameText & WELDING_PRICE_SHEET_SUFFIX
                    Exit Do
                End If
            End If
            recordset.MoveNext
        Loop
    End If

Cleanup:
    If Err.Number <> 0 Then
        LogCI "レール溶接単価マスタADO読込エラー Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    CommonCloseAdoRecordset recordset
    CommonCloseAdoConnection connection
    ResolveWeldingPriceSheetName = resultName
End Function

' レール溶接単価シート A列の「積算線区：」行から線区名(B列)を読み取り、
' 整理番号行(B列=数値)と組み合わせた照合キーを構築する。
Public Sub BuildWeldingUnitPriceRowCache(ByVal priceSheet As Worksheet, _
                                          ByVal priceValueMap As Object, _
                                          Optional ByVal storeRowIndex As Boolean = False, _
                                          Optional ByVal rowIndexMap As Object = Nothing)
    Dim lastRow As Long
    lastRow = priceSheet.Cells(priceSheet.rows.Count, WELDING_PRICE_SEIRI_COL).End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    Dim currentLineSection As String
    currentLineSection = ""

    Dim r As Long
    For r = 1 To lastRow
        If IsWeldingLineSectionLabelCell(CommonNzText(priceSheet.Cells(r, 1).value)) Then
            currentLineSection = NormalizeWeldingPriceLineSectionName( _
                CommonNzText(priceSheet.Cells(r, WELDING_PRICE_SEIRI_COL).value))
        End If

        If IsWeldingPriceSeiriCellValue(priceSheet.Cells(r, WELDING_PRICE_SEIRI_COL).value) Then
            Dim seiriKey As String
            seiriKey = BuildWeldingLookupKey(priceSheet.Cells(r, WELDING_PRICE_SEIRI_COL).value)
            If seiriKey <> "" Then
                If storeRowIndex Then
                    RegisterWeldingPriceRowIndex rowIndexMap, currentLineSection, seiriKey, r
                ElseIf Not priceValueMap Is Nothing Then
                    RegisterWeldingPriceCacheEntry priceValueMap, currentLineSection, seiriKey, Array( _
                        priceSheet.Cells(r, UNIT_PRICE_DAY_PRICE_COL).value, _
                        priceSheet.Cells(r, UNIT_PRICE_NIGHT_PRICE_COL).value)
                End If
            End If
        End If
    Next r
End Sub

Public Sub RegisterWeldingPriceCacheEntry(ByVal priceValueMap As Object, _
                                           ByVal unitLineSection As String, _
                                           ByVal seiriKey As String, _
                                           ByVal values As Variant)
    Dim primaryKey As String
    primaryKey = BuildWeldingLineSeiriLookupKeyFromParts(unitLineSection, seiriKey)
    If primaryKey = "" Then Exit Sub
    If Not priceValueMap.Exists(primaryKey) Then priceValueMap.Add primaryKey, values

    EnsureProjectLineNameAliasMapsLoaded
    If mProjectLineNameReverseAliasWelding Is Nothing Then Exit Sub
    If Len(unitLineSection) = 0 Then Exit Sub
    If Not mProjectLineNameReverseAliasWelding.Exists(unitLineSection) Then Exit Sub

    Dim sourceNames As Collection
    Set sourceNames = mProjectLineNameReverseAliasWelding(unitLineSection)
    Dim sourceName As Variant
    For Each sourceName In sourceNames
        Dim aliasKey As String
        aliasKey = BuildWeldingLineSeiriLookupKeyFromParts(CStr(sourceName), seiriKey)
        If aliasKey <> "" And Not priceValueMap.Exists(aliasKey) Then
            priceValueMap.Add aliasKey, values
        End If
    Next sourceName
End Sub

Public Sub RegisterWeldingPriceRowIndex(ByVal rowIndexMap As Object, _
                                         ByVal unitLineSection As String, _
                                         ByVal seiriKey As String, _
                                         ByVal rowIndex As Long)
    If rowIndexMap Is Nothing Then Exit Sub

    Dim primaryKey As String
    primaryKey = BuildWeldingLineSeiriLookupKeyFromParts(unitLineSection, seiriKey)
    If primaryKey = "" Then Exit Sub
    If Not rowIndexMap.Exists(primaryKey) Then rowIndexMap.Add primaryKey, rowIndex

    EnsureProjectLineNameAliasMapsLoaded
    If mProjectLineNameReverseAliasWelding Is Nothing Then Exit Sub
    If Len(unitLineSection) = 0 Then Exit Sub
    If Not mProjectLineNameReverseAliasWelding.Exists(unitLineSection) Then Exit Sub

    Dim sourceNames As Collection
    Set sourceNames = mProjectLineNameReverseAliasWelding(unitLineSection)
    Dim sourceName As Variant
    For Each sourceName In sourceNames
        Dim aliasKey As String
        aliasKey = BuildWeldingLineSeiriLookupKeyFromParts(CStr(sourceName), seiriKey)
        If aliasKey <> "" And Not rowIndexMap.Exists(aliasKey) Then
            rowIndexMap.Add aliasKey, rowIndex
        End If
    Next sourceName
End Sub

Public Function BuildWeldingLineSeiriLookupKeyFromParts(ByVal normalizedLineSection As String, _
                                                         ByVal seiriKey As String) As String
    If seiriKey = "" Then Exit Function
    If Len(normalizedLineSection) = 0 Then
        BuildWeldingLineSeiriLookupKeyFromParts = seiriKey
    Else
        BuildWeldingLineSeiriLookupKeyFromParts = normalizedLineSection & "|" & seiriKey
    End If
End Function

Public Function BuildWeldingLineSeiriLookupKey(ByVal lineText As String, _
                                                ByVal seiriValue As Variant, _
                                                Optional ByVal useLineAlias As Boolean = True) As String
    Dim seiriKey As String
    seiriKey = BuildWeldingLookupKey(seiriValue)
    If seiriKey = "" Then Exit Function

    Dim normalizedLine As String
    normalizedLine = NormalizeWeldingPriceLineSectionName(lineText)
    If useLineAlias Then
        normalizedLine = mod_Construction_LineMapping.ResolveWeldingLineSectionAlias(normalizedLine)
    End If
    BuildWeldingLineSeiriLookupKey = BuildWeldingLineSeiriLookupKeyFromParts(normalizedLine, seiriKey)
End Function

Public Function NormalizeWeldingPriceLineSectionName(ByVal lineText As String) As String
    NormalizeWeldingPriceLineSectionName = mod_Construction_LineMapping.NormalizeLineLookupText(lineText, True)
End Function

Public Function IsWeldingPriceSeiriCellValue(ByVal value As Variant) As Boolean
    Dim textValue As String
    textValue = Trim$(StrConv(CommonNzText(value), vbNarrow))
    If Len(textValue) = 0 Then Exit Function
    IsWeldingPriceSeiriCellValue = IsNumeric(textValue)
End Function

Public Function IsWeldingLineSectionLabelCell(ByVal labelText As String) As Boolean
    Dim normalized As String
    normalized = CommonRemoveAllSpaces(CommonNormalizeText(labelText))
    If Len(normalized) = 0 Then Exit Function
    IsWeldingLineSectionLabelCell = _
        (InStr(1, normalized, WeldingLineSectionLabelKeywordText(), vbTextCompare) > 0)
End Function

Public Function WeldingLineSectionLabelKeywordText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H7A4D) & ChrW$(&H7B97) & ChrW$(&H7DDA) & ChrW$(&H533A)
    WeldingLineSectionLabelKeywordText = cached
End Function

Public Function BuildWeldingLookupKey(ByVal seiriValue As Variant) As String
    Dim keyText As String
    keyText = CommonRemoveAllSpaces(CommonNzText(seiriValue))
    If keyText = "" Then Exit Function
    If IsNumeric(keyText) Then
        Dim seiriNumber As Long
        seiriNumber = CLng(CDbl(keyText))
        ' 5桁の整理番号は下4桁を溶接単価シートの整理番号として照合する。
        ' 下4桁の先頭ゼロは除去して正規化する(例:20130->0130->130)。
        ' 単価シート側が先頭ゼロ無し(130)で格納されていても一致させるため。
        If seiriNumber >= 10000 And seiriNumber <= 99999 Then
            BuildWeldingLookupKey = CStr(CLng(Right$(Format$(seiriNumber, "00000"), 4)))
        Else
            BuildWeldingLookupKey = CStr(seiriNumber)
        End If
    Else
        If Len(keyText) = 5 And IsNumeric(keyText) Then
            BuildWeldingLookupKey = CStr(CLng(Right$(keyText, 4)))
        Else
            BuildWeldingLookupKey = keyText
        End If
    End If
End Function

Public Function SheetExistsByName(ByVal sheetName As String) As Boolean
    If sheetName = "" Then Exit Function
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.worksheets(sheetName)
    On Error GoTo 0
    SheetExistsByName = Not ws Is Nothing
End Function

Public Function OpenUnitPriceMasterAdoConnection(ByRef resolvedPath As String) As Object
    Dim candidates As Collection
    Set candidates = New Collection

    AddUniqueText candidates, mod_Construction_Import_Load.ResolveMasterFilePath()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim documentRoot As String
    If Len(ThisWorkbook.Path) > 0 Then
        documentRoot = fso.GetParentFolderName(ThisWorkbook.Path)
        AddUniqueText candidates, fso.BuildPath(documentRoot, _
            MASTER_DATA_FOLDER & "\" & UNIT_PRICE_LINE_MASTER_FILE)
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then
        userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    End If
    If Len(Trim$(userProfilePath)) > 0 Then
        AddUniqueText candidates, userProfilePath & "\" & CommonCompanyNameText() & "\" & _
            "線路出張所用_注文書_請求書アクセスサイト - ドキュメント\" & _
            MASTER_DATA_FOLDER & "\" & UNIT_PRICE_LINE_MASTER_FILE
    End If

    Dim candidate As Variant
    Dim connection As Object
    For Each candidate In candidates
        If fso.FileExists(CStr(candidate)) Then
            Set connection = CommonOpenExcelAdoConnection(CStr(candidate))
            If Not connection Is Nothing Then
                resolvedPath = CStr(candidate)
                Set OpenUnitPriceMasterAdoConnection = connection
                Exit Function
            End If
        End If
    Next candidate
End Function

Public Function FindAdoWorksheetName(ByVal connection As Object, _
                                      ByVal expectedSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(connection)

    Dim expectedKey As String
    expectedKey = CommonRemoveAllSpaces(CommonNormalizeText(expectedSheetName))

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If StrComp(CommonRemoveAllSpaces(CommonNormalizeText(CStr(sheetName))), _
                   expectedKey, vbTextCompare) = 0 Then
            FindAdoWorksheetName = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

Public Function BuildAdoSheetTableName(ByVal sheetName As String) As String
    BuildAdoSheetTableName = "[" & Replace$(sheetName, "]", "]]") & "$]"
End Function

' "線路出張所用_注文書_請求書アクセスサイト - ドキュメント"
Public Function OrderInvoiceDocumentFolderTextCI() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7DDA) & ChrW$(&H8DEF) & ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & ChrW$(&H7528) & _
                 "_" & ChrW$(&H6CE8) & ChrW$(&H6587) & ChrW$(&H66F8) & "_" & _
                 ChrW$(&H8ACB) & ChrW$(&H6C42) & ChrW$(&H66F8) & _
                 ChrW$(&H30A2) & ChrW$(&H30AF) & ChrW$(&H30BB) & ChrW$(&H30B9) & _
                 ChrW$(&H30B5) & ChrW$(&H30A4) & ChrW$(&H30C8) & _
                 " - " & ChrW$(&H30C9) & ChrW$(&H30AD) & ChrW$(&H30E5) & ChrW$(&H30E1) & ChrW$(&H30F3) & ChrW$(&H30C8)
    End If
    OrderInvoiceDocumentFolderTextCI = cached
End Function

' "在来線及び新幹線"
Public Function CombinedLineTypeMasterFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5728) & ChrW$(&H6765) & ChrW$(&H7DDA) & ChrW$(&H53CA) & ChrW$(&H3073) & _
                 ChrW$(&H65B0) & ChrW$(&H5E79) & ChrW$(&H7DDA)
    End If
    CombinedLineTypeMasterFolderText = cached
End Function

Public Sub WriteReferenceValueToBasicInfo(ByVal refValue As Variant, ByVal sourceCell As String)
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet(ThisWorkbook)
    If wsInfo Is Nothing Then
        LogCI "基本情報シートが見つからないため " & BASIC_INFO_REF_VALUE_CELL & " への転記をスキップ"
        Exit Sub
    End If

    With wsInfo.Range(BASIC_INFO_REF_VALUE_CELL)
        .value = refValue
        .HorizontalAlignment = xlCenter
        .Font.Name = BASIC_INFO_REF_FONT_NAME
    End With
    LogCI BASIC_INFO_REF_VALUE_CELL & " に参照シート " & sourceCell & " を転記"
End Sub
