Option Explicit

'==========================================================================
'  工事単価インポートモジュール
'    改修内容（#17）：
'      - ResetUnitPriceValidation の AlertStyle を
'        xlValidAlertStop → xlValidAlertInformation に変更。
'        非表示列をリスト元に使用する場合、xlValidAlertStop では
'        Excel がリスト外値と判定して入力をブロックするため。
'    改修内容（#18）：
'      - ResetUnitPriceValidation を MergeArea 対応に修正。
'        C22 等がセル結合されている場合、targetCell が結合副セル
'        （D22 等）を指してしまい入力規則が誤ったセルに設定される
'        問題を解消。MergeArea.Cells(1,1) で代表セルを確実に取得する。
'==========================================================================
Public SharedMasterData As Variant

Private Type UnitPriceRequest
    Nendo As String
    BranchName As String
    OfficeName As String
    lineType As String
    projectName As String
    UnitPriceKind As String
End Type

Private Type UnitPriceMasterRow
    BranchName As String
    OfficeName As String
    BranchGroupName As String
    UnitPriceSectionName As String
End Type

Private Const MASTER_SHEET_NAME As String = "単価適用線区"
Private Const PROJECT_MASTER_SHEET_NAME As String = "単価適用工事件名マスタ"
Private Const ZAIRAISEN_DISTINCTION_SHEET_NAME As String = "在幹区分"
Private Const UNIT_PRICE_DATA_FOLDER As String = "単価データ"
Private Const UNIT_PRICE_REFERENCE_FOLDER As String = "工事件名別マスタ"
Private Const UNIT_PRICE_MASTER_FILE As String = "出張所別_単価適用線区.xlsx"

Private Const ZAIRAISEN_NAME As String = "在来線"
Private Const SHINKANSEN_NAME As String = "新幹線"
Private Const INITIAL_PRICE_NAME As String = "年初単価"
Private Const DESIGN_CHANGE_PRICE_NAME As String = "設計変更単価"
Private Const NORMAL_PRICE_FOLDER As String = "通常単価"

Private Const BASIC_INFO_YEAR_CELL As String = "B4"
Private Const BASIC_INFO_BRANCH_CELL As String = "B6"
Private Const BASIC_INFO_OFFICE_CELL As String = "C6"
Private Const BASIC_INFO_WORK_NAME_CELL As String = "C10"
Private Const BASIC_INFO_LINE_TYPE_CELL As String = "C20"
Private Const BASIC_INFO_PROJECT_NAME_CELL As String = "C21"
Private Const BASIC_INFO_PRICE_KIND_CELL As String = "C22"
Private Const BASIC_INFO_PRICE_KIND_FALLBACK_CELL As String = "B22"
Private Const BASIC_INFO_IMPORTED_LINE_NAMES_CELL As String = "C24"
Private Const BASIC_INFO_WELDING_FLAG_CELL As String = "C23"
Private Const PROJECT_NAME_LIST_COL As String = "AE"
Private Const LINE_TYPE_LIST_COL As String = "AF"
Private Const PRICE_KIND_LIST_COL As String = "AG"
Private Const LIST_START_ROW As Long = 2
Private Const PROJECT_NAME_MASTER_START_ROW As Long = 2
Private Const PROJECT_NAME_MASTER_LAST_ROW As Long = 1048576
Private Const IMPORTED_SHEET_PROPERTY As String = "UnitPriceImported"

Private Const UNIT_PRICE_FILE_KEYWORD As String = "軌道材料購入充当"
Private Const PURCHASE_SHEET_NAME_SUFFIX As String = "_購入充当単価"
Private Const PURCHASE_SHEET_TAB_R As Long = 255
Private Const PURCHASE_SHEET_TAB_G As Long = 204
Private Const PURCHASE_SHEET_TAB_B As Long = 153
Private Const WELDING_REQUIRED_VALUE As String = "溶接工事あり"
Private Const WELDING_FILE_KEYWORD As String = "⑧レール溶接工事"
Private Const WELDING_FILE_FALLBACK_KEYWORD As String = "レール溶接工事"
Private Const WELDING_SHEET_NAME_SUFFIX As String = "_レール溶接単価"
Private Const WELDING_SHEET_TAB_R As Long = 252
Private Const WELDING_SHEET_TAB_G As Long = 213
Private Const WELDING_SHEET_TAB_B As Long = 180
Private Const SOURCE_HEADER_ROW As Long = 1
Private Const IMPORTED_LINE_NAME_BASE_ROW_HEIGHT As Double = 28.5
Private Const IMPORTED_LINE_NAME_MAX_LINES As Long = 7

Private Const ZAIRAISEN_PROJECT_NAME_FIRST_END_ROW  As Long = 8
Private Const ZAIRAISEN_PROJECT_NAME_SKIP_ROW       As Long = 9
Private Const ZAIRAISEN_PROJECT_NAME_SECOND_END_ROW As Long = 11
Private Const SHINKANSEN_PROJECT_NAME_END_ROW       As Long = 6

Public Function GetMasterFilePath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim candidates As Collection
    Set candidates = New Collection

    If Len(ThisWorkbook.Path) > 0 Then
        candidates.Add fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.Path), _
                       UnitPriceMasterFolderText() & "\" & UNIT_PRICE_REFERENCE_FOLDER & "\" & UNIT_PRICE_MASTER_FILE)
        candidates.Add fso.BuildPath(ThisWorkbook.Path, _
                       UNIT_PRICE_REFERENCE_FOLDER & "\" & UNIT_PRICE_MASTER_FILE)
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    If Len(Trim$(userProfilePath)) > 0 Then
        candidates.Add userProfilePath & "\" & CommonCompanyNameText() & "\" & _
                       OrderInvoiceDocumentFolderText() & "\" & UnitPriceMasterFolderText() & "\" & _
                       UNIT_PRICE_REFERENCE_FOLDER & "\" & UNIT_PRICE_MASTER_FILE
    End If

    GetMasterFilePath = FirstExistingFilePath(candidates)
End Function

Public Sub AutoFillLineTypeFromWorkName(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim workName As String
    workName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_WORK_NAME_CELL).Value))
    If workName = "" Then Exit Sub

    Dim masterFilePath As String
    masterFilePath = GetMasterFilePath()
    If masterFilePath = "" Then Exit Sub

    Dim lineType As String
    lineType = LookupLineTypeByWorkName(masterFilePath, workName)
    If lineType = "" Then Exit Sub

    If StrComp(CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).Value)), lineType, vbTextCompare) = 0 Then Exit Sub

    wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).Value = lineType
End Sub

Public Sub AutoFillProjectNameFromWorkName(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim workName As String
    workName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_WORK_NAME_CELL).Value))
    If workName = "" Then Exit Sub

    Dim lineType As String
    lineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).Value))
    If lineType = "" Then Exit Sub

    Dim masterFilePath As String
    masterFilePath = GetMasterFilePath()
    If masterFilePath = "" Then Exit Sub

    Dim sheetName As String
    Dim dummy As Long
    ResolveUnitPriceProjectNameMasterConfig lineType, sheetName, dummy
    If sheetName = "" Then Exit Sub

    Dim projectName As String
    projectName = LookupProjectNameByWorkName(masterFilePath, sheetName, workName)
    If projectName = "" Then Exit Sub

    If StrComp(CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).Value)), projectName, vbTextCompare) = 0 Then Exit Sub

    wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).Value = projectName
End Sub

Private Function LookupProjectNameByWorkName(ByVal masterFilePath As String, _
                                             ByVal sheetName As String, _
                                             ByVal workName As String) As String
    Dim cn As Object
    Dim rs As Object

    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(masterFilePath)
    If cn Is Nothing Then Exit Function

    Dim sql As String
    sql = "SELECT F1, F2 FROM " & _
          BuildAdoSheetRangeName(sheetName, "A", PROJECT_NAME_MASTER_START_ROW, "B", PROJECT_NAME_MASTER_LAST_ROW) & _
          " WHERE F2 IS NOT NULL"

    Set rs = cn.Execute(sql)

    Dim normalizedWorkName As String
    normalizedWorkName = NormalizeMatchText(workName)

    Do Until rs.EOF
        Dim keyword As String
        keyword = NormalizeMatchText(CommonNzText(CommonGetAdoFieldValue(rs, 1)))
        If keyword <> "" Then
            If InStr(1, normalizedWorkName, keyword, vbTextCompare) > 0 Then
                Dim projectNameRaw As String
                projectNameRaw = Trim$(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
                If projectNameRaw <> "" Then
                    LookupProjectNameByWorkName = CommonNormalizeText(projectNameRaw)
                    GoTo Cleanup
                End If
            End If
        End If
        rs.MoveNext
    Loop

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    Resume Cleanup
End Function

Private Function LookupLineTypeByWorkName(ByVal masterFilePath As String, _
                                          ByVal workName As String) As String
    Dim cn As Object
    Dim rs As Object

    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(masterFilePath)
    If cn Is Nothing Then Exit Function

    Dim sql As String
    sql = "SELECT F1, F2 FROM " & _
          BuildAdoSheetRangeName(ZAIRAISEN_DISTINCTION_SHEET_NAME, "A", 2, "B", PROJECT_NAME_MASTER_LAST_ROW) & _
          " WHERE F1 IS NOT NULL"

    Set rs = cn.Execute(sql)

    Dim normalizedWorkName As String
    normalizedWorkName = NormalizeMatchText(workName)

    Do Until rs.EOF
        Dim keyword As String
        keyword = NormalizeMatchText(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
        If keyword <> "" Then
            If InStr(1, normalizedWorkName, keyword, vbTextCompare) > 0 Then
                Dim lineTypeRaw As String
                lineTypeRaw = Trim$(CommonNzText(CommonGetAdoFieldValue(rs, 1)))
                If lineTypeRaw <> "" Then
                    LookupLineTypeByWorkName = CommonNormalizeText(lineTypeRaw)
                    GoTo Cleanup
                End If
            End If
        End If
        rs.MoveNext
    Loop

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    Resume Cleanup
End Function

Public Sub RefreshUnitPriceProjectNameValidation(Optional ByVal wsInfo As Worksheet, _
                                                  Optional ByVal keepProjectName As Boolean = True)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    WriteUnitPriceLineTypeValidation wsInfo
    WriteUnitPriceKindValidation wsInfo

    Dim lineType As String
    lineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).Value))
    Dim projectNames As Collection
    Set projectNames = LoadUnitPriceProjectNamesForBasicInfo(GetMasterFilePath(), lineType)
    If projectNames Is Nothing Then
        ClearUnitPriceProjectNameValidation wsInfo, True
        Exit Sub
    End If

    If projectNames.Count = 0 Then
        ClearUnitPriceProjectNameValidation wsInfo, True
        Exit Sub
    End If

    WriteUnitPriceProjectNameValidation wsInfo, projectNames, keepProjectName
    Exit Sub

ErrorHandler:
    ClearUnitPriceProjectNameValidation wsInfo, True
End Sub

Public Sub ImportConstructionUnitPriceForBasicInfo()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    ClearAndImportUnitPriceForBasicInfo wsInfo
End Sub

Public Sub ClearAndImportUnitPriceForBasicInfo(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ImportUnitPriceData ws
End Sub

Private Sub ImportUnitPriceData(ByVal wsInfo As Worksheet)
    Dim request As UnitPriceRequest
    If Not TryReadUnitPriceRequest(wsInfo, request) Then Exit Sub

    Dim masterRow As UnitPriceMasterRow
    If Not TryLoadUnitPriceMasterRow(request, masterRow) Then Exit Sub

    Dim priceFolderPath As String
    Dim sectionFolderPath As String
    Dim sourceFilePath As String
    sourceFilePath = ResolveUnitPriceSourceFilePath(request, masterRow, priceFolderPath, sectionFolderPath)
    If sourceFilePath = "" Then Exit Sub

    Dim sheetNames As Collection
    Set sheetNames = LoadWorksheetNamesFromWorkbook(sourceFilePath)
    If sheetNames Is Nothing Then
        MsgBox "単価表ブックに取り込み可能なシートが見つかりませんでした。" & vbCrLf & sourceFilePath, vbExclamation
        Exit Sub
    End If
    If sheetNames.Count = 0 Then
        MsgBox "単価表ブックに取り込み可能なシートが見つかりませんでした。" & vbCrLf & sourceFilePath, vbExclamation
        Exit Sub
    End If

    Dim selectedSheetNames As Collection
    Set selectedSheetNames = PromptLineNameSelection(sheetNames)
    If selectedSheetNames Is Nothing Then Exit Sub
    If selectedSheetNames.Count = 0 Then Exit Sub

    If Not ImportSelectedUnitPriceSheets(sourceFilePath, selectedSheetNames, wsInfo.Parent) Then Exit Sub
    WriteSelectedLineNames wsInfo, selectedSheetNames

    Dim purchaseSheetName As String
    Call ImportPurchaseUnitPriceSheetsByReference(request, masterRow, priceFolderPath, sectionFolderPath, wsInfo.Parent, purchaseSheetName)

    Dim weldingSheetName As String
    Call ImportWeldingUnitPriceSheetsIfRequired(wsInfo, masterRow, priceFolderPath, sectionFolderPath, selectedSheetNames, wsInfo.Parent, weldingSheetName)

    MsgBox BuildImportCompleteMessage(selectedSheetNames, sourceFilePath, purchaseSheetName, weldingSheetName), vbInformation, "完了"
End Sub

Private Function TryReadUnitPriceRequest(ByVal wsInfo As Worksheet, ByRef request As UnitPriceRequest) As Boolean
    request.Nendo = CommonExtractYear4Digits(CStr(wsInfo.Range(BASIC_INFO_YEAR_CELL).Value))
    request.BranchName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_BRANCH_CELL).Value))
    request.OfficeName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_OFFICE_CELL).Value))
    request.lineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).Value))
    request.projectName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).Value))
    request.UnitPriceKind = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PRICE_KIND_CELL).Value))
    If request.UnitPriceKind = "" Then
        request.UnitPriceKind = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PRICE_KIND_FALLBACK_CELL).Value))
    End If
    If InStr(1, NormalizeMatchText(request.UnitPriceKind), NormalizeMatchText("単価適用区分"), vbTextCompare) > 0 Then
        request.UnitPriceKind = ""
    End If

    If request.Nendo = "" Then
        MsgBox "基本情報シート B4 に4桁の年度が見つかりません。", vbExclamation
        Exit Function
    End If
    If request.BranchName = "" Or request.OfficeName = "" Then
        MsgBox "基本情報シート B6 または C6 が空です。支店名・出張所名を確認してください。", vbExclamation
        Exit Function
    End If
    If request.lineType = "" Then
        MsgBox "基本情報シート C20 の線区区分を選択してください。", vbExclamation
        Exit Function
    End If
    If request.projectName = "" Then
        MsgBox "基本情報シート C21 の単価適用工事件名を選択してください。", vbExclamation
        Exit Function
    End If
    If IsPurchaseUnitPriceProjectName(request.projectName) Then
        MsgBox "基本情報シート C21 は軌道材料購入充当以外の単価適用工事件名を選択してください。" & vbCrLf & _
               "購入充当単価は C24 確定後に自動作成します。", vbExclamation
        Exit Function
    End If
    If request.UnitPriceKind = "" Then
        MsgBox "基本情報シート C22 の単価区分を選択してください。", vbExclamation
        Exit Function
    End If

    TryReadUnitPriceRequest = True
End Function

Private Function TryLoadUnitPriceMasterRow(ByRef request As UnitPriceRequest, _
                                           ByRef masterRow As UnitPriceMasterRow) As Boolean
    Dim sourceFilePath As String
    sourceFilePath = GetMasterFilePath()
    If sourceFilePath = "" Then
        MsgBox "出張所別_単価適用線区.xlsx が見つかりません。", vbExclamation
        Exit Function
    End If

    Dim cn As Object
    Dim rs As Object
    On Error GoTo ErrorHandler

    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then
        MsgBox "出張所別_単価適用線区.xlsx を参照できませんでした。" & vbCrLf & sourceFilePath, vbExclamation
        Exit Function
    End If

    Dim sql As String
    sql = "SELECT F1, F2, F3, F5 FROM " & _
          BuildAdoSheetRangeName(MASTER_SHEET_NAME, "B", 2, "F", PROJECT_NAME_MASTER_LAST_ROW) & _
          " WHERE F1 IS NOT NULL"
    Set rs = cn.Execute(sql)

    Do Until rs.EOF
        If MasterTextMatches(CommonNzText(CommonGetAdoFieldValue(rs, 0)), request.BranchName) And _
           MasterTextMatches(CommonNzText(CommonGetAdoFieldValue(rs, 1)), request.OfficeName) Then
            masterRow.BranchName = CommonNzText(CommonGetAdoFieldValue(rs, 0))
            masterRow.OfficeName = CommonNzText(CommonGetAdoFieldValue(rs, 1))
            masterRow.BranchGroupName = CommonNzText(CommonGetAdoFieldValue(rs, 2))
            masterRow.UnitPriceSectionName = CommonNzText(CommonGetAdoFieldValue(rs, 3))
            TryLoadUnitPriceMasterRow = True
            GoTo Cleanup
        End If
        rs.MoveNext
    Loop

    MsgBox "単価適用線区シートに該当する支店・出張所が見つかりませんでした。" & vbCrLf & _
           "支店：" & request.BranchName & vbCrLf & _
           "出張所：" & request.OfficeName, vbExclamation

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    MsgBox "単価適用線区データの読み込みに失敗しました。" & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function ResolveUnitPriceSourceFilePath(ByRef request As UnitPriceRequest, _
                                                ByRef masterRow As UnitPriceMasterRow, _
                                                ByRef priceFolderPath As String, _
                                                ByRef sectionFolderPath As String) As String
    priceFolderPath = ResolveUnitPricePriceFolderPath(request, masterRow, sectionFolderPath)
    If priceFolderPath = "" Then Exit Function

    ResolveUnitPriceSourceFilePath = FindUnitPriceWorkbook(priceFolderPath, request.projectName)
    If ResolveUnitPriceSourceFilePath = "" Then
        MsgBox "工事件名に一致する単価表が見つかりません。" & vbCrLf & _
               "工事件名：" & request.projectName & vbCrLf & _
               priceFolderPath, vbExclamation
    End If
End Function

Private Function ResolveUnitPricePriceFolderPath(ByRef request As UnitPriceRequest, _
                                                 ByRef masterRow As UnitPriceMasterRow, _
                                                 ByRef sectionFolderPath As String) As String
    Dim dataRoot As String
    dataRoot = GetUnitPriceDataRootPath()
    If dataRoot = "" Or Dir(dataRoot, vbDirectory) = "" Then
        MsgBox "単価データフォルダが見つかりません。" & vbCrLf & dataRoot, vbExclamation
        Exit Function
    End If

    Dim lineFolder As String
    lineFolder = FindChildFolderByKey(dataRoot, request.lineType, True)
    If lineFolder = "" Then
        MsgBox "線区区分フォルダが見つかりません。" & vbCrLf & request.lineType & vbCrLf & dataRoot, vbExclamation
        Exit Function
    End If

    Dim branchGroupFolder As String
    branchGroupFolder = FindChildFolderByKey(lineFolder, masterRow.BranchGroupName, False)
    If branchGroupFolder = "" Then
        MsgBox "支社フォルダが見つかりません。" & vbCrLf & masterRow.BranchGroupName & vbCrLf & lineFolder, vbExclamation
        Exit Function
    End If

    Dim sectionFolder As String
    sectionFolder = FindChildFolderByKey(branchGroupFolder, masterRow.UnitPriceSectionName, True)
    If sectionFolder = "" Then
        MsgBox "単価適用保線区フォルダが見つかりません。" & vbCrLf & masterRow.UnitPriceSectionName & vbCrLf & branchGroupFolder, vbExclamation
        Exit Function
    End If

    Dim yearFolder As String
    yearFolder = sectionFolder & "\" & request.Nendo
    If Dir(yearFolder, vbDirectory) = "" Then
        MsgBox "年度フォルダが見つかりません。" & vbCrLf & yearFolder, vbExclamation
        Exit Function
    End If

    Dim priceFolder As String
    priceFolder = yearFolder & "\" & ResolveUnitPriceFolderName(request.UnitPriceKind)
    If Dir(priceFolder, vbDirectory) = "" Then
        MsgBox "単価区分フォルダが見つかりません。" & vbCrLf & priceFolder, vbExclamation
        Exit Function
    End If

    sectionFolderPath = sectionFolder
    ResolveUnitPricePriceFolderPath = priceFolder
End Function

Private Function ResolveUnitPriceFolderName(ByVal priceKind As String) As String
    Dim normalizedKind As String
    normalizedKind = NormalizeMatchText(priceKind)
    If InStr(1, normalizedKind, NormalizeMatchText(INITIAL_PRICE_NAME), vbTextCompare) > 0 Then
        ResolveUnitPriceFolderName = NORMAL_PRICE_FOLDER
    ElseIf InStr(1, normalizedKind, NormalizeMatchText(DESIGN_CHANGE_PRICE_NAME), vbTextCompare) > 0 Then
        ResolveUnitPriceFolderName = DESIGN_CHANGE_PRICE_NAME
    ElseIf InStr(1, normalizedKind, NormalizeMatchText(NORMAL_PRICE_FOLDER), vbTextCompare) > 0 Then
        ResolveUnitPriceFolderName = NORMAL_PRICE_FOLDER
    Else
        ResolveUnitPriceFolderName = priceKind
    End If
End Function

Private Function LoadWorksheetNamesFromWorkbook(ByVal sourceFilePath As String) As Collection
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim result As Collection
    Dim previousScreenUpdating As Boolean

    On Error GoTo ErrorHandler
    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    Set result = New Collection
    For Each sourceSheet In sourceBook.Worksheets
        result.Add sourceSheet.Name
    Next sourceSheet
    Set LoadWorksheetNamesFromWorkbook = result

Cleanup:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
    Exit Function

ErrorHandler:
    Set LoadWorksheetNamesFromWorkbook = Nothing
    MsgBox "単価表ブックを開けませんでした。" & vbCrLf & sourceFilePath & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function PromptLineNameSelection(ByVal sheetNames As Collection) As Collection
    On Error GoTo ErrorHandler

    SelectLineName.InitLineNames sheetNames
    SelectLineName.Show vbModal

    If SelectLineName.SelectionConfirmed Then
        Set PromptLineNameSelection = SelectLineName.GetSelectedLineNames()
    End If

Cleanup:
    On Error Resume Next
    Unload SelectLineName
    On Error GoTo 0
    Exit Function

ErrorHandler:
    MsgBox "積算線区選択フォームを表示できませんでした。" & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function ImportSelectedUnitPriceSheets(ByVal sourceFilePath As String, _
                                               ByVal selectedSheetNames As Collection, _
                                               ByVal targetBook As Workbook) As Boolean
    Dim sourceBook As Workbook
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean

    On Error GoTo ErrorHandler
    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    DeleteImportedUnitPriceSheets targetBook

    Dim sheetName As Variant
    For Each sheetName In selectedSheetNames
        DeleteWorksheetIfExists targetBook, CStr(sheetName)
        sourceBook.Worksheets(CStr(sheetName)).Copy After:=targetBook.Worksheets(targetBook.Worksheets.Count)
        With targetBook.Worksheets(targetBook.Worksheets.Count)
            .Name = MakeUniqueWorksheetName(targetBook, CStr(sheetName), .Name)
            .Tab.Color = RGB(221, 235, 247)
            MarkImportedUnitPriceSheet targetBook.Worksheets(.Name)
        End With
    Next sheetName

    ImportSelectedUnitPriceSheets = True

Cleanup:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating
    CommonGetBasicInfoWorksheet(targetBook).Activate
    On Error GoTo 0
    Exit Function

ErrorHandler:
    MsgBox "単価表の取り込みに失敗しました。" & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function ImportAndMergePurchaseUnitPriceSheets(ByVal sourceFilePath As String, _
                                                       ByVal selectedSheetNames As Collection, _
                                                       ByVal targetBook As Workbook, _
                                                       ByVal newSheetName As String) As Boolean
    Dim sourceBook As Workbook
    Dim newSheet As Worksheet
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    On Error GoTo ErrorHandler
    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    DeleteWorksheetIfExists targetBook, newSheetName
    Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)

    Dim isFirst As Boolean
    isFirst = True

    Dim sheetName As Variant
    For Each sheetName In selectedSheetNames
        Dim srcSheet As Worksheet
        Set srcSheet = sourceBook.Worksheets(CStr(sheetName))

        If isFirst Then
            srcSheet.Copy After:=targetBook.Worksheets(targetBook.Worksheets.Count)
            Set newSheet = targetBook.Worksheets(targetBook.Worksheets.Count)
            newSheet.Name = MakeUniqueWorksheetName(targetBook, newSheetName, newSheet.Name)
            isFirst = False
        Else
            AppendSheetDataExcludingHeader srcSheet, newSheet
        End If
    Next sheetName

    If Not newSheet Is Nothing Then
        newSheet.Tab.Color = RGB(PURCHASE_SHEET_TAB_R, PURCHASE_SHEET_TAB_G, PURCHASE_SHEET_TAB_B)
        MarkImportedUnitPriceSheet newSheet
    End If

    ImportAndMergePurchaseUnitPriceSheets = True

Cleanup:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
    Application.CutCopyMode = False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    MsgBox "購入充当単価表の取り込みに失敗しました。" & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Function ImportAndMergeWeldingUnitPriceSheets(ByVal sourceFilePath As String, _
                                                      ByVal selectedSheetNames As Collection, _
                                                      ByVal targetBook As Workbook, _
                                                      ByVal newSheetName As String) As Boolean
    Dim sourceBook As Workbook
    Dim newSheet As Worksheet
    Dim previousDisplayAlerts As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousCalculation As XlCalculation

    On Error GoTo ErrorHandler
    previousDisplayAlerts = Application.DisplayAlerts
    previousScreenUpdating = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    DeleteWorksheetIfExists targetBook, newSheetName
    Set sourceBook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)

    Dim isFirst As Boolean
    isFirst = True

    Dim sheetName As Variant
    For Each sheetName In selectedSheetNames
        Dim srcSheet As Worksheet
        Set srcSheet = sourceBook.Worksheets(CStr(sheetName))

        If isFirst Then
            srcSheet.Copy After:=targetBook.Worksheets(targetBook.Worksheets.Count)
            Set newSheet = targetBook.Worksheets(targetBook.Worksheets.Count)
            newSheet.Name = MakeUniqueWorksheetName(targetBook, newSheetName, newSheet.Name)
            isFirst = False
        Else
            AppendSheetDataExcludingHeader srcSheet, newSheet
        End If
    Next sheetName

    If Not newSheet Is Nothing Then
        newSheet.Tab.Color = RGB(WELDING_SHEET_TAB_R, WELDING_SHEET_TAB_G, WELDING_SHEET_TAB_B)
        MarkImportedUnitPriceSheet newSheet
    End If

    ImportAndMergeWeldingUnitPriceSheets = True

Cleanup:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating
    Application.Calculation = previousCalculation
    Application.CutCopyMode = False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    MsgBox "レール溶接単価表の取り込みに失敗しました。" & vbCrLf & Err.Description, vbExclamation
    Resume Cleanup
End Function

Private Sub AppendSheetDataExcludingHeader(ByVal srcSheet As Worksheet, ByVal destSheet As Worksheet)
    Dim srcUsed As Range
    Set srcUsed = srcSheet.UsedRange
    If srcUsed Is Nothing Then Exit Sub

    Dim srcLastRow As Long, srcLastCol As Long
    srcLastRow = srcUsed.Row + srcUsed.Rows.Count - 1
    srcLastCol = srcUsed.Column + srcUsed.Columns.Count - 1
    If srcLastRow <= SOURCE_HEADER_ROW Then Exit Sub

    Dim destLastRow As Long
    destLastRow = FindLastUsedRow(destSheet)
    If destLastRow < 1 Then destLastRow = 1

    srcSheet.Range(srcSheet.Cells(SOURCE_HEADER_ROW + 1, 1), srcSheet.Cells(srcLastRow, srcLastCol)).Copy
    destSheet.Cells(destLastRow + 1, 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
End Sub

Private Function FindLastUsedRow(ByVal targetSheet As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = targetSheet.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If foundCell Is Nothing Then FindLastUsedRow = 0 Else FindLastUsedRow = foundCell.Row
End Function

Private Function ImportPurchaseUnitPriceSheetsByReference(ByRef request As UnitPriceRequest, _
                                                          ByRef masterRow As UnitPriceMasterRow, _
                                                          ByVal priceFolderPath As String, _
                                                          ByVal sectionFolderPath As String, _
                                                          ByVal targetBook As Workbook, _
                                                          ByRef createdSheetName As String) As Boolean
    createdSheetName = ""

    Dim purchaseFilePath As String
    purchaseFilePath = FindPurchaseUnitPriceWorkbook(priceFolderPath)
    If purchaseFilePath = "" Then Exit Function

    Dim referenceKey As String
    referenceKey = BuildPurchaseReferenceKey(sectionFolderPath)
    If referenceKey = "" Then
        MsgBox "購入充当単価表の参照キーを取得できませんでした。" & vbCrLf & sectionFolderPath, vbExclamation
        Exit Function
    End If

    Dim purchaseSheetNames As Collection
    Set purchaseSheetNames = LoadPurchaseSheetNamesByReference(purchaseFilePath, referenceKey)
    If purchaseSheetNames Is Nothing Then Exit Function
    If purchaseSheetNames.Count = 0 Then
        MsgBox "購入充当単価表に参照キー「" & referenceKey & "」に一致するシートが見つかりません。" & vbCrLf & purchaseFilePath, vbExclamation
        Exit Function
    End If

    Dim newSheetName As String
    newSheetName = BuildPurchaseSheetName(GetPathBaseName(sectionFolderPath))
    If newSheetName = "" Then newSheetName = BuildPurchaseSheetName(masterRow.UnitPriceSectionName)
    If newSheetName = "" Then
        MsgBox "購入充当単価表の取込先シート名を生成できませんでした。", vbExclamation
        Exit Function
    End If

    If Not ImportAndMergePurchaseUnitPriceSheets(purchaseFilePath, purchaseSheetNames, targetBook, newSheetName) Then Exit Function

    createdSheetName = newSheetName
    ImportPurchaseUnitPriceSheetsByReference = True
End Function

Private Function ImportWeldingUnitPriceSheetsIfRequired(ByVal wsInfo As Worksheet, _
                                                        ByRef masterRow As UnitPriceMasterRow, _
                                                        ByVal priceFolderPath As String, _
                                                        ByVal sectionFolderPath As String, _
                                                        ByVal selectedLineNames As Collection, _
                                                        ByVal targetBook As Workbook, _
                                                        ByRef createdSheetName As String) As Boolean
    createdSheetName = ""
    If Not IsWeldingUnitPriceRequired(wsInfo) Then
        ImportWeldingUnitPriceSheetsIfRequired = True
        Exit Function
    End If

    Dim weldingFilePath As String
    weldingFilePath = FindWeldingUnitPriceWorkbook(priceFolderPath)
    If weldingFilePath = "" Then
        MsgBox "レール溶接単価表が見つかりません。" & vbCrLf & _
               "検索値：" & WELDING_FILE_KEYWORD & vbCrLf & priceFolderPath, vbExclamation
        Exit Function
    End If

    Dim weldingSheetNames As Collection
    Set weldingSheetNames = LoadWeldingSheetNames(weldingFilePath, selectedLineNames)
    If weldingSheetNames Is Nothing Then Exit Function
    If weldingSheetNames.Count = 0 Then
        MsgBox "レール溶接単価表に取り込み可能なシートが見つかりません。" & vbCrLf & weldingFilePath, vbExclamation
        Exit Function
    End If

    Dim newSheetName As String
    newSheetName = BuildWeldingSheetName(GetPathBaseName(sectionFolderPath))
    If newSheetName = "" Then newSheetName = BuildWeldingSheetName(masterRow.UnitPriceSectionName)
    If newSheetName = "" Then
        MsgBox "レール溶接単価表の取込先シート名を生成できませんでした。", vbExclamation
        Exit Function
    End If

    If Not ImportAndMergeWeldingUnitPriceSheets(weldingFilePath, weldingSheetNames, targetBook, newSheetName) Then Exit Function

    createdSheetName = newSheetName
    ImportWeldingUnitPriceSheetsIfRequired = True
End Function

Private Function IsWeldingUnitPriceRequired(ByVal wsInfo As Worksheet) As Boolean
    If wsInfo Is Nothing Then Exit Function
    IsWeldingUnitPriceRequired = (StrComp(NormalizeMatchText(CStr(wsInfo.Range(BASIC_INFO_WELDING_FLAG_CELL).Value)), _
                                         NormalizeMatchText(WELDING_REQUIRED_VALUE), vbTextCompare) = 0)
End Function

Private Function LoadWeldingSheetNames(ByVal sourceFilePath As String, _
                                       ByVal selectedLineNames As Collection) As Collection
    Dim sourceSheetNames As Collection
    Set sourceSheetNames = LoadWorksheetNamesFromWorkbook(sourceFilePath)
    If sourceSheetNames Is Nothing Then Exit Function

    Dim result As Collection
    Set result = New Collection

    Dim sheetName As Variant
    If Not selectedLineNames Is Nothing Then
        For Each sheetName In sourceSheetNames
            If CollectionContainsText(selectedLineNames, CStr(sheetName)) Then result.Add CStr(sheetName)
        Next sheetName
    End If

    If result.Count = 0 Then
        For Each sheetName In sourceSheetNames
            result.Add CStr(sheetName)
        Next sheetName
    End If

    Set LoadWeldingSheetNames = result
End Function

Private Function LoadPurchaseSheetNamesByReference(ByVal sourceFilePath As String, ByVal referenceKey As String) As Collection
    Dim sheetNames As Collection
    Set sheetNames = LoadWorksheetNamesFromWorkbook(sourceFilePath)
    If sheetNames Is Nothing Then Exit Function

    Dim result As Collection
    Set result = New Collection
    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If PurchaseSheetNameMatchesReferenceKey(CStr(sheetName), referenceKey) Then result.Add CStr(sheetName)
    Next sheetName
    Set LoadPurchaseSheetNamesByReference = result
End Function

Private Function PurchaseSheetNameMatchesReferenceKey(ByVal sheetName As String, ByVal referenceKey As String) As Boolean
    Dim a As String, b As String
    a = NormalizeMatchText(sheetName)
    b = NormalizeMatchText(referenceKey)
    If a = "" Or b = "" Then Exit Function
    PurchaseSheetNameMatchesReferenceKey = (a = b Or _
                                            Left$(a, Len(b) + 1) = b & "-" Or _
                                            Left$(a, Len(b) + 1) = b & "_")
End Function

Private Function BuildPurchaseReferenceKey(ByVal sectionFolderPath As String) As String
    BuildPurchaseReferenceKey = ExtractLeadingDigits(GetPathBaseName(sectionFolderPath))
End Function

Private Function ExtractLeadingDigits(ByVal value As String) As String
    Dim result As String, i As Long, ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch Like "[0-9]" Then
            result = result & ch
        ElseIf Len(result) > 0 Then
            Exit For
        End If
    Next i
    ExtractLeadingDigits = result
End Function

Private Function GetPathBaseName(ByVal sourcePath As String) As String
    Dim pos As Long
    pos = InStrRev(sourcePath, "\")
    If pos > 0 Then GetPathBaseName = Mid$(sourcePath, pos + 1) Else GetPathBaseName = sourcePath
End Function

Private Function BuildPurchaseSheetName(ByVal rawSectionName As String) As String
    Dim trimmed As String
    trimmed = TrimLeadingDigitsAndSeparators(rawSectionName)
    If trimmed = "" Then Exit Function
    BuildPurchaseSheetName = Left$(MakeSafeWorksheetName(trimmed & PURCHASE_SHEET_NAME_SUFFIX), 31)
End Function

Private Function BuildWeldingSheetName(ByVal rawSectionName As String) As String
    Dim trimmed As String
    trimmed = TrimLeadingDigitsAndSeparators(rawSectionName)
    If trimmed = "" Then Exit Function
    BuildWeldingSheetName = Left$(MakeSafeWorksheetName(trimmed & WELDING_SHEET_NAME_SUFFIX), 31)
End Function

Private Function TrimLeadingDigitsAndSeparators(ByVal rawValue As String) As String
    Dim result As String
    result = Trim$(CommonNormalizeText(rawValue))
    Do While Len(result) > 0
        If Mid$(result, 1, 1) Like "[0-9]" Then result = Mid$(result, 2) Else Exit Do
    Loop
    Do While Len(result) > 0 And (Left$(result, 1) = "_" Or Left$(result, 1) = "-" Or Left$(result, 1) = " ")
        result = Mid$(result, 2)
    Loop
    TrimLeadingDigitsAndSeparators = Trim$(result)
End Function

Private Sub WriteSelectedLineNames(ByVal wsInfo As Worksheet, ByVal selectedSheetNames As Collection)
    If wsInfo Is Nothing Then Exit Sub
    wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).Value = JoinCollectionText(selectedSheetNames, ChrW$(&H3001))
    FormatImportedLineNamesCell wsInfo
End Sub

Public Sub FormatImportedLineNamesCell(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim targetCell As Range
    Set targetCell = wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).MergeArea.Cells(1, 1)

    Dim formattedText As String
    formattedText = NormalizeImportedLineNameText(CStr(targetCell.Value))
    If CStr(targetCell.Value) <> formattedText Then targetCell.Value = formattedText

    targetCell.MergeArea.WrapText = True
    targetCell.MergeArea.VerticalAlignment = xlVAlignCenter

    Dim lineCount As Long
    lineCount = CountImportedLineNameLines(formattedText)
    If lineCount < 1 Then lineCount = 1
    If lineCount > IMPORTED_LINE_NAME_MAX_LINES Then lineCount = IMPORTED_LINE_NAME_MAX_LINES
    targetCell.EntireRow.RowHeight = IMPORTED_LINE_NAME_BASE_ROW_HEIGHT * lineCount
End Sub

Private Function NormalizeImportedLineNameText(ByVal sourceText As String) As String
    Dim normalized As String
    normalized = Replace$(Replace$(sourceText, vbCrLf, vbLf), vbCr, vbLf)
    normalized = Replace$(normalized, ChrW$(&H3001), vbLf)

    Dim parts As Variant
    parts = Split(normalized, vbLf)

    Dim result As String
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim lineText As String
        lineText = Trim$(CStr(parts(i)))
        If lineText <> "" Then
            If result <> "" Then result = result & vbLf
            result = result & lineText
        End If
    Next i
    NormalizeImportedLineNameText = result
End Function

Private Function CountImportedLineNameLines(ByVal lineNameText As String) As Long
    If lineNameText = "" Then
        CountImportedLineNameLines = 1
        Exit Function
    End If

    Dim i As Long
    CountImportedLineNameLines = 1
    For i = 1 To Len(lineNameText)
        If Mid$(lineNameText, i, 1) = vbLf Then CountImportedLineNameLines = CountImportedLineNameLines + 1
    Next i
End Function

Private Function JoinCollectionText(ByVal values As Collection, ByVal delimiter As String) As String
    If values Is Nothing Then Exit Function
    Dim result As String, item As Variant
    For Each item In values
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(item)
    Next item
    JoinCollectionText = result
End Function

Public Sub ConfirmAndClearUnitPriceForBasicInfo(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    If MsgBox("単価情報をクリアしますか？" & vbCrLf & _
              "はい：C24の選択内容と、作成済みの単価シートを削除します。" & vbCrLf & _
              "いいえ：単価情報を残します。", vbQuestion + vbYesNo, "単価情報クリア") <> vbYes Then Exit Sub
    ClearUnitPriceSheets wsInfo.Parent
    wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).ClearContents
    FormatImportedLineNamesCell wsInfo
End Sub

Public Sub SilentClearUnitPriceForBasicInfo(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    ClearUnitPriceSheets wsInfo.Parent
    wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).ClearContents
    FormatImportedLineNamesCell wsInfo
End Sub

Public Sub ClearUnitPriceSheets(Optional ByVal targetBook As Workbook)
    Dim savedErrNumber As Long, savedErrSource As String, savedErrDescription As String
    If targetBook Is Nothing Then Set targetBook = ThisWorkbook
    Dim previousDisplayAlerts As Boolean
    previousDisplayAlerts = Application.DisplayAlerts
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    DeleteImportedUnitPriceSheets targetBook
Cleanup:
    Application.DisplayAlerts = previousDisplayAlerts
    If savedErrNumber <> 0 Then Err.Raise savedErrNumber, savedErrSource, savedErrDescription
    Exit Sub
ErrorHandler:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    Resume Cleanup
End Sub

Private Sub DeleteImportedUnitPriceSheets(ByVal targetBook As Workbook)
    Dim i As Long
    For i = targetBook.Worksheets.Count To 1 Step -1
        If IsImportedUnitPriceSheet(targetBook.Worksheets(i)) Then
            If targetBook.Worksheets.Count > 1 Then targetBook.Worksheets(i).Delete
        End If
    Next i
End Sub

Private Sub DeleteWorksheetIfExists(ByVal targetBook As Workbook, ByVal sheetName As String)
    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = targetBook.Worksheets(sheetName)
    On Error GoTo 0
    If targetSheet Is Nothing Then Exit Sub
    If targetBook.Worksheets.Count <= 1 Then Exit Sub
    targetSheet.Delete
End Sub

Private Function IsImportedUnitPriceSheet(ByVal targetSheet As Worksheet) As Boolean
    On Error Resume Next
    Dim prop As Object
    For Each prop In targetSheet.CustomProperties
        If StrComp(prop.Name, IMPORTED_SHEET_PROPERTY, vbTextCompare) = 0 Then
            IsImportedUnitPriceSheet = (CStr(prop.Value) = "1")
            Exit Function
        End If
    Next prop
    On Error GoTo 0
End Function

Private Sub MarkImportedUnitPriceSheet(ByVal targetSheet As Worksheet)
    On Error Resume Next
    targetSheet.CustomProperties.Add Name:=IMPORTED_SHEET_PROPERTY, Value:="1"
    On Error GoTo 0
End Sub

Private Function MakeUniqueWorksheetName(ByVal targetBook As Workbook, _
                                         ByVal requestedName As String, _
                                         ByVal currentName As String) As String
    Dim baseName As String
    baseName = Left$(MakeSafeWorksheetName(requestedName), 31)
    If baseName = "" Then baseName = "単価表"
    If StrComp(baseName, currentName, vbTextCompare) = 0 Then
        MakeUniqueWorksheetName = currentName
        Exit Function
    End If
    If Not WorksheetExists(targetBook, baseName) Then
        MakeUniqueWorksheetName = baseName
        Exit Function
    End If
    Dim i As Long, candidate As String
    For i = 2 To 99
        candidate = Left$(baseName, 28) & "(" & CStr(i) & ")"
        If Not WorksheetExists(targetBook, candidate) Then
            MakeUniqueWorksheetName = candidate
            Exit Function
        End If
    Next i
    MakeUniqueWorksheetName = currentName
End Function

Private Function MakeSafeWorksheetName(ByVal sourceName As String) As String
    Dim result As String
    result = Trim$(sourceName)
    Dim i As Long
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", "[", "]")
    For i = LBound(invalidChars) To UBound(invalidChars)
        result = Replace$(result, CStr(invalidChars(i)), "_")
    Next i
    MakeSafeWorksheetName = result
End Function

Private Function WorksheetExists(ByVal targetBook As Workbook, ByVal sheetName As String) As Boolean
    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = targetBook.Worksheets(sheetName)
    WorksheetExists = Not targetSheet Is Nothing
    On Error GoTo 0
End Function

Private Function LoadUnitPriceProjectNamesForBasicInfo(ByVal sourceFilePath As String, _
                                                       ByVal lineType As String) As Collection
    Select Case CommonNormalizeText(lineType)
        Case ZAIRAISEN_NAME
            Set LoadUnitPriceProjectNamesForBasicInfo = LoadUnitPriceProjectNames(sourceFilePath, ZAIRAISEN_NAME)
        Case SHINKANSEN_NAME
            Set LoadUnitPriceProjectNamesForBasicInfo = LoadUnitPriceProjectNames(sourceFilePath, SHINKANSEN_NAME)
        Case Else
            Dim result As Collection
            Set result = New Collection
            AddUnitPriceProjectNames result, LoadUnitPriceProjectNames(sourceFilePath, ZAIRAISEN_NAME)
            AddUnitPriceProjectNames result, LoadUnitPriceProjectNames(sourceFilePath, SHINKANSEN_NAME)
            Set LoadUnitPriceProjectNamesForBasicInfo = result
    End Select
End Function

Private Sub AddUnitPriceProjectNames(ByVal target As Collection, ByVal source As Collection)
    If target Is Nothing Or source Is Nothing Then Exit Sub
    Dim item As Variant
    For Each item In source
        If Not CollectionContainsText(target, CStr(item)) Then target.Add CStr(item)
    Next item
End Sub

Private Function LoadUnitPriceProjectNames(ByVal sourceFilePath As String, _
                                           ByVal lineType As String) As Collection
    Dim sheetName As String, endRow As Long
    On Error GoTo ErrorHandler
    ResolveUnitPriceProjectNameMasterConfig lineType, sheetName, endRow
    If sheetName = "" Then Exit Function

    If CommonNormalizeText(lineType) = ZAIRAISEN_NAME Then
        Dim result As Collection
        Set result = New Collection
        AddUnitPriceProjectNames result, LoadUnitPriceProjectNamesByAdo(sourceFilePath, sheetName, ZAIRAISEN_PROJECT_NAME_FIRST_END_ROW)
        AddUnitPriceProjectNames result, LoadUnitPriceProjectNamesByAdo2(sourceFilePath, sheetName, ZAIRAISEN_PROJECT_NAME_SKIP_ROW + 1, endRow)
        Set LoadUnitPriceProjectNames = result
    Else
        Set LoadUnitPriceProjectNames = LoadUnitPriceProjectNamesByAdo(sourceFilePath, sheetName, endRow)
    End If
    Exit Function
ErrorHandler:
    Set LoadUnitPriceProjectNames = Nothing
End Function

Private Sub ResolveUnitPriceProjectNameMasterConfig(ByVal lineType As String, _
                                                    ByRef sheetName As String, _
                                                    ByRef endRow As Long)
    Select Case CommonNormalizeText(lineType)
        Case ZAIRAISEN_NAME
            sheetName = PROJECT_MASTER_SHEET_NAME
            endRow = ZAIRAISEN_PROJECT_NAME_SECOND_END_ROW
        Case SHINKANSEN_NAME
            sheetName = SHINKANSEN_NAME & "_" & PROJECT_MASTER_SHEET_NAME
            endRow = SHINKANSEN_PROJECT_NAME_END_ROW
        Case Else
            sheetName = ""
            endRow = 0
    End Select
End Sub

Private Function LoadUnitPriceProjectNamesByAdo(ByVal sourceFilePath As String, _
                                                ByVal sheetName As String, _
                                                ByVal endRow As Long) As Collection
    Dim cn As Object, rs As Object, result As Collection
    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then Exit Function

    Set rs = cn.Execute("SELECT F1 FROM " & _
                        BuildAdoSheetRangeName(sheetName, "A", PROJECT_NAME_MASTER_START_ROW, "A", endRow) & _
                        " WHERE F1 IS NOT NULL")
    Set result = New Collection
    Do Until rs.EOF
        Dim itemText As String
        itemText = Trim$(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
        If itemText <> "" Then result.Add itemText
        rs.MoveNext
    Loop
    Set LoadUnitPriceProjectNamesByAdo = result
Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function
ErrorHandler:
    Set LoadUnitPriceProjectNamesByAdo = Nothing
    Resume Cleanup
End Function

Private Function LoadUnitPriceProjectNamesByAdo2(ByVal sourceFilePath As String, _
                                                 ByVal sheetName As String, _
                                                 ByVal startRow As Long, _
                                                 ByVal endRow As Long) As Collection
    Dim cn As Object, rs As Object, result As Collection
    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then Exit Function

    Set rs = cn.Execute("SELECT F1 FROM " & _
                        BuildAdoSheetRangeName(sheetName, "A", startRow, "A", endRow) & _
                        " WHERE F1 IS NOT NULL")
    Set result = New Collection
    Do Until rs.EOF
        Dim itemText As String
        itemText = Trim$(CommonNzText(CommonGetAdoFieldValue(rs, 0)))
        If itemText <> "" Then result.Add itemText
        rs.MoveNext
    Loop
    Set LoadUnitPriceProjectNamesByAdo2 = result
Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function
ErrorHandler:
    Set LoadUnitPriceProjectNamesByAdo2 = Nothing
    Resume Cleanup
End Function

Private Function BuildAdoSheetRangeName(ByVal sheetName As String, _
                                        ByVal startCol As String, ByVal startRow As Long, _
                                        ByVal endCol As String, ByVal endRow As Long) As String
    BuildAdoSheetRangeName = "[" & Replace$(sheetName, "]", "]]") & "$" & _
                             startCol & CStr(startRow) & ":" & endCol & CStr(endRow) & "]"
End Function

Private Sub WriteUnitPriceLineTypeValidation(ByVal wsInfo As Worksheet)
    Dim listRange As Range
    Set listRange = wsInfo.Range(LINE_TYPE_LIST_COL & LIST_START_ROW).Resize(2, 1)
    wsInfo.Columns(LINE_TYPE_LIST_COL & ":" & LINE_TYPE_LIST_COL).Hidden = False
    listRange.ClearContents
    listRange.Cells(1, 1).Value = ZAIRAISEN_NAME
    listRange.Cells(2, 1).Value = SHINKANSEN_NAME
    ResetUnitPriceValidation wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL), listRange
    wsInfo.Columns(LINE_TYPE_LIST_COL & ":" & LINE_TYPE_LIST_COL).Hidden = True
End Sub

Private Sub WriteUnitPriceKindValidation(ByVal wsInfo As Worksheet)
    Dim listRange As Range
    Set listRange = wsInfo.Range(PRICE_KIND_LIST_COL & LIST_START_ROW).Resize(2, 1)
    wsInfo.Columns(PRICE_KIND_LIST_COL & ":" & PRICE_KIND_LIST_COL).Hidden = False
    listRange.ClearContents
    listRange.Cells(1, 1).Value = INITIAL_PRICE_NAME
    listRange.Cells(2, 1).Value = DESIGN_CHANGE_PRICE_NAME
    ResetUnitPriceValidation wsInfo.Range(BASIC_INFO_PRICE_KIND_CELL), listRange
    wsInfo.Columns(PRICE_KIND_LIST_COL & ":" & PRICE_KIND_LIST_COL).Hidden = True
End Sub

Private Sub WriteUnitPriceProjectNameValidation(ByVal wsInfo As Worksheet, _
                                                ByVal projectNames As Collection, _
                                                ByVal keepProjectName As Boolean)
    Dim maxRows As Long
    maxRows = Application.Max(1, projectNames.Count)
    Dim listRange As Range
    Set listRange = wsInfo.Range(PROJECT_NAME_LIST_COL & LIST_START_ROW).Resize(maxRows, 1)

    Dim currentProjectName As String
    currentProjectName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).Value))

    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = False
    wsInfo.Range(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).ClearContents

    Dim i As Long
    For i = 1 To projectNames.Count
        listRange.Cells(i, 1).Value = projectNames(i)
    Next i

    ResetUnitPriceValidation wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL), _
                             listRange.Resize(Application.Max(1, projectNames.Count))

    If Not keepProjectName Or currentProjectName = "" Or Not CollectionContainsText(projectNames, currentProjectName) Then
        wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).ClearContents
    End If

    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = True
End Sub

'--------------------------------------------------------------------------
'  ResetUnitPriceValidation  (#17, #18 修正)
'    セル結合されている場合、MergeArea の左上セルに対して入力規則を設定。
'    targetCell が D22 等の結合副セルを指してしまう場合でも
'    確実に結合範囲の代表セル（C22 等）に適用する。
'    AlertStyle = xlValidAlertInformation：非表示列参照時に
'    xlValidAlertStop だと入力がブロックされるため。
'--------------------------------------------------------------------------
Private Sub ResetUnitPriceValidation(ByVal targetCell As Range, ByVal listRange As Range)
    ' 結合セルの左上（代表セル）に対して入力規則を設定する
    Dim topLeft As Range
    Set topLeft = targetCell.MergeArea.Cells(1, 1)
    With topLeft.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, _
             Formula1:="=" & listRange.Address(True, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With
End Sub

Private Sub ClearUnitPriceProjectNameValidation(ByVal wsInfo As Worksheet, _
                                                Optional ByVal clearProjectName As Boolean = True)
    On Error Resume Next
    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = False
    wsInfo.Range(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).ClearContents
    wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).Validation.Delete
    If clearProjectName Then wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).ClearContents
    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = True
    On Error GoTo 0
End Sub

Private Function CollectionContainsText(ByVal values As Collection, ByVal searchText As String) As Boolean
    If values Is Nothing Then Exit Function
    Dim item As Variant
    For Each item In values
        If StrComp(CommonNormalizeText(CStr(item)), CommonNormalizeText(searchText), vbTextCompare) = 0 Then
            CollectionContainsText = True
            Exit Function
        End If
    Next item
End Function

Private Function FindUnitPriceWorkbook(ByVal folderPath As String, ByVal projectName As String) As String
    Dim normalizedProjectName As String
    normalizedProjectName = NormalizeMatchText(projectName)
    If normalizedProjectName = "" Then Exit Function

    Dim extensions As Variant
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")
    Dim ext As Variant, fileName As String
    For Each ext In extensions
        fileName = Dir(folderPath & "\" & CStr(ext), vbNormal)
        Do While fileName <> ""
            If Left$(fileName, 2) <> "~$" Then
                If InStr(1, fileName, UNIT_PRICE_FILE_KEYWORD, vbTextCompare) = 0 Then
                    If InStr(1, NormalizeMatchText(RemoveFileExtension(fileName)), normalizedProjectName, vbTextCompare) > 0 Then
                        FindUnitPriceWorkbook = folderPath & "\" & fileName
                        Exit Function
                    End If
                End If
            End If
            fileName = Dir()
        Loop
    Next ext
End Function

Private Function FindPurchaseUnitPriceWorkbook(ByVal folderPath As String) As String
    Dim extensions As Variant
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")
    Dim ext As Variant, fileName As String
    For Each ext In extensions
        fileName = Dir(folderPath & "\" & CStr(ext), vbNormal)
        Do While fileName <> ""
            If Left$(fileName, 2) <> "~$" Then
                If InStr(1, fileName, UNIT_PRICE_FILE_KEYWORD, vbTextCompare) > 0 Then
                    FindPurchaseUnitPriceWorkbook = folderPath & "\" & fileName
                    Exit Function
                End If
            End If
            fileName = Dir()
        Loop
    Next ext
End Function

Private Function FindWeldingUnitPriceWorkbook(ByVal folderPath As String) As String
    FindWeldingUnitPriceWorkbook = FindWorkbookByKeyword(folderPath, WELDING_FILE_KEYWORD)
    If FindWeldingUnitPriceWorkbook = "" Then
        FindWeldingUnitPriceWorkbook = FindWorkbookByKeyword(folderPath, WELDING_FILE_FALLBACK_KEYWORD)
    End If
End Function

Private Function FindWorkbookByKeyword(ByVal folderPath As String, ByVal keyword As String) As String
    Dim normalizedKeyword As String
    normalizedKeyword = NormalizeMatchText(keyword)
    If normalizedKeyword = "" Then Exit Function

    Dim extensions As Variant
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")
    Dim ext As Variant, fileName As String
    For Each ext In extensions
        fileName = Dir(folderPath & "\" & CStr(ext), vbNormal)
        Do While fileName <> ""
            If Left$(fileName, 2) <> "~$" Then
                If InStr(1, NormalizeMatchText(fileName), normalizedKeyword, vbTextCompare) > 0 Then
                    FindWorkbookByKeyword = folderPath & "\" & fileName
                    Exit Function
                End If
            End If
            fileName = Dir()
        Loop
    Next ext
End Function

Private Function RemoveFileExtension(ByVal fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then RemoveFileExtension = Left$(fileName, pos - 1) Else RemoveFileExtension = fileName
End Function

Private Function IsPurchaseUnitPriceProjectName(ByVal projectName As String) As Boolean
    IsPurchaseUnitPriceProjectName = (InStr(1, NormalizeMatchText(projectName), NormalizeMatchText(UNIT_PRICE_FILE_KEYWORD), vbTextCompare) > 0 Or _
                                      InStr(1, NormalizeMatchText(projectName), NormalizeMatchText("購入充当"), vbTextCompare) > 0)
End Function

Private Function FindChildFolderByKey(ByVal parentFolder As String, _
                                      ByVal keyText As String, _
                                      ByVal ignoreLeadingNumbers As Boolean) As String
    Dim folderName As String
    folderName = Dir(parentFolder & "\*", vbDirectory)
    Do While folderName <> ""
        If folderName <> "." And folderName <> ".." Then
            If (GetAttr(parentFolder & "\" & folderName) And vbDirectory) = vbDirectory Then
                If FolderTextMatches(keyText, NormalizeFolderName(folderName, ignoreLeadingNumbers)) Then
                    FindChildFolderByKey = parentFolder & "\" & folderName
                    Exit Function
                End If
            End If
        End If
        folderName = Dir()
    Loop
End Function

Private Function FolderTextMatches(ByVal keyText As String, ByVal folderText As String) As Boolean
    Dim a As String, b As String
    a = NormalizeMatchText(keyText)
    b = NormalizeMatchText(folderText)
    If a = "" Or b = "" Then Exit Function
    FolderTextMatches = (StrComp(a, b, vbTextCompare) = 0 Or _
                         Left$(a, Len(b)) = b Or Left$(b, Len(a)) = a Or _
                         InStr(1, a, b, vbTextCompare) > 0 Or InStr(1, b, a, vbTextCompare) > 0)
End Function

Private Function MasterTextMatches(ByVal masterText As String, ByVal searchText As String) As Boolean
    Dim a As String, b As String
    a = NormalizeMatchText(masterText)
    b = NormalizeMatchText(searchText)
    If a = "" Or b = "" Then Exit Function
    MasterTextMatches = (StrComp(a, b, vbTextCompare) = 0 Or Left$(a, Len(b)) = b)
End Function

Private Function NormalizeMatchText(ByVal value As String) As String
    NormalizeMatchText = CommonRemoveAllSpaces(CommonNormalizeText(value))
End Function

Private Function NormalizeFolderName(ByVal folderName As String, ByVal ignoreLeadingNumbers As Boolean) As String
    Dim result As String
    result = folderName
    If ignoreLeadingNumbers Then
        Do While Len(result) > 0 And Mid$(result, 1, 1) Like "[0-9]"
            result = Mid$(result, 2)
        Loop
        Do While Len(result) > 0 And (Left$(result, 1) = "_" Or Left$(result, 1) = "-" Or Left$(result, 1) = " ")
            result = Mid$(result, 2)
        Loop
    End If
    NormalizeFolderName = result
End Function

Private Function GetUnitPriceDataRootPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim masterFilePath As String
    masterFilePath = GetMasterFilePath()
    If masterFilePath <> "" Then
        GetUnitPriceDataRootPath = fso.BuildPath(fso.GetParentFolderName(fso.GetParentFolderName(masterFilePath)), UNIT_PRICE_DATA_FOLDER)
        Exit Function
    End If
    If Len(ThisWorkbook.Path) > 0 Then
        GetUnitPriceDataRootPath = fso.BuildPath(fso.GetParentFolderName(ThisWorkbook.Path), _
                                   UnitPriceMasterFolderText() & "\" & UNIT_PRICE_DATA_FOLDER)
    End If
End Function

Private Function FirstExistingFilePath(ByVal candidates As Collection) As String
    Dim candidate As Variant, found As Boolean
    For Each candidate In candidates
        found = False
        On Error Resume Next
        found = (Len(Dir(CStr(candidate), vbNormal)) > 0)
        On Error GoTo 0
        If found Then
            FirstExistingFilePath = CStr(candidate)
            Exit Function
        End If
    Next candidate
End Function

Private Function BuildImportCompleteMessage(ByVal selectedSheetNames As Collection, _
                                            ByVal sourceFilePath As String, _
                                            ByVal purchaseSheetName As String, _
                                            ByVal weldingSheetName As String) As String
    BuildImportCompleteMessage = CStr(selectedSheetNames.Count) & "件の積算線区単価表を取り込みました。" & vbCrLf & sourceFilePath
    If Len(purchaseSheetName) > 0 Then
        BuildImportCompleteMessage = BuildImportCompleteMessage & vbCrLf & "購入充当単価表を「" & purchaseSheetName & "」シートに作成しました。"
    End If
    If Len(weldingSheetName) > 0 Then
        BuildImportCompleteMessage = BuildImportCompleteMessage & vbCrLf & "レール溶接単価表を「" & weldingSheetName & "」シートに作成しました。"
    End If
End Function

Private Function OrderInvoiceDocumentFolderText() As String
    Static cached As String
    If cached = "" Then cached = "線路出張所用_注文書_請求書アクセスサイト - ドキュメント"
    OrderInvoiceDocumentFolderText = cached
End Function

Private Function UnitPriceMasterFolderText() As String
    Static cached As String
    If cached = "" Then cached = "単価マスタ"
    UnitPriceMasterFolderText = cached
End Function
