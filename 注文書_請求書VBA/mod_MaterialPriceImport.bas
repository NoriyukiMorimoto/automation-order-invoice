Option Explicit

'==========================================================================
'  工事単価インポートモジュール
'    基本情報 B4/B6/C6/C25/C26/C27 から条件を読み取り、
'    単価マスタ配下の通常単価/設計変更単価ブックを探す。
'    対象ブックのシート名を SelectLineName フォームへ表示し、
'    選択された積算線区シートをこのブックへコピーする。
'==========================================================================

Public SharedMasterData As Variant

Private Type UnitPriceRequest
    Nendo As String
    BranchName As String
    OfficeName As String
    LineType As String
    ProjectName As String
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
Private Const BASIC_INFO_LINE_TYPE_CELL As String = "C25"
Private Const BASIC_INFO_PROJECT_NAME_CELL As String = "C26"
Private Const BASIC_INFO_PRICE_KIND_CELL As String = "C27"
Private Const BASIC_INFO_PRICE_KIND_FALLBACK_CELL As String = "B27"
Private Const BASIC_INFO_IMPORTED_LINE_NAMES_CELL As String = "C28"
Private Const PROJECT_NAME_LIST_COL As String = "AE"
Private Const LINE_TYPE_LIST_COL As String = "AF"
Private Const PRICE_KIND_LIST_COL As String = "AG"
Private Const LIST_START_ROW As Long = 2
Private Const PROJECT_NAME_MASTER_START_ROW As Long = 2
Private Const PROJECT_NAME_MASTER_LAST_ROW As Long = 1048576
Private Const IMPORTED_SHEET_PROPERTY As String = "UnitPriceImported"

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

Public Sub RefreshUnitPriceProjectNameValidation(Optional ByVal wsInfo As Worksheet, _
                                                  Optional ByVal keepProjectName As Boolean = True)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ErrorHandler

    WriteUnitPriceLineTypeValidation wsInfo
    WriteUnitPriceKindValidation wsInfo

    Dim lineType As String
    lineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).Value))
    If lineType = "" Then
        ClearUnitPriceProjectNameValidation wsInfo, Not keepProjectName
        Exit Sub
    End If

    Dim projectNames As Collection
    Set projectNames = LoadUnitPriceProjectNames(GetMasterFilePath(), lineType)
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

    Dim sourceFilePath As String
    sourceFilePath = ResolveUnitPriceSourceFilePath(request, masterRow)
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
    MsgBox BuildImportCompleteMessage(selectedSheetNames, sourceFilePath), vbInformation, "完了"
End Sub

Private Function TryReadUnitPriceRequest(ByVal wsInfo As Worksheet, ByRef request As UnitPriceRequest) As Boolean
    request.Nendo = CommonExtractYear4Digits(CStr(wsInfo.Range(BASIC_INFO_YEAR_CELL).Value))
    request.BranchName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_BRANCH_CELL).Value))
    request.OfficeName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_OFFICE_CELL).Value))
    request.LineType = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL).Value))
    request.ProjectName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).Value))
    request.UnitPriceKind = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PRICE_KIND_CELL).Value))
    If request.UnitPriceKind = "" Then
        request.UnitPriceKind = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PRICE_KIND_FALLBACK_CELL).Value))
    End If

    If request.Nendo = "" Then
        MsgBox "基本情報シート B4 に4桁の年度が見つかりません。", vbExclamation
        Exit Function
    End If
    If request.BranchName = "" Or request.OfficeName = "" Then
        MsgBox "基本情報シート B6 または C6 が空です。支店名・出張所名を確認してください。", vbExclamation
        Exit Function
    End If
    If request.LineType = "" Then
        MsgBox "基本情報シート C25 の線区区分を選択してください。", vbExclamation
        Exit Function
    End If
    If request.ProjectName = "" Then
        MsgBox "基本情報シート C26 の工事件名を選択してください。", vbExclamation
        Exit Function
    End If
    If request.UnitPriceKind = "" Then
        MsgBox "基本情報シート C27 の単価区分を選択してください。", vbExclamation
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
                                                ByRef masterRow As UnitPriceMasterRow) As String
    Dim dataRoot As String
    dataRoot = GetUnitPriceDataRootPath()
    If dataRoot = "" Then
        MsgBox "単価データフォルダが見つかりません。" & vbCrLf & dataRoot, vbExclamation
        Exit Function
    End If
    If Dir(dataRoot, vbDirectory) = "" Then
        MsgBox "単価データフォルダが見つかりません。" & vbCrLf & dataRoot, vbExclamation
        Exit Function
    End If

    Dim lineFolder As String
    lineFolder = FindChildFolderByKey(dataRoot, request.LineType, True)
    If lineFolder = "" Then
        MsgBox "線区区分フォルダが見つかりません。" & vbCrLf & _
               "線区区分：" & request.LineType & vbCrLf & _
               dataRoot, vbExclamation
        Exit Function
    End If

    Dim branchGroupFolder As String
    branchGroupFolder = FindChildFolderByKey(lineFolder, masterRow.BranchGroupName, False)
    If branchGroupFolder = "" Then
        MsgBox "支社フォルダが見つかりません。" & vbCrLf & _
               "支社：" & masterRow.BranchGroupName & vbCrLf & _
               lineFolder, vbExclamation
        Exit Function
    End If

    Dim sectionFolder As String
    sectionFolder = FindChildFolderByKey(branchGroupFolder, masterRow.UnitPriceSectionName, True)
    If sectionFolder = "" Then
        MsgBox "単価適用保線区フォルダが見つかりません。" & vbCrLf & _
               "単価適用保線区：" & masterRow.UnitPriceSectionName & vbCrLf & _
               branchGroupFolder, vbExclamation
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

    ResolveUnitPriceSourceFilePath = FindUnitPriceWorkbook(priceFolder, request.ProjectName)
    If ResolveUnitPriceSourceFilePath = "" Then
        MsgBox "工事件名に一致する単価表が見つかりません。" & vbCrLf & _
               "工事件名：" & request.ProjectName & vbCrLf & _
               priceFolder, vbExclamation
    End If
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

    Set sourceBook = Workbooks.Open(FileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
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

    Set sourceBook = Workbooks.Open(FileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
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

Private Sub WriteSelectedLineNames(ByVal wsInfo As Worksheet, ByVal selectedSheetNames As Collection)
    If wsInfo Is Nothing Then Exit Sub
    wsInfo.Range(BASIC_INFO_IMPORTED_LINE_NAMES_CELL).Value = JoinCollectionText(selectedSheetNames, JapaneseCommaText())
End Sub

Private Function JoinCollectionText(ByVal values As Collection, ByVal delimiter As String) As String
    If values Is Nothing Then Exit Function

    Dim result As String
    Dim item As Variant
    For Each item In values
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(item)
    Next item

    JoinCollectionText = result
End Function

Private Function JapaneseCommaText() As String
    JapaneseCommaText = ChrW$(&H3001)
End Function

Public Sub ClearUnitPriceSheets(Optional ByVal targetBook As Workbook)
    If targetBook Is Nothing Then Set targetBook = ThisWorkbook
    DeleteImportedUnitPriceSheets targetBook
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

    Dim i As Long
    Dim candidate As String
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
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", "[", "]")

    Dim result As String
    result = Trim$(sourceName)

    Dim i As Long
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

Private Function LoadUnitPriceProjectNames(ByVal sourceFilePath As String, _
                                           ByVal lineType As String) As Collection
    Dim sheetName As String

    On Error GoTo ErrorHandler
    sheetName = ResolveUnitPriceProjectNameMasterSheetName(lineType)
    If sheetName = "" Then Exit Function

    Set LoadUnitPriceProjectNames = LoadUnitPriceProjectNamesByAdo(sourceFilePath, sheetName)
    Exit Function

ErrorHandler:
    Set LoadUnitPriceProjectNames = Nothing
End Function

Private Function ResolveUnitPriceProjectNameMasterSheetName(ByVal lineType As String) As String
    Select Case CommonNormalizeText(lineType)
        Case ZAIRAISEN_NAME
            ResolveUnitPriceProjectNameMasterSheetName = PROJECT_MASTER_SHEET_NAME
        Case SHINKANSEN_NAME
            ResolveUnitPriceProjectNameMasterSheetName = SHINKANSEN_NAME & "_" & PROJECT_MASTER_SHEET_NAME
    End Select
End Function

Private Function LoadUnitPriceProjectNamesByAdo(ByVal sourceFilePath As String, _
                                                ByVal sheetName As String) As Collection
    Dim cn As Object
    Dim rs As Object
    Dim result As Collection

    On Error GoTo ErrorHandler
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then Exit Function

    Dim sql As String
    sql = "SELECT F1 FROM " & BuildAdoSheetRangeName(sheetName, "A", PROJECT_NAME_MASTER_START_ROW, "A", PROJECT_NAME_MASTER_LAST_ROW) & _
          " WHERE F1 IS NOT NULL"

    Set rs = cn.Execute(sql)
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

Private Function BuildAdoSheetRangeName(ByVal sheetName As String, _
                                        ByVal startCol As String, _
                                        ByVal startRow As Long, _
                                        ByVal endCol As String, _
                                        ByVal endRow As Long) As String
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

    ResetUnitPriceProjectNameValidation wsInfo.Range(BASIC_INFO_LINE_TYPE_CELL), listRange
    wsInfo.Columns(LINE_TYPE_LIST_COL & ":" & LINE_TYPE_LIST_COL).Hidden = True
End Sub

Private Sub WriteUnitPriceKindValidation(ByVal wsInfo As Worksheet)
    Dim listRange As Range
    Set listRange = wsInfo.Range(PRICE_KIND_LIST_COL & LIST_START_ROW).Resize(2, 1)

    wsInfo.Columns(PRICE_KIND_LIST_COL & ":" & PRICE_KIND_LIST_COL).Hidden = False
    listRange.ClearContents
    listRange.Cells(1, 1).Value = INITIAL_PRICE_NAME
    listRange.Cells(2, 1).Value = DESIGN_CHANGE_PRICE_NAME

    ResetUnitPriceProjectNameValidation wsInfo.Range(BASIC_INFO_PRICE_KIND_CELL), listRange
    wsInfo.Columns(PRICE_KIND_LIST_COL & ":" & PRICE_KIND_LIST_COL).Hidden = True
End Sub

Private Sub WriteUnitPriceProjectNameValidation(ByVal wsInfo As Worksheet, _
                                                ByVal projectNames As Collection, _
                                                ByVal keepProjectName As Boolean)
    Dim maxProjectRows As Long
    maxProjectRows = Application.Max(1, projectNames.Count)

    Dim listRange As Range
    Set listRange = wsInfo.Range(PROJECT_NAME_LIST_COL & LIST_START_ROW).Resize(maxProjectRows, 1)

    Dim currentProjectName As String
    currentProjectName = CommonNormalizeText(CStr(wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).Value))

    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = False
    wsInfo.Range(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).ClearContents

    Dim i As Long
    For i = 1 To projectNames.Count
        listRange.Cells(i, 1).Value = projectNames(i)
    Next i

    ResetUnitPriceProjectNameValidation wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL), _
                                        listRange.Resize(Application.Max(1, projectNames.Count))

    If Not keepProjectName Or currentProjectName = "" Or Not CollectionContainsText(projectNames, currentProjectName) Then
        wsInfo.Range(BASIC_INFO_PROJECT_NAME_CELL).ClearContents
    End If

    wsInfo.Columns(PROJECT_NAME_LIST_COL & ":" & PROJECT_NAME_LIST_COL).Hidden = True
End Sub

Private Sub ResetUnitPriceProjectNameValidation(ByVal targetCell As Range, ByVal listRange As Range)
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & listRange.Address(True, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
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
    Dim extensions As Variant
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")

    Dim normalizedProjectName As String
    normalizedProjectName = NormalizeMatchText(projectName)

    Dim ext As Variant
    For Each ext In extensions
        Dim fileName As String
        fileName = Dir(folderPath & "\" & CStr(ext), vbNormal)
        Do While fileName <> ""
            If Left$(fileName, 2) <> "~$" Then
                If InStr(1, NormalizeMatchText(RemoveFileExtension(fileName)), normalizedProjectName, vbTextCompare) > 0 Then
                    FindUnitPriceWorkbook = folderPath & "\" & fileName
                    Exit Function
                End If
            End If
            fileName = Dir()
        Loop
    Next ext
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
    Dim a As String
    Dim b As String
    a = NormalizeMatchText(keyText)
    b = NormalizeMatchText(folderText)

    If a = "" Or b = "" Then Exit Function
    FolderTextMatches = (StrComp(a, b, vbTextCompare) = 0 Or _
                         Left$(a, Len(b)) = b Or _
                         Left$(b, Len(a)) = a Or _
                         InStr(1, a, b, vbTextCompare) > 0 Or _
                         InStr(1, b, a, vbTextCompare) > 0)
End Function

Private Function MasterTextMatches(ByVal masterText As String, ByVal searchText As String) As Boolean
    Dim masterValue As String
    Dim searchValue As String
    masterValue = NormalizeMatchText(masterText)
    searchValue = NormalizeMatchText(searchText)

    If masterValue = "" Or searchValue = "" Then Exit Function
    MasterTextMatches = (StrComp(masterValue, searchValue, vbTextCompare) = 0 Or _
                         Left$(masterValue, Len(searchValue)) = searchValue)
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

Private Function RemoveFileExtension(ByVal fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        RemoveFileExtension = Left$(fileName, pos - 1)
    Else
        RemoveFileExtension = fileName
    End If
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

'--------------------------------------------------------------------------
'  FirstExistingFilePath
'    候補パスを順に検索し、最初に存在するファイルパスを返す。
'    SharePoint同期フォルダ未同期時など、ThisWorkbook.Path が
'    https:// 形式になる環境では Dir() がエラー52を返すため、
'    On Error Resume Next でスキップして次の候補へ進む。
'--------------------------------------------------------------------------
Private Function FirstExistingFilePath(ByVal candidates As Collection) As String
    Dim candidate As Variant
    Dim found As Boolean

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
                                            ByVal sourceFilePath As String) As String
    BuildImportCompleteMessage = CStr(selectedSheetNames.Count) & "件の積算線区単価表を取り込みました。" & vbCrLf & _
                                 sourceFilePath
End Function

Private Function OrderInvoiceDocumentFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = "線路出張所用_注文書_請求書アクセスサイト - ドキュメント"
    End If
    OrderInvoiceDocumentFolderText = cached
End Function

Private Function UnitPriceMasterFolderText() As String
    Static cached As String
    If cached = "" Then cached = "単価マスタ"
    UnitPriceMasterFolderText = cached
End Function
