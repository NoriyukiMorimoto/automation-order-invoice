Option Explicit

' ??????????10?(AH???)?????????AJ(????)???????
Private Const LIST_BRANCH_COL As String = "AK"
Private Const LIST_OFFICE_COL As String = "AL"
Private Const LIST_START_ROW As Long = 2
Private Const OFFICE_COMBO_NAME As String = "ComboBox1"
Private Const OFFICE_COMBO_WIDTH_POINTS As Double = 310.5

Private mSuppressC6Change As Boolean
Private mInPromptOffice As Boolean
Private mOfficePromptTime As Date

Public Function IsSuppressingC6Change() As Boolean
    IsSuppressingC6Change = mSuppressC6Change
End Function

Public Function IsPromptingOfficeComboBox() As Boolean
    IsPromptingOfficeComboBox = mInPromptOffice
End Function

Public Sub FillManagerNameToBasicInfo()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox UiMsgBasicInfoSheetNotFoundCheckNameText(), vbExclamation
        Exit Sub
    End If

    Dim yearText As String
    yearText = CommonExtractYear4Digits(Trim$(CStr(wsInfo.Range("B4").value)))
    If yearText = "" Then
        MsgBox UiMsgBasicInfoYearNotFoundB4ExampleText(), vbExclamation
        Exit Sub
    End If

    Dim BranchName As String, OfficeName As String
    BranchName = CommonNormalizeText(CStr(wsInfo.Range("B6").value))
    OfficeName = CommonNormalizeText(CStr(wsInfo.Range("C6").value))
    If BranchName = "" Or OfficeName = "" Then
        MsgBox UiMsgBasicInfoBranchOfficeEmptyText(), vbExclamation
        Exit Sub
    End If

    Dim sourceFilePath As String
    sourceFilePath = GetManagerListFilePath(yearText)
    If sourceFilePath = "" Then Exit Sub

    Dim rows As Collection
    Set rows = LoadManagerListRows(sourceFilePath)
    If rows Is Nothing Then
        MsgBox UiMsgManagerListFileUnreadableText() & vbCrLf & sourceFilePath, vbExclamation
        Exit Sub
    End If

    Dim foundName As String
    Dim rowData As Variant
    For Each rowData In rows
        If StrComp(CommonNormalizeText(CStr(rowData(0))), BranchName, vbTextCompare) = 0 And _
           StrComp(CommonNormalizeText(CStr(rowData(1))), OfficeName, vbTextCompare) = 0 Then
            foundName = Trim$(CStr(rowData(2)))
            Exit For
        End If
    Next rowData

    If foundName = "" Then
        MsgBox UiMsgManagerBranchOfficeNotFoundText() & vbCrLf & _
               UiMsgBranchNameLabelText() & BranchName & vbCrLf & _
               UiMsgOfficeNameLabelText() & OfficeName, vbExclamation
    Else
        wsInfo.Range("F6").value = foundName
    End If
End Sub

Public Function RefreshBranchOfficeValidation(Optional ByVal keepOffice As Boolean = True) As Boolean
    RefreshBranchOfficeValidation = False
    mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation 開始 keepOffice=" & keepOffice

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox UiMsgBasicInfoSheetNotFoundCheckNameText(), vbExclamation
        Exit Function
    End If

    Dim yearText As String
    yearText = CommonExtractYear4Digits(Trim$(CStr(wsInfo.Range("B4").value)))
    If yearText = "" Then
        mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: yearText 空 -> Exit"
        MsgBox UiMsgBasicInfoYearNotFoundB4ExampleText(), vbExclamation
        Exit Function
    End If

    Dim sourceFilePath As String
    sourceFilePath = GetManagerListFilePath(yearText)
    mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: sourceFilePath=[" & sourceFilePath & "]"
    If sourceFilePath = "" Then
        mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: sourceFilePath 空 -> Exit"
        Exit Function
    End If

    Dim rows As Collection
    Set rows = LoadManagerListRows(sourceFilePath)
    If rows Is Nothing Then
        mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: rows Is Nothing -> Exit"
        MsgBox UiMsgManagerListFileUnreadableText() & vbCrLf & sourceFilePath, vbExclamation
        Exit Function
    End If
    mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: rows.Count=" & rows.Count

    Dim branchList As Object, officeList As Object
    Set branchList = CreateObject("Scripting.Dictionary")
    Set officeList = CreateObject("Scripting.Dictionary")
    branchList.CompareMode = vbTextCompare
    officeList.CompareMode = vbTextCompare

    Dim selectedBranch As String
    selectedBranch = CommonNormalizeText(CStr(wsInfo.Range("B6").value))
    mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: selectedBranch=[" & selectedBranch & "]"

    Dim rowData As Variant, BranchName As String, OfficeName As String
    For Each rowData In rows
        BranchName = CommonNormalizeText(CStr(rowData(0)))
        OfficeName = CommonNormalizeText(CStr(rowData(1)))
        If IsManagerListDataRow(BranchName, OfficeName) Then
            If StrComp(BranchName, HeadOfficeText(), vbTextCompare) <> 0 Then
                If Not branchList.Exists(BranchName) Then branchList.Add BranchName, BranchName
                If selectedBranch <> "" Then
                    If StrComp(BranchName, selectedBranch, vbTextCompare) = 0 Then
                        If Not officeList.Exists(OfficeName) Then officeList.Add OfficeName, OfficeName
                    End If
                End If
            End If
        End If
    Next rowData

    mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: 集計完了 branchList=" & branchList.Count & " officeList=" & officeList.Count & " -> WriteValidationLists 呼び出し"
    RefreshBranchOfficeValidation = WriteValidationLists(wsInfo, branchList, officeList, keepOffice)
End Function

Private Function WriteValidationLists(ByVal wsInfo As Worksheet, _
                                      ByVal branchList As Object, _
                                      ByVal officeList As Object, _
                                      ByVal keepOffice As Boolean) As Boolean
    WriteValidationLists = False
    mod_DebugLog.Log "[FillMgr] WriteValidationLists: keepOffice=" & keepOffice & " branchCount=" & branchList.Count & " officeCount=" & officeList.Count

    Dim branchCol As String, officeCol As String
    branchCol = LIST_BRANCH_COL
    officeCol = LIST_OFFICE_COL

    wsInfo.Columns(branchCol & ":" & officeCol).Hidden = False
    wsInfo.Range(branchCol & ":" & officeCol).ClearContents

    WriteDictionaryKeysToColumn wsInfo, branchList, branchCol
    WriteDictionaryKeysToColumn wsInfo, officeList, officeCol

    ResetListValidation wsInfo.Range("B6"), _
                        wsInfo.Range(branchCol & LIST_START_ROW).Resize(Application.Max(1, branchList.Count, 1))
    ResetListValidation wsInfo.Range("C6"), _
                        wsInfo.Range(officeCol & LIST_START_ROW).Resize(Application.Max(1, officeList.Count, 1))

    If branchList.Count = 0 Then wsInfo.Range("B6").ClearContents

    Dim currentOffice As String
    currentOffice = CommonNormalizeText(CStr(wsInfo.Range("C6").value))
    If Not keepOffice Or currentOffice = "" Or Not officeList.Exists(currentOffice) Then
        wsInfo.Range("C6").ClearContents
    End If

    UpdateOfficeComboBox wsInfo, officeList
    wsInfo.Columns(branchCol & ":" & officeCol).Hidden = True

    If Not keepOffice And officeList.Count > 0 Then
        mod_DebugLog.Log "[FillMgr] WriteValidationLists -> True（コンボ表示要）"
        WriteValidationLists = True
    Else
        mod_DebugLog.Log "[FillMgr] WriteValidationLists -> False (keepOffice=" & keepOffice & " officeCount=" & officeList.Count & ")"
    End If
End Function

Private Sub WriteDictionaryKeysToColumn(ByVal wsInfo As Worksheet, _
                                         ByVal dict As Object, _
                                         ByVal colLetter As String)
    If dict Is Nothing Then Exit Sub
    Dim total As Long
    total = dict.Count
    If total = 0 Then Exit Sub

    Dim keysArr As Variant
    keysArr = dict.Keys

    Dim outArr() As Variant
    ReDim outArr(1 To total, 1 To 1)
    Dim i As Long
    For i = 0 To total - 1
        outArr(i + 1, 1) = keysArr(i)
    Next i

    wsInfo.Range(colLetter & LIST_START_ROW).Resize(total, 1).value = outArr
End Sub

' ?????????????????(T:AG??2?9)??????
' ????????(10????)??????????????? AK/AL(??????)?
Public Sub CleanupLegacyBasicInfoValidationListDebris(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    Const LEGACY_LIST_RANGE As String = "T2:AG9"

    On Error Resume Next
    wsInfo.Range(LEGACY_LIST_RANGE).ClearContents
    On Error GoTo 0
End Sub

Private Sub ResetListValidation(ByVal targetCell As Range, ByVal listRange As Range)
    With targetCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & listRange.Address(True, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With
End Sub

Public Sub ScheduleOfficeComboBoxPrompt()
    CancelScheduledOfficeComboBoxPrompt
    mOfficePromptTime = Now + TimeSerial(0, 0, 1)

    On Error Resume Next
    Application.OnTime mOfficePromptTime, "'" & ThisWorkbook.Name & "'!PromptOfficeComboBox"
    If Err.Number <> 0 Then
        Err.Clear
        mOfficePromptTime = 0
        PromptOfficeComboBox
    End If
    On Error GoTo 0
End Sub

Public Sub CancelScheduledOfficeComboBoxPrompt()
    If mOfficePromptTime = 0 Then Exit Sub

    On Error Resume Next
    Application.OnTime mOfficePromptTime, _
                       "'" & ThisWorkbook.Name & "'!PromptOfficeComboBox", _
                       Schedule:=False
    On Error GoTo 0
    mOfficePromptTime = 0
End Sub

Public Sub PromptOfficeComboBox()
    mOfficePromptTime = 0
    mod_DebugLog.Log "[FillMgr] PromptOfficeComboBox 開始 EnableEvents=" & Application.EnableEvents
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        mod_DebugLog.Log "[FillMgr] PromptOfficeComboBox: wsInfo Is Nothing -> Exit"
        Exit Sub
    End If

    If ActiveSheet Is Nothing Then
        mod_DebugLog.Log "[FillMgr] PromptOfficeComboBox: ActiveSheet Is Nothing -> Exit"
        Exit Sub
    End If
    If Not ActiveSheet Is wsInfo Then
        mod_DebugLog.Log "[FillMgr] PromptOfficeComboBox: ActiveSheet=[" & ActiveSheet.Name & "] <> wsInfo=[" & wsInfo.Name & "] -> Exit"
        Exit Sub
    End If

    Dim prevEnableEvents As Boolean
    prevEnableEvents = Application.EnableEvents
    mod_DebugLog.Log "[FillMgr] PromptOfficeComboBox: prevEnableEvents=" & prevEnableEvents & " -> EnableEvents=True にして ShowOfficeComboBox 呼び出し"

    On Error GoTo ExitHandler
    mInPromptOffice = True
    Application.EnableEvents = True
    ShowOfficeComboBox wsInfo

ExitHandler:
    mod_DebugLog.Log "[FillMgr] PromptOfficeComboBox 終了 Err=" & Err.Number & " mInPromptOffice=" & mInPromptOffice
    mInPromptOffice = False
    Application.EnableEvents = prevEnableEvents
End Sub

Public Sub HideOfficeComboBox(Optional ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error Resume Next
    Dim ole As OLEObject
    Set ole = wsInfo.OLEObjects(OFFICE_COMBO_NAME)
    If Not ole Is Nothing Then
        ClearOfficeComboBoxLinkedCell ole
        ole.Object.value = CStr(wsInfo.Range("C6").value)
        ole.Visible = False
    End If
    On Error GoTo 0
End Sub

Public Sub CommitOfficeComboBoxSelection(Optional ByVal selectC6 As Boolean = True)
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ExitHandler

    Dim ole As OLEObject
    Set ole = wsInfo.OLEObjects(OFFICE_COMBO_NAME)
    If ole Is Nothing Then GoTo ExitHandler

    Dim selectedOffice As String
    Dim previousOffice As String
    selectedOffice = CommonNormalizeText(CStr(ole.Object.value))
    If selectedOffice = "" Then GoTo ExitHandler
    previousOffice = CommonNormalizeText(CStr(wsInfo.Range("C6").value))

    mSuppressC6Change = True
    ClearOfficeComboBoxLinkedCell ole
    wsInfo.Range("C6").value = selectedOffice
    mSuppressC6Change = False

    If StrComp(previousOffice, selectedOffice, vbTextCompare) <> 0 Then
        If Not mod_MaterialPriceImport.IsImportingUnitPriceData() Then
            mod_BasicInfoUpdate.SilentClearBasicInfo wsInfo
        End If
    End If
    FillManagerNameToBasicInfo

ExitHandler:
    mSuppressC6Change = False
    HideOfficeComboBox wsInfo
    If selectC6 Then
        On Error Resume Next
        wsInfo.Range("C6").Select
        On Error GoTo 0
    End If
End Sub

Private Sub ShowOfficeComboBox(ByVal wsInfo As Worksheet)
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox 開始 EnableEvents=" & Application.EnableEvents
    On Error GoTo ErrorHandler

    Dim ole As OLEObject
    Set ole = GetOfficeComboBox(wsInfo)
    If ole Is Nothing Then
        mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: ole Is Nothing -> ValidationDropdown へ"
        ShowC6ValidationDropdown wsInfo
        Exit Sub
    End If

    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: FitOfficeComboBoxToC6 開始"
    FitOfficeComboBoxToC6 wsInfo, ole
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: wsInfo.Activate 開始"
    wsInfo.Activate
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: C6.Select 開始"
    wsInfo.Range("C6").Select
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: OLEObject.LinkedCell="""" 開始"
    ClearOfficeComboBoxLinkedCell ole
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: ListCount チェック"
    If ole.Object.ListCount = 0 Then
        mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: ListCount=0 -> ValidationDropdown へ"
        HideOfficeComboBox wsInfo
        ShowC6ValidationDropdown wsInfo
        Exit Sub
    End If
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: ListCount=" & ole.Object.ListCount & " ole.ProgID=[" & ole.progID & "] ole.Name=[" & ole.Name & "]"
    ole.Visible = True
    ole.Activate
    On Error Resume Next
    ole.Object.DropDown
    Dim dropDownErr As Long
    Dim dropDownDesc As String
    dropDownErr = Err.Number
    dropDownDesc = Err.Description
    On Error GoTo ErrorHandler
    If dropDownErr <> 0 Then
        mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: DropDown Err=" & dropDownErr & " " & dropDownDesc & " -> SendKeys未使用で表示維持"
        ole.Visible = True
        ole.Activate
    Else
        mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: DropDown 完了"
    End If
    Exit Sub

ErrorHandler:
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: ErrorHandler Err=" & Err.Number & " " & Err.Description
    HideOfficeComboBox wsInfo
    ShowC6ValidationDropdown wsInfo
End Sub

Private Sub FitOfficeComboBoxToC6(ByVal wsInfo As Worksheet, ByVal ole As OLEObject)
    On Error Resume Next
    ole.Left = wsInfo.Range("C6").Left
    mod_DebugLog.Log "[FillMgr] Fit: Left Err=" & Err.Number: Err.Clear
    ole.Top = wsInfo.Range("C6").Top
    mod_DebugLog.Log "[FillMgr] Fit: Top Err=" & Err.Number: Err.Clear
    ole.Width = OFFICE_COMBO_WIDTH_POINTS
    mod_DebugLog.Log "[FillMgr] Fit: Width Err=" & Err.Number: Err.Clear
    ole.Height = wsInfo.Range("C6").Height
    mod_DebugLog.Log "[FillMgr] Fit: Height Err=" & Err.Number: Err.Clear
    ole.Placement = xlMoveAndSize
    mod_DebugLog.Log "[FillMgr] Fit: Placement Err=" & Err.Number: Err.Clear
    On Error GoTo 0
End Sub

Private Function GetOfficeComboBox(ByVal wsInfo As Worksheet) As OLEObject
    On Error Resume Next
    Set GetOfficeComboBox = wsInfo.OLEObjects(OFFICE_COMBO_NAME)
    On Error GoTo 0
    If Not GetOfficeComboBox Is Nothing Then Exit Function

    On Error Resume Next
    Set GetOfficeComboBox = wsInfo.OLEObjects.Add(ClassType:="Forms.ComboBox.1", _
                                                  Link:=False, _
                                                  DisplayAsIcon:=False, _
                                                  Left:=wsInfo.Range("C6").Left, _
                                                  Top:=wsInfo.Range("C6").Top, _
                                                  Width:=OFFICE_COMBO_WIDTH_POINTS, _
                                                  Height:=wsInfo.Range("C6").Height)
    If Not GetOfficeComboBox Is Nothing Then
        GetOfficeComboBox.Name = OFFICE_COMBO_NAME
        GetOfficeComboBox.Visible = False
        FitOfficeComboBoxToC6 wsInfo, GetOfficeComboBox
    End If
    On Error GoTo 0
End Function

Private Sub ShowC6ValidationDropdown(ByVal wsInfo As Worksheet)
    On Error Resume Next
    wsInfo.Activate
    wsInfo.Range("C6").Select
    mod_DebugLog.Log "[FillMgr] ShowC6ValidationDropdown: SendKeys未使用（NumLock保護）"
    On Error GoTo 0
End Sub

Private Sub ClearOfficeComboBoxLinkedCell(ByVal ole As OLEObject)
    On Error Resume Next
    ole.LinkedCell = ""
    mod_DebugLog.Log "[FillMgr] ClearLinkedCell Err=" & Err.Number
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub UpdateOfficeComboBox(ByVal wsInfo As Worksheet, ByVal officeList As Object)
    On Error Resume Next
    Dim ole As OLEObject
    Set ole = GetOfficeComboBox(wsInfo)
    If ole Is Nothing Then Exit Sub

    FitOfficeComboBoxToC6 wsInfo, ole

    With ole.Object
        .Clear
        Dim i As Long
        Dim keysArr As Variant
        If officeList.Count > 0 Then
            keysArr = officeList.Keys
            For i = 0 To officeList.Count - 1
                .AddItem keysArr(i)
            Next i
        End If
        .ListRows = Application.Max(1, Application.Min(12, officeList.Count))
        .MatchRequired = False
        .value = CStr(wsInfo.Range("C6").value)
    End With
    ClearOfficeComboBoxLinkedCell ole
    ole.Visible = False
    On Error GoTo 0
End Sub

Private Function LoadManagerListRows(ByVal sourceFilePath As String) As Collection
    mod_DebugLog.Log "[FillMgr] LoadManagerListRows: ADO試行 path=[" & sourceFilePath & "]"
    Set LoadManagerListRows = LoadManagerListRowsFromAdo(sourceFilePath)
    If LoadManagerListRows Is Nothing Then
        mod_DebugLog.Log "[FillMgr] LoadManagerListRows: ADO失敗 -> Workbook.Open試行"
        Set LoadManagerListRows = LoadManagerListRowsFromWorkbook(sourceFilePath)
        If LoadManagerListRows Is Nothing Then
            mod_DebugLog.Log "[FillMgr] LoadManagerListRows: Workbook.Open も失敗"
        Else
            mod_DebugLog.Log "[FillMgr] LoadManagerListRows: Workbook.Open 成功 Count=" & LoadManagerListRows.Count
        End If
    Else
        mod_DebugLog.Log "[FillMgr] LoadManagerListRows: ADO成功 Count=" & LoadManagerListRows.Count
    End If
End Function

Private Function LoadManagerListRowsFromAdo(ByVal sourceFilePath As String) As Collection
    Dim cn As Object
    Set cn = CommonOpenExcelAdoConnection(sourceFilePath)
    If cn Is Nothing Then Exit Function

    Dim rs As Object

    On Error GoTo ErrorHandler

    Dim sheetName As String
    sheetName = GetFirstWorksheetTableName(cn)
    If sheetName = "" Then GoTo Cleanup

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT [F2], [F3], [F6] FROM [" & sheetName & "]", cn, 0, 1, 1

    Dim rows As Collection
    Set rows = New Collection

    If Not rs.EOF Then rs.MoveNext
    Do Until rs.EOF
        rows.Add Array(CommonNzText(rs.Fields(0).value), _
                       CommonNzText(rs.Fields(1).value), _
                       CommonNzText(rs.Fields(2).value))
        rs.MoveNext
    Loop

    Set LoadManagerListRowsFromAdo = rows

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    mod_DebugLog.Log "[FillMgr] LoadManagerListRowsFromAdo: Err=" & Err.Number & " " & Err.Description
    Set LoadManagerListRowsFromAdo = Nothing
    Resume Cleanup
End Function

Private Function LoadManagerListRowsFromWorkbook(ByVal sourceFilePath As String) As Collection
    Dim sourceBook As Workbook
    Dim previousDisplayAlerts As Boolean
    previousDisplayAlerts = Application.DisplayAlerts

    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False

    Set sourceBook = Application.Workbooks.Open(fileName:=sourceFilePath, _
                                                UpdateLinks:=False, _
                                                ReadOnly:=True, _
                                                AddToMru:=False)

    Dim sourceSheet As Worksheet
    Set sourceSheet = sourceBook.worksheets(1)

    Dim rows As Collection
    Set rows = New Collection

    Dim lastRow As Long
    lastRow = sourceSheet.Cells(sourceSheet.rows.Count, 2).End(xlUp).Row

    Dim rr As Long
    For rr = 2 To lastRow
        rows.Add Array(CommonNzText(sourceSheet.Cells(rr, 2).value), _
                       CommonNzText(sourceSheet.Cells(rr, 3).value), _
                       CommonNzText(sourceSheet.Cells(rr, 6).value))
    Next rr

    Set LoadManagerListRowsFromWorkbook = rows

Cleanup:
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.DisplayAlerts = previousDisplayAlerts
    Exit Function

ErrorHandler:
    mod_DebugLog.Log "[FillMgr] LoadManagerListRowsFromWorkbook: Err=" & Err.Number & " " & Err.Description
    Set LoadManagerListRowsFromWorkbook = Nothing
    Resume Cleanup
End Function

Private Function GetFirstWorksheetTableName(ByVal cn As Object) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(cn)
    If sheetNames Is Nothing Then Exit Function
    If sheetNames.Count = 0 Then Exit Function
    GetFirstWorksheetTableName = CStr(sheetNames(1)) & "$"
End Function

Private Function IsManagerListDataRow(ByVal BranchName As String, ByVal OfficeName As String) As Boolean
    If BranchName = "" Or OfficeName = "" Then Exit Function
    If StrComp(BranchName, BranchHeaderText(), vbTextCompare) = 0 Then Exit Function
    If StrComp(OfficeName, OfficeHeaderText(), vbTextCompare) = 0 Then Exit Function
    If StrComp(OfficeName, OfficeBranchHeaderText(), vbTextCompare) = 0 Then Exit Function
    IsManagerListDataRow = True
End Function

Private Function GetManagerListFilePath(ByVal yearText As String) As String
    Dim folderPath As String
    folderPath = GetManagerListFolderPath()
    If Right$(folderPath, 1) <> Chr$(92) Then folderPath = folderPath & Chr$(92)
    mod_DebugLog.Log "[FillMgr] GetManagerListFilePath: folderPath=[" & folderPath & "]"
    If Dir(folderPath, vbDirectory) = "" Then
        mod_DebugLog.Log "[FillMgr] GetManagerListFilePath: フォルダ不存在 -> Exit"
        MsgBox UiMsgManagerListFolderNotFoundText() & vbCrLf & folderPath, vbExclamation
        Exit Function
    End If
    GetManagerListFilePath = FindManagerListFile(folderPath, yearText)
    mod_DebugLog.Log "[FillMgr] GetManagerListFilePath: 結果=[" & GetManagerListFilePath & "]"
    If GetManagerListFilePath = "" Then
        mod_DebugLog.Log "[FillMgr] GetManagerListFilePath: ファイル不存在 -> Exit"
        MsgBox yearText & UiMsgManagerListFileNotFoundSuffixText(), vbExclamation
    End If
End Function

Private Function GetManagerListFolderPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim candidates As Collection
    Set candidates = New Collection

    If Len(ThisWorkbook.Path) > 0 Then
        candidates.Add BuildManagerListFolderPath(fso.GetParentFolderName(ThisWorkbook.Path), fso)
        candidates.Add BuildManagerListFolderPath(ThisWorkbook.Path, fso)
    End If

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then
        userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    End If

    If Len(Trim$(userProfilePath)) > 0 Then
        candidates.Add BuildManagerListFolderPath(userProfilePath & Chr$(92) & CommonCompanyNameText() & Chr$(92) & _
                                                  OrderInvoiceDocumentFolderText(), fso)
    End If

    Dim resolvedPath As String
    resolvedPath = FirstExistingManagerListFolderPath(candidates)
    mod_DebugLog.Log "[FillMgr] GetManagerListFolderPath: 結果=[" & resolvedPath & "]"
    GetManagerListFolderPath = resolvedPath
End Function

Private Function BuildManagerListFolderPath(ByVal documentRootPath As String, _
                                            ByVal fso As Object) As String
    If Len(Trim$(documentRootPath)) = 0 Then Exit Function
    Dim folderPath As String
    folderPath = fso.BuildPath(documentRootPath, MasterDataFolderText())
    folderPath = fso.BuildPath(folderPath, ManagerNameFolderText())
    BuildManagerListFolderPath = folderPath & Chr$(92)
    mod_DebugLog.Log "[FillMgr] BuildManagerListFolderPath: root=[" & documentRootPath & "] -> [" & BuildManagerListFolderPath & "]"
End Function

Private Function FirstExistingManagerListFolderPath(ByVal candidates As Collection) As String
    Dim candidate As Variant
    For Each candidate In candidates
        If Len(CStr(candidate)) > 0 Then
            If Left$(CStr(candidate), 8) = "https://" Then
                mod_DebugLog.Log "[FillMgr] FirstExistingPath: [" & CStr(candidate) & "] -> https スキップ"
            Else
                Dim existResult As String
                On Error Resume Next
                existResult = Dir(CStr(candidate), vbDirectory)
                On Error GoTo 0
                mod_DebugLog.Log "[FillMgr] FirstExistingPath: [" & CStr(candidate) & "] Dir=[" & existResult & "]"
                If existResult <> "" Then
                    FirstExistingManagerListFolderPath = CStr(candidate)
                    Exit Function
                End If
            End If
        End If
    Next candidate
    mod_DebugLog.Log "[FillMgr] FirstExistingPath: 全候補が存在しない"
End Function

Private Function FindManagerListFile(ByVal folderPath As String, ByVal yearText As String) As String
    Dim listKeyword As String
    listKeyword = ManagerListKeywordText()
    Dim fileName As String
    fileName = ManagerListFileNameText(yearText)
    If Dir(folderPath & fileName, vbNormal) <> "" Then
        FindManagerListFile = folderPath & fileName
        Exit Function
    End If
    fileName = Dir(folderPath & yearText & "_*" & listKeyword & "*.*")
    If fileName <> "" Then
        FindManagerListFile = folderPath & fileName
        Exit Function
    End If
    fileName = Dir(folderPath & "*" & yearText & "*" & listKeyword & "*.*")
    If fileName <> "" Then
        FindManagerListFile = folderPath & fileName
        Exit Function
    End If
    fileName = Dir(folderPath & "*" & listKeyword & "*.*")
    Do While fileName <> ""
        If InStr(fileName, yearText) > 0 Then
            FindManagerListFile = folderPath & fileName
            Exit Function
        End If
        fileName = Dir()
    Loop
End Function

Private Function ManagerListFileNameText(ByVal yearText As String) As String
    ManagerListFileNameText = yearText & "_" & ManagerListKeywordText() & ".xlsx"
End Function

Private Function OrderInvoiceDocumentFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7DDA) & ChrW$(&H8DEF) & ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & _
                 ChrW$(&H7528) & ChrW$(&H5F) & ChrW$(&H6CE8) & ChrW$(&H6587) & ChrW$(&H66F8) & ChrW$(&H5F) & _
                 ChrW$(&H8ACB) & ChrW$(&H6C42) & ChrW$(&H66F8) & ChrW$(&H30A2) & ChrW$(&H30AF) & ChrW$(&H30BB) & _
                 ChrW$(&H30B9) & ChrW$(&H30B5) & ChrW$(&H30A4) & ChrW$(&H30C8) & " - " & _
                 ChrW$(&H30C9) & ChrW$(&H30AD) & ChrW$(&H30E5) & ChrW$(&H30E1) & ChrW$(&H30F3) & ChrW$(&H30C8)
    End If
    OrderInvoiceDocumentFolderText = cached
End Function

Private Function MasterDataFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & ChrW$(&H30C7) & _
                 ChrW$(&H30FC) & ChrW$(&H30BF)
    End If
    MasterDataFolderText = cached
End Function

Private Function ManagerNameFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & ChrW$(&H9577) & ChrW$(&H540D)
    End If
    ManagerNameFolderText = cached
End Function

Private Function ManagerListKeywordText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7DDA) & ChrW$(&H8DEF) & ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & _
                 ChrW$(&H9577) & ChrW$(&H30EA) & ChrW$(&H30B9) & ChrW$(&H30C8)
    End If
    ManagerListKeywordText = cached
End Function

Private Function HeadOfficeText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H672C) & ChrW$(&H793E)
    HeadOfficeText = cached
End Function

Private Function BranchHeaderText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H652F) & ChrW$(&H5E97) & ChrW$(&H540D)
    BranchHeaderText = cached
End Function

Private Function OfficeHeaderText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & ChrW$(&H540D)
    OfficeHeaderText = cached
End Function

Private Function OfficeBranchHeaderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & ChrW$(&H30FB) & _
                 ChrW$(&H652F) & ChrW$(&H6240) & ChrW$(&H540D)
    End If
    OfficeBranchHeaderText = cached
End Function
