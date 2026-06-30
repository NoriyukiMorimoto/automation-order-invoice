Option Explicit

Public SelectionConfirmed As Boolean
Private SharedMasterData As Variant

Public Sub ClearSharedMasterData()
    If IsArray(SharedMasterData) Then
        Erase SharedMasterData
    End If
    SharedMasterData = Empty
End Sub

Private Sub UserForm_Initialize()
    SelectionConfirmed = False
    SetupListView
    If IsEmpty(SharedMasterData) Then
        LoadMasterDataToMemory
    Else
        RefreshList ""
    End If
End Sub

Private Sub UserForm_Activate()
    Me.TextBox1.SetFocus
End Sub

Private Sub SetupListView()
    With Me.ListView1
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideColumnHeaders = False
        .MultiSelect = False
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , ProjectStatusColumnProjectNoText(), 100, lvwColumnCenter
        .ColumnHeaders.Add , , ProjectStatusColumnProjectNameText(), 500
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "", 0
    End With
End Sub

Private Sub LoadMasterDataToMemory()
    Dim previousScreenUpdating As Boolean
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    previousScreenUpdating = Application.screenUpdating
    On Error GoTo ErrorHandler

    Dim folderPath As String: folderPath = GetProjectStatusDataFolderPath()
    Dim targetYear As Long
    Dim targetBranch As String
    Dim targetOffice As String

    targetYear = ProjectSelectionYear
    targetBranch = ProjectSelectionBranchName
    targetOffice = ProjectSelectionOfficeName

    Dim targetBranchOffice As String
    targetBranchOffice = GetProjectSelectionBranchOfficeSearchKey(targetBranch, targetOffice)

    Dim sourceFilePath As String
    sourceFilePath = folderPath & CStr(targetYear) & "_" & _
        GetProjectSelectionBranchNameForFile(targetBranch) & ProjectStatusFileSuffixText()
    LogProjectSelection "folder=[" & folderPath & "] file=[" & sourceFilePath & "] key=[" & targetBranchOffice & "] year=" & CStr(targetYear)

    Me.Caption = UiMsgProjectStatusLoadingCaptionText(): DoEvents
    Application.screenUpdating = False

    Dim sourceArr As Variant
    sourceArr = GetProjectStatusSourceArray(folderPath, CStr(targetYear), targetBranch)

    Dim i As Long
    Dim hitCount As Long
    hitCount = 0

    If Not IsEmpty(sourceArr) Then
        For i = UBound(sourceArr, 1) To 1 Step -1
            If targetBranchOffice = "" Or _
               InStr(1, RemoveProjectSelectionSpaces(GetProjectSourceValue(sourceArr, i, 7)), targetBranchOffice, vbTextCompare) > 0 Then
                hitCount = hitCount + 1
            End If
        Next i
    End If

    ReDim SharedMasterData(1 To Application.Max(1, hitCount), 1 To 6)

    If hitCount = 0 Then
        SharedMasterData(1, 1) = ""
        SharedMasterData(1, 2) = ""
        SharedMasterData(1, 3) = ""
        SharedMasterData(1, 4) = ""
        SharedMasterData(1, 5) = ""
        SharedMasterData(1, 6) = ""
    Else
        Dim writeIndex As Long
        writeIndex = 1
        For i = UBound(sourceArr, 1) To 1 Step -1
            If targetBranchOffice = "" Or _
               InStr(1, RemoveProjectSelectionSpaces(GetProjectSourceValue(sourceArr, i, 7)), targetBranchOffice, vbTextCompare) > 0 Then
                SharedMasterData(writeIndex, 1) = GetProjectSourceValue(sourceArr, i, 10)
                SharedMasterData(writeIndex, 2) = RemoveProjectSelectionSpaces(GetProjectSourceValue(sourceArr, i, 29) & GetProjectSourceValue(sourceArr, i, 30))
                SharedMasterData(writeIndex, 3) = SharedMasterData(writeIndex, 2)
                SharedMasterData(writeIndex, 4) = GetProjectSourceValue(sourceArr, i, 69)
                SharedMasterData(writeIndex, 5) = GetProjectSourceRawValue(sourceArr, i, 34)
                SharedMasterData(writeIndex, 6) = GetProjectSourceRawValue(sourceArr, i, 35)
                writeIndex = writeIndex + 1
            End If
        Next i
    End If

    RefreshList ""
    LogProjectSelection "loaded=" & CStr(Not IsEmpty(sourceArr)) & " hits=" & CStr(hitCount)
    If hitCount = 0 Then
        MsgBox UiMsgProjectDataNotFoundText() & vbCrLf & vbCrLf & _
               BuildProjectStatusNotFoundDetailText(sourceFilePath, targetBranchOffice, targetYear), vbInformation
    End If
    Me.Caption = targetBranch & " " & targetOffice & UiMsgProjectSelectionCaptionSuffixText()
    GoTo FinallyExit

ErrorHandler:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description

FinallyExit:
    Application.screenUpdating = previousScreenUpdating
    If savedErrNumber <> 0 Then
        Me.Caption = UiMsgProjectNumberSelectionCaptionText()
        MsgBox UiMsgProjectStatusLoadFailedText() & vbCrLf & savedErrDescription, vbExclamation
    End If
End Sub
Private Function GetProjectStatusDataFolderPath() As String
    Dim candidates As Collection
    Set candidates = New Collection

    Dim defaultFolder As String
    defaultFolder = BuildProjectStatusDataFolderPath(Environ$("USERPROFILE"))
    AppendProjectStatusFolderCandidate candidates, defaultFolder

    On Error Resume Next
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    If Len(wbPath) > 0 Then
        If Left$(LCase$(wbPath), 8) <> "https://" Then
            AppendProjectStatusFolderCandidate candidates, BuildProjectStatusDataFolderPath(wbPath)
            AppendProjectStatusFolderCandidate candidates, _
                BuildProjectStatusDataFolderPath(Left$(wbPath, InStrRev(wbPath, Chr$(92)) - 1))
        End If
    End If
    On Error GoTo 0

    Dim i As Long
    For i = 1 To candidates.Count
        Dim folderPath As String
        folderPath = CStr(candidates(i))
        If Len(Dir(folderPath, vbDirectory)) > 0 Then
            GetProjectStatusDataFolderPath = folderPath
            Exit Function
        End If
    Next i

    GetProjectStatusDataFolderPath = defaultFolder
End Function

Private Sub AppendProjectStatusFolderCandidate(ByVal candidates As Collection, ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub

    Dim normalizedPath As String
    normalizedPath = folderPath
    If Right$(normalizedPath, 1) <> Chr$(92) Then normalizedPath = normalizedPath & Chr$(92)

    Dim i As Long
    For i = 1 To candidates.Count
        If StrComp(CStr(candidates(i)), normalizedPath, vbTextCompare) = 0 Then Exit Sub
    Next i
    candidates.Add normalizedPath
End Sub

Private Function BuildProjectStatusDataFolderPath(ByVal rootPath As String) As String
    If Len(rootPath) = 0 Then Exit Function
    BuildProjectStatusDataFolderPath = rootPath & Chr$(92) & mod_common.CommonCompanyNameText() & Chr$(92) & _
        ProjectStatusParentFolderText() & Chr$(92) & ProjectStatusDataSubFolderText() & Chr$(92)
End Function

Private Function GetProjectStatusSourceArray(ByVal folderPath As String, ByVal targetYear As String, ByVal targetBranch As String) As Variant
    Dim sourcePath As String
    sourcePath = folderPath & targetYear & "_" & GetProjectSelectionBranchNameForFile(targetBranch) & ProjectStatusFileSuffixText()

    If Len(Dir(sourcePath, vbNormal)) = 0 Then Exit Function
    GetProjectStatusSourceArray = ReadProjectStatusFileToArray(sourcePath, "Sheet1")
End Function
Private Function ReadProjectStatusFileToArray(ByVal sourcePath As String, ByVal sheetName As String) As Variant
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim lastRow As Long

    If Len(Dir(sourcePath, vbNormal)) = 0 Then Exit Function

    On Error GoTo Cleanup
    Set sourceBook = Workbooks.Open(fileName:=sourcePath, ReadOnly:=True, UpdateLinks:=False)

    On Error Resume Next
    Set sourceSheet = sourceBook.worksheets(sheetName)
    On Error GoTo Cleanup
    If sourceSheet Is Nothing Then Set sourceSheet = sourceBook.worksheets(1)

    lastRow = sourceSheet.Cells(sourceSheet.rows.Count, "G").End(xlUp).Row
    If lastRow < sourceSheet.Cells(sourceSheet.rows.Count, "J").End(xlUp).Row Then lastRow = sourceSheet.Cells(sourceSheet.rows.Count, "J").End(xlUp).Row
    If lastRow < 1 Then GoTo Cleanup

    ReadProjectStatusFileToArray = sourceSheet.Range("A1:BQ" & lastRow).value

Cleanup:
    On Error Resume Next
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
End Function
Private Function GetProjectSourceValue(ByVal sourceArr As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    On Error GoTo ErrorHandler
    If colIndex <= UBound(sourceArr, 2) Then GetProjectSourceValue = CStr(sourceArr(rowIndex, colIndex))
    Exit Function

ErrorHandler:
    GetProjectSourceValue = ""
End Function

Private Function GetProjectSourceRawValue(ByVal sourceArr As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    On Error GoTo ErrorHandler
    If colIndex <= UBound(sourceArr, 2) Then GetProjectSourceRawValue = sourceArr(rowIndex, colIndex)
    Exit Function

ErrorHandler:
    GetProjectSourceRawValue = Empty
End Function

Private Function GetProjectSelectionBranchNameForSearch(ByVal BranchName As String) As String
    BranchName = RemoveProjectSelectionSpaces(BranchName)
    If BranchName = "" Then Exit Function

    If Right$(BranchName, 2) = BranchSuffixText() Then
        GetProjectSelectionBranchNameForSearch = BranchName
    Else
        GetProjectSelectionBranchNameForSearch = BranchName & BranchSuffixText()
    End If
End Function

Private Function GetProjectSelectionBranchNameForFile(ByVal BranchName As String) As String
    BranchName = RemoveProjectSelectionSpaces(BranchName)
    If BranchName = "" Then Exit Function

    If Right$(BranchName, 2) = BranchSuffixText() Then
        GetProjectSelectionBranchNameForFile = BranchName
    Else
        GetProjectSelectionBranchNameForFile = BranchName & BranchSuffixText()
    End If
End Function
Private Function GetProjectSelectionBranchOfficeSearchKey(ByVal BranchName As String, ByVal OfficeName As String) As String
    Dim normalizedBranch As String
    Dim normalizedOffice As String

    normalizedBranch = GetProjectSelectionBranchNameForSearch(BranchName)
    normalizedOffice = RemoveProjectSelectionSpaces(OfficeName)

    If StrComp(normalizedBranch, KobeBranchText(), vbTextCompare) = 0 And _
       StrComp(normalizedOffice, SanyoShinkansenTrackMaintenanceOfficeText(), vbTextCompare) = 0 Then
        GetProjectSelectionBranchOfficeSearchKey = KobeSanyoShinkansenTrackSearchText()
    Else
        GetProjectSelectionBranchOfficeSearchKey = RemoveProjectSelectionSpaces(normalizedBranch & normalizedOffice)
    End If
End Function

Private Function KobeBranchText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H795E) & ChrW$(&H6238) & BranchSuffixText()
    KobeBranchText = cached
End Function

Private Function SanyoShinkansenTrackMaintenanceOfficeText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5C71) & ChrW$(&H967D) & ChrW$(&H65B0) & ChrW$(&H5E79) & ChrW$(&H7DDA) & _
                 ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H30E1) & ChrW$(&H30F3) & ChrW$(&H30C6) & _
                 ChrW$(&H30CA) & ChrW$(&H30F3) & ChrW$(&H30B9) & ChrW$(&H51FA) & ChrW$(&H5F35) & _
                 ChrW$(&H6240)
    End If
    SanyoShinkansenTrackMaintenanceOfficeText = cached
End Function

Private Function KobeSanyoShinkansenTrackSearchText() As String
    Static cached As String
    If cached = "" Then
        cached = KobeBranchText() & ChrW$(&H5C71) & ChrW$(&H967D) & ChrW$(&H65B0) & _
                 ChrW$(&H5E79) & ChrW$(&H7DDA) & ChrW$(&H8ECC)
    End If
    KobeSanyoShinkansenTrackSearchText = cached
End Function
Private Function BranchSuffixText() As String
    BranchSuffixText = ChrW$(&H652F) & ChrW$(&H5E97)
End Function
Private Function RemoveProjectSelectionSpaces(ByVal value As String) As String
    value = Replace$(value, " ", "")
    value = Replace$(value, ChrW$(&H3000), "")
    value = Replace$(value, vbTab, "")
    value = Replace$(value, vbCr, "")
    value = Replace$(value, vbLf, "")
    RemoveProjectSelectionSpaces = value
End Function

Private Sub SetBasicInfoProjectSelection(ByVal projectNo As String, ByVal projectName As String, ByVal dataIndex As Long, Optional ByVal selectedProjectDetail As String = "", Optional ByVal selectedContractDate As String = "", Optional ByVal selectedWorkStartDate As Variant, Optional ByVal selectedWorkEndDate As Variant)
    Dim previousEnableEvents As Boolean
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    previousEnableEvents = Application.EnableEvents
    On Error GoTo ErrorHandler

    Dim targetWs As Worksheet
    On Error Resume Next
    Set targetWs = ThisWorkbook.worksheets(ProjectSelectionTargetSheetName)
    On Error GoTo 0
    If targetWs Is Nothing Then Set targetWs = ActiveCell.Worksheet

    Dim targetCell As Range
    On Error Resume Next
    Set targetCell = targetWs.Range(ProjectSelectionTargetAddress)
    On Error GoTo 0
    If targetCell Is Nothing Then Set targetCell = targetWs.Range("C9")

    Dim projectDetail As String
    Dim contractDate As String
    Dim workStartDate As Variant
    Dim workEndDate As Variant
    projectDetail = selectedProjectDetail
    contractDate = selectedContractDate
    workStartDate = selectedWorkStartDate
    workEndDate = selectedWorkEndDate

    Dim lowerRow As Long
    Dim upperRow As Long
    Dim upperCol As Long

    If IsArray(SharedMasterData) Then
        On Error Resume Next
        lowerRow = LBound(SharedMasterData, 1)
        upperRow = UBound(SharedMasterData, 1)
        upperCol = UBound(SharedMasterData, 2)
        If Err.Number = 0 Then
            If dataIndex >= lowerRow And dataIndex <= upperRow Then
                If projectDetail = "" And upperCol >= 3 Then projectDetail = CStr(SharedMasterData(dataIndex, 3))
                If contractDate = "" And upperCol >= 4 Then contractDate = CStr(SharedMasterData(dataIndex, 4))
                If IsProjectDateValueBlank(workStartDate) And upperCol >= 5 Then workStartDate = SharedMasterData(dataIndex, 5)
                If IsProjectDateValueBlank(workEndDate) And upperCol >= 6 Then workEndDate = SharedMasterData(dataIndex, 6)
            End If
        End If
        Err.Clear
        On Error GoTo 0
    End If

    Application.EnableEvents = False
    targetCell.value = projectNo
    targetWs.Range("C10").value = projectDetail
    mod_MaterialPriceImport.AutoFillLineTypeFromWorkName targetWs
    mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation targetWs, False
    mod_MaterialPriceImport.AutoFillProjectNameFromWorkName targetWs
    SetBasicInfoContractDateValue targetWs.Range("C11"), contractDate
    SetBasicInfoDateValueLikeCell targetWs.Range("C15"), workStartDate, targetWs.Range("C11")
    SetBasicInfoDateValueLikeCell targetWs.Range("C16"), workEndDate, targetWs.Range("C11")
    mod_BasicInfoUpdate.ApplyWorkDaysFromWorkDates targetWs
    ' ????(C9)??EnableEvents=False???????Change???????
    ' ???(???????)???????????????????
    mod_BasicInfoGuide.OnCellChanged targetWs, targetWs.Range("C9")
    Application.EnableEvents = previousEnableEvents

    SelectionConfirmed = True
    Unload Me
    Exit Sub

ErrorHandler:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    Application.EnableEvents = previousEnableEvents
    MsgBox UiMsgProjectApplyToBasicInfoFailedText() & vbCrLf & _
           "Err " & CStr(savedErrNumber) & ": " & savedErrDescription, vbExclamation
End Sub

Private Sub SetBasicInfoContractDateValue(ByVal targetCell As Range, ByVal contractDate As String)
    Dim parsedDate As Date
    Dim displayRange As Range
    Set displayRange = targetCell.MergeArea

    If TryParseProjectContractDate(contractDate, parsedDate) Then
        targetCell.value = parsedDate
        displayRange.NumberFormatLocal = ProjectStatusDateNumberFormatText()
    Else
        targetCell.value = RemoveProjectContractWeekday(contractDate)
        displayRange.NumberFormatLocal = ProjectStatusDateNumberFormatText()
    End If
End Sub

Private Function IsProjectDateValueBlank(ByVal sourceValue As Variant) As Boolean
    If IsError(sourceValue) Or IsNull(sourceValue) Or IsEmpty(sourceValue) Then
        IsProjectDateValueBlank = True
    Else
        IsProjectDateValueBlank = (Len(Trim$(CStr(sourceValue))) = 0)
    End If
End Function

Private Sub SetBasicInfoDateValueLikeCell(ByVal targetCell As Range, ByVal sourceValue As Variant, ByVal formatSourceCell As Range)
    Dim parsedDate As Date
    Dim displayRange As Range
    Set displayRange = targetCell.MergeArea

    If IsProjectDateValueBlank(sourceValue) Then
        targetCell.value = ""
    ElseIf TryParseProjectDateValue(sourceValue, parsedDate) Then
        targetCell.value = parsedDate
    Else
        targetCell.value = RemoveProjectContractWeekday(CStr(sourceValue))
    End If
    displayRange.NumberFormatLocal = formatSourceCell.MergeArea.NumberFormatLocal
End Sub

Private Function TryParseProjectDateValue(ByVal sourceValue As Variant, ByRef parsedDate As Date) As Boolean
    If IsProjectDateValueBlank(sourceValue) Then Exit Function

    If IsDate(sourceValue) Then
        parsedDate = CDate(sourceValue)
        TryParseProjectDateValue = True
        Exit Function
    End If

    If IsNumeric(sourceValue) Then
        parsedDate = DateSerial(1899, 12, 30) + CDbl(sourceValue)
        TryParseProjectDateValue = True
        Exit Function
    End If

    TryParseProjectDateValue = TryParseProjectContractDate(CStr(sourceValue), parsedDate)
End Function

Private Function TryParseProjectContractDate(ByVal sourceText As String, ByRef parsedDate As Date) As Boolean
    sourceText = RemoveProjectContractWeekday(sourceText)
    sourceText = Trim$(sourceText)
    If sourceText = "" Then Exit Function

    On Error GoTo ErrorHandler
    If IsDate(sourceText) Then
        parsedDate = CDate(sourceText)
        TryParseProjectContractDate = True
        Exit Function
    End If

    sourceText = Replace$(sourceText, ".", "/")
    sourceText = Replace$(sourceText, "-", "/")
    sourceText = Replace$(sourceText, ProjectStatusYearSuffixText(), "/")
    sourceText = Replace$(sourceText, ProjectStatusMonthSuffixText(), "/")
    sourceText = Replace$(sourceText, ProjectStatusDaySuffixText(), "")
    If IsDate(sourceText) Then
        parsedDate = CDate(sourceText)
        TryParseProjectContractDate = True
    End If
    Exit Function

ErrorHandler:
    TryParseProjectContractDate = False
End Function

Private Function RemoveProjectContractWeekday(ByVal sourceText As String) As String
    Dim p As Long
    sourceText = Trim$(sourceText)

    p = InStr(sourceText, "(")
    If p > 0 Then sourceText = Left$(sourceText, p - 1)

    p = InStr(sourceText, ProjectStatusFullWidthOpenParenText())
    If p > 0 Then sourceText = Left$(sourceText, p - 1)

    Dim weekdayNames As Variant
    weekdayNames = ProjectStatusWeekdayNamesLongArray()
    Dim j As Long
    For j = 0 To UBound(weekdayNames)
        Dim suffixH As String
        suffixH = " " & weekdayNames(j)
        If Right$(sourceText, Len(suffixH)) = suffixH Then
            sourceText = Left$(sourceText, Len(sourceText) - Len(suffixH))
            sourceText = Trim$(sourceText)
            Exit For
        End If
        Dim suffixZ As String
        suffixZ = ProjectStatusFullWidthSpaceText() & weekdayNames(j)
        If Right$(sourceText, Len(suffixZ)) = suffixZ Then
            sourceText = Left$(sourceText, Len(sourceText) - Len(suffixZ))
            sourceText = Trim$(sourceText)
            Exit For
        End If
    Next j

    Dim weekdays As Variant
    weekdays = ProjectStatusWeekdayNamesShortArray()
    Dim i As Long
    For i = 0 To UBound(weekdays)
        Dim sLen As Long
        sLen = Len(sourceText)
        If sLen >= 2 Then
            If Right$(sourceText, 1) = weekdays(i) Then
                sourceText = Left$(sourceText, sLen - 1)
                sourceText = Trim$(sourceText)
                Exit For
            End If
        End If
    Next i

    RemoveProjectContractWeekday = Trim$(sourceText)
End Function
Private Sub TextBox1_Change(): RefreshList Me.TextBox1.text: End Sub

Private Sub RefreshList(ByVal keyword As String)
    Dim i As Long
    Dim itmX As ListItem
    If IsEmpty(SharedMasterData) Then Exit Sub
    Me.ListView1.ListItems.Clear
    For i = 1 To UBound(SharedMasterData, 1)
        If (SharedMasterData(i, 1) & SharedMasterData(i, 2) <> "") And (keyword = "" Or InStr(1, SharedMasterData(i, 1) & SharedMasterData(i, 2), keyword, vbTextCompare) > 0) Then
            Set itmX = Me.ListView1.ListItems.Add(, , "")
            itmX.Tag = CStr(i)
            itmX.SubItems(1) = SharedMasterData(i, 1)
            itmX.SubItems(2) = SharedMasterData(i, 2)
            If UBound(SharedMasterData, 2) >= 3 Then itmX.SubItems(3) = CStr(SharedMasterData(i, 3))
            If UBound(SharedMasterData, 2) >= 4 Then itmX.SubItems(4) = CStr(SharedMasterData(i, 4))
            If UBound(SharedMasterData, 2) >= 5 Then itmX.SubItems(5) = CStr(SharedMasterData(i, 5))
            If UBound(SharedMasterData, 2) >= 6 Then itmX.SubItems(6) = CStr(SharedMasterData(i, 6))
        End If
    Next i
End Sub

Private Sub CommandButton1_Click(): SetSelectedValue: End Sub
Private Sub ListView1_DblClick(): SetSelectedValue: End Sub

Private Sub SetSelectedValue()
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub

    Dim projectNo As String
    Dim projectName As String
    projectNo = Me.ListView1.SelectedItem.SubItems(1)
    projectName = Me.ListView1.SelectedItem.SubItems(2)
    SetBasicInfoProjectSelection projectNo, _
                                 projectName, _
                                 CLng(Me.ListView1.SelectedItem.Tag), _
                                 Me.ListView1.SelectedItem.SubItems(3), _
                                 Me.ListView1.SelectedItem.SubItems(4), _
                                 Me.ListView1.SelectedItem.SubItems(5), _
                                 Me.ListView1.SelectedItem.SubItems(6)
End Sub

Private Function ProjectStatusParentFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5E33) & ChrW$(&H7968) & "7_" & _
                 ChrW$(&H652F) & ChrW$(&H6255) & ChrW$(&H91D1) & ChrW$(&H984D) & _
                 ChrW$(&H8A08) & ChrW$(&H7B97) & ChrW$(&H30B7) & ChrW$(&H30FC) & ChrW$(&H30C8) & _
                 " - " & ChrW$(&H5E33) & ChrW$(&H7968) & "7_" & _
                 ChrW$(&H652F) & ChrW$(&H6255) & ChrW$(&H91D1) & ChrW$(&H984D) & _
                 ChrW$(&H8A08) & ChrW$(&H7B97) & ChrW$(&H30B7) & ChrW$(&H30FC) & ChrW$(&H30C8)
    End If
    ProjectStatusParentFolderText = cached
End Function

Private Function ProjectStatusDataSubFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H3010) & ChrW$(&H5404) & ChrW$(&H652F) & ChrW$(&H5E97) & _
                 ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H756A) & ChrW$(&H53F7) & _
                 ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF) & ChrW$(&H3011)
    End If
    ProjectStatusDataSubFolderText = cached
End Function

Private Function ProjectStatusFileSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = "_" & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H73FE) & ChrW$(&H6CC1) & _
                 ChrW$(&H8868) & ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF) & ".xlsx"
    End If
    ProjectStatusFileSuffixText = cached
End Function

Private Function ProjectStatusColumnProjectNoText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H756A) & ChrW$(&H53F7)
    End If
    ProjectStatusColumnProjectNoText = cached
End Function

Private Function ProjectStatusColumnProjectNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H4EF6) & ChrW$(&H540D)
    End If
    ProjectStatusColumnProjectNameText = cached
End Function

Private Function ProjectStatusDateNumberFormatText() As String
    Static cached As String
    If cached = "" Then
        cached = "yyyy" & ChrW$(&H5E74) & "m" & ChrW$(&H6708) & "d" & ChrW$(&H65E5)
    End If
    ProjectStatusDateNumberFormatText = cached
End Function

Private Function ProjectStatusYearSuffixText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H5E74)
    ProjectStatusYearSuffixText = cached
End Function

Private Function ProjectStatusMonthSuffixText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H6708)
    ProjectStatusMonthSuffixText = cached
End Function

Private Function ProjectStatusDaySuffixText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H65E5)
    ProjectStatusDaySuffixText = cached
End Function

Private Function ProjectStatusFullWidthOpenParenText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&HFF08)
    ProjectStatusFullWidthOpenParenText = cached
End Function

Private Function ProjectStatusFullWidthSpaceText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H3000)
    ProjectStatusFullWidthSpaceText = cached
End Function

Private Function ProjectStatusWeekdayNamesLongArray() As Variant
    ProjectStatusWeekdayNamesLongArray = Array( _
        ProjectStatusWeekdayLongText(1), ProjectStatusWeekdayLongText(2), _
        ProjectStatusWeekdayLongText(3), ProjectStatusWeekdayLongText(4), _
        ProjectStatusWeekdayLongText(5), ProjectStatusWeekdayLongText(6), _
        ProjectStatusWeekdayLongText(7), ProjectStatusWeekdayMediumText(1), _
        ProjectStatusWeekdayMediumText(2), ProjectStatusWeekdayMediumText(3), _
        ProjectStatusWeekdayMediumText(4), ProjectStatusWeekdayMediumText(5), _
        ProjectStatusWeekdayMediumText(6), ProjectStatusWeekdayMediumText(7))
End Function

Private Function ProjectStatusWeekdayNamesShortArray() As Variant
    ProjectStatusWeekdayNamesShortArray = Array( _
        ProjectStatusWeekdayShortText(1), ProjectStatusWeekdayShortText(2), _
        ProjectStatusWeekdayShortText(3), ProjectStatusWeekdayShortText(4), _
        ProjectStatusWeekdayShortText(5), ProjectStatusWeekdayShortText(6), _
        ProjectStatusWeekdayShortText(7))
End Function

Private Function ProjectStatusWeekdayLongText(ByVal weekdayIndex As Long) As String
    ProjectStatusWeekdayLongText = ProjectStatusWeekdayShortText(weekdayIndex) & _
                                   ChrW$(&H66DC) & ChrW$(&H65E5)
End Function

Private Function ProjectStatusWeekdayMediumText(ByVal weekdayIndex As Long) As String
    ProjectStatusWeekdayMediumText = ProjectStatusWeekdayShortText(weekdayIndex) & ChrW$(&H66DC)
End Function

Private Function ProjectStatusWeekdayShortText(ByVal weekdayIndex As Long) As String
    Select Case weekdayIndex
        Case 1: ProjectStatusWeekdayShortText = ChrW$(&H6708)
        Case 2: ProjectStatusWeekdayShortText = ChrW$(&H706B)
        Case 3: ProjectStatusWeekdayShortText = ChrW$(&H6C34)
        Case 4: ProjectStatusWeekdayShortText = ChrW$(&H6728)
        Case 5: ProjectStatusWeekdayShortText = ChrW$(&H91D1)
        Case 6: ProjectStatusWeekdayShortText = ChrW$(&H571F)
        Case 7: ProjectStatusWeekdayShortText = ChrW$(&H65E5)
    End Select
End Function

Private Sub LogProjectSelection(ByVal msg As String)
    mod_DebugLog.Log "[ProjSel] " & msg
End Sub

Private Function BuildProjectStatusNotFoundDetailText(ByVal sourceFilePath As String, _
                                                      ByVal targetBranchOffice As String, _
                                                      ByVal targetYear As Long) As String
    BuildProjectStatusNotFoundDetailText = ProjectStatusDetailFileLabelText() & sourceFilePath & vbCrLf & _
        ProjectStatusDetailSearchKeyLabelText() & targetBranchOffice & vbCrLf & _
        ProjectStatusDetailYearLabelText() & CStr(targetYear)
    If targetYear <= 0 Then
        BuildProjectStatusNotFoundDetailText = BuildProjectStatusNotFoundDetailText & vbCrLf & _
            ProjectStatusDetailYearHintText()
    End If
End Function

Private Function ProjectStatusDetailFileLabelText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H53C2) & ChrW$(&H7167) & ChrW$(&H30D5) & ChrW$(&H30A1) & ChrW$(&H30A4) & ChrW$(&H30EB) & ": "
    ProjectStatusDetailFileLabelText = cached
End Function

Private Function ProjectStatusDetailSearchKeyLabelText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H691C) & ChrW$(&H7D22) & ChrW$(&H30AD) & ChrW$(&H30FC) & "(G" & ChrW$(&H5217) & "): "
    ProjectStatusDetailSearchKeyLabelText = cached
End Function

Private Function ProjectStatusDetailYearLabelText() As String
    Static cached As String
    If cached = "" Then cached = "B4" & ChrW$(&H5E74) & ChrW$(&H5EA6) & ": "
    ProjectStatusDetailYearLabelText = cached
End Function

Private Function ProjectStatusDetailYearHintText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H203B) & "B4" & ChrW$(&H306B) & "4" & ChrW$(&H6841) & ChrW$(&H306E) & _
                 ChrW$(&H897F) & ChrW$(&H66A6) & ChrW$(&H5E74) & "(2026" & ChrW$(&H306A) & ChrW$(&H3069) & ")" & _
                 ChrW$(&H3092) & ChrW$(&H5165) & ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H304F) & _
                 ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044)
    End If
    ProjectStatusDetailYearHintText = cached
End Function
