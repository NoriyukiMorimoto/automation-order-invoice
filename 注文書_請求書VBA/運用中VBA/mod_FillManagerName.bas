Option Explicit

'==========================================================================
'  出張所長名 自動入力／支店・出張所バリデーション再構篈モジュール
'    改修内容：
'      #23: B6変更時コンボ表示フロー追跡用ログを挿入（mod_DebugLog使用）
'           デバッグ確認後は mod_DebugLog の呼び出しを削除すること。
'      #22: B6変更時に C6 コンボが開かない問題を修正。
'           EnableEvents=False 中は OLEObject.Activate が正しく動かないため、
'           PromptOfficeComboBox の処理中だけ一時的に True に切り替える。
'           同時に mInPromptOffice フラグを追加し、Sheet1.cls の
'           SelectionChange / ComboBox1_LostFocus でコンボを消さないようにする。
'           終了時に EnableEvents を呼び出し元の状態へ復元する。
'      #19: B6変更後にC6コンボが表示されない問題を修正。
'           RefreshBranchOfficeValidation の戻り値（Boolean）で
'           「コンボ表示スケジュールが必要か」を Sheet1.cls へ返すよう変更。
'      #7 : WriteValidationLists での Dictionary→セル単位書き込みを
'           Variant 2次元配列＋Range 一括代入へ置換。
'      #9 : NormalizeText / CommonGetBasicInfoWorksheet / 日本語名生成 /
'           ADO 接続生成は mod_Common に集約。重複定義を撤去。
'      #10: CommitOfficeComboBoxSelection で C6 書き込み時に
'           Worksheet_Change の C6 処理が再起動するのを防ぐため
'           mSuppressC6Change フラグを追加。
'      #15: CommitOfficeComboBoxSelection の「現在値と同じならスキップ」
'           ガードを撤廃。フラグは常に ON にして Change イベントの
'           二重発火を抑制する。
'      #16: 出張所長リストファイル名を
'           「年度_線路出張所長リスト.xlsx」優先に変更。
'      #17: 出張所長リストの参照先を
'           単価マスタ\工事件名別マスタ\出張所長名 に変更。
'==========================================================================

Private Const LIST_BRANCH_COL As String = "AA"
Private Const LIST_OFFICE_COL As String = "AB"
Private Const LIST_START_ROW As Long = 2
Private Const OFFICE_COMBO_NAME As String = "ComboBox1"
Private Const OFFICE_COMBO_WIDTH_POINTS As Double = 310.5

' C6 への書き込み中に Worksheet_Change の C6 処理をスキップするフラグ
Private mSuppressC6Change As Boolean
' (#22) PromptOfficeComboBox 処理中を示すフラグ
'       True の間は SelectionChange / LostFocus のコンボ非表示処理をスキップする
Private mInPromptOffice As Boolean

Public Function IsSuppressingC6Change() As Boolean
    IsSuppressingC6Change = mSuppressC6Change
End Function

' (#22) PromptOfficeComboBox 処理中かどうかを Sheet1.cls から参照するための Public Function
Public Function IsPromptingOfficeComboBox() As Boolean
    IsPromptingOfficeComboBox = mInPromptOffice
End Function

Public Sub FillManagerNameToBasicInfo()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つかりません。シート名を確認してください。", vbExclamation
        Exit Sub
    End If

    Dim yearText As String
    yearText = CommonExtractYear4Digits(Trim$(CStr(wsInfo.Range("B4").Value)))
    If yearText = "" Then
        MsgBox "基本情報シート B4 に4桁の年度が見つかりません。例: 2026", vbExclamation
        Exit Sub
    End If

    Dim BranchName As String, OfficeName As String
    BranchName = CommonNormalizeText(CStr(wsInfo.Range("B6").Value))
    OfficeName = CommonNormalizeText(CStr(wsInfo.Range("C6").Value))
    If BranchName = "" Or OfficeName = "" Then
        MsgBox "基本情報シート B6 または C6 が空です。支店名・出張所名を確認してください。", vbExclamation
        Exit Sub
    End If

    Dim sourceFilePath As String
    sourceFilePath = GetManagerListFilePath(yearText)
    If sourceFilePath = "" Then Exit Sub

    Dim rows As Collection
    Set rows = LoadManagerListRows(sourceFilePath)
    If rows Is Nothing Then
        MsgBox "出張所長リストファイルを参照できませんでした。" & vbCrLf & sourceFilePath, vbExclamation
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
        MsgBox "該当する支店名・出張所名が見つかりませんでした。" & vbCrLf & _
               "支店名：" & BranchName & vbCrLf & _
               "出張所名：" & OfficeName, vbExclamation
    Else
        wsInfo.Range("F6").Value = foundName
    End If
End Sub

'--------------------------------------------------------------------------
'  RefreshBranchOfficeValidation
'    戻り値：True = 呼び出し元でコンボ表示処理を呼ぶ必要あり
'             False = 不要
'--------------------------------------------------------------------------
Public Function RefreshBranchOfficeValidation(Optional ByVal keepOffice As Boolean = True) As Boolean
    RefreshBranchOfficeValidation = False
    mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation 開始 keepOffice=" & keepOffice

    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox "基本情報シートが見つかりません。シート名を確認してください。", vbExclamation
        Exit Function
    End If

    Dim yearText As String
    yearText = CommonExtractYear4Digits(Trim$(CStr(wsInfo.Range("B4").Value)))
    If yearText = "" Then
        mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: yearText 空 -> Exit"
        MsgBox "基本情報シート B4 に4桁の年度が見つかりません。例: 2026", vbExclamation
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
        MsgBox "出張所長リストファイルを参照できませんでした。" & vbCrLf & sourceFilePath, vbExclamation
        Exit Function
    End If
    mod_DebugLog.Log "[FillMgr] RefreshBranchOfficeValidation: rows.Count=" & rows.Count

    Dim branchList As Object, officeList As Object
    Set branchList = CreateObject("Scripting.Dictionary")
    Set officeList = CreateObject("Scripting.Dictionary")
    branchList.CompareMode = vbTextCompare
    officeList.CompareMode = vbTextCompare

    Dim selectedBranch As String
    selectedBranch = CommonNormalizeText(CStr(wsInfo.Range("B6").Value))
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
    currentOffice = CommonNormalizeText(CStr(wsInfo.Range("C6").Value))
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

    wsInfo.Range(colLetter & LIST_START_ROW).Resize(total, 1).Value = outArr
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
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!PromptOfficeComboBox"
    If Err.Number <> 0 Then
        Err.Clear
        PromptOfficeComboBox
    End If
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
'  PromptOfficeComboBox  (#22 修正)
'--------------------------------------------------------------------------
Public Sub PromptOfficeComboBox()
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
        ole.Object.LinkedCell = ""
        ole.Object.Value = CStr(wsInfo.Range("C6").Value)
        ole.Visible = False
    End If
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
'  CommitOfficeComboBoxSelection  (#15 修正)
'--------------------------------------------------------------------------
Public Sub CommitOfficeComboBoxSelection(Optional ByVal selectC6 As Boolean = True)
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ExitHandler

    Dim ole As OLEObject
    Set ole = wsInfo.OLEObjects(OFFICE_COMBO_NAME)
    If ole Is Nothing Then GoTo ExitHandler

    Dim selectedOffice As String
    selectedOffice = CommonNormalizeText(CStr(ole.Object.Value))
    If selectedOffice = "" Then GoTo ExitHandler

    mSuppressC6Change = True

    ole.Object.LinkedCell = ""
    wsInfo.Range("C6").Value = selectedOffice
    mSuppressC6Change = False

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

    FitOfficeComboBoxToC6 wsInfo, ole
    wsInfo.Activate
    wsInfo.Range("C6").Select
    ole.Object.LinkedCell = ""
    If ole.Object.ListCount = 0 Then
        mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: ListCount=0 -> ValidationDropdown へ"
        HideOfficeComboBox wsInfo
        ShowC6ValidationDropdown wsInfo
        Exit Sub
    End If
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: ListCount=" & ole.Object.ListCount & " -> DropDown 呼び出し"
    ole.Visible = True
    ole.Activate
    ole.Object.DropDown
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: DropDown 完了 Err=" & Err.Number
    Exit Sub

ErrorHandler:
    mod_DebugLog.Log "[FillMgr] ShowOfficeComboBox: ErrorHandler Err=" & Err.Number & " " & Err.Description
    HideOfficeComboBox wsInfo
    ShowC6ValidationDropdown wsInfo
End Sub

Private Sub FitOfficeComboBoxToC6(ByVal wsInfo As Worksheet, ByVal ole As OLEObject)
    With ole
        .Left = wsInfo.Range("C6").Left
        .Top = wsInfo.Range("C6").Top
        .Width = OFFICE_COMBO_WIDTH_POINTS
        .Height = wsInfo.Range("C6").Height
        .Placement = xlMoveAndSize
    End With
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
    Application.SendKeys "%{DOWN}", True
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
        .LinkedCell = ""
        .ListRows = Application.Max(1, Application.Min(12, officeList.Count))
        .MatchRequired = False
        .Value = CStr(wsInfo.Range("C6").Value)
    End With
    ole.Visible = False
    On Error GoTo 0
End Sub

Private Function LoadManagerListRows(ByVal sourceFilePath As String) As Collection
    Set LoadManagerListRows = LoadManagerListRowsFromAdo(sourceFilePath)
    If LoadManagerListRows Is Nothing Then
        Set LoadManagerListRows = LoadManagerListRowsFromWorkbook(sourceFilePath)
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
        rows.Add Array(CommonNzText(rs.Fields(0).Value), _
                       CommonNzText(rs.Fields(1).Value), _
                       CommonNzText(rs.Fields(2).Value))
        rs.MoveNext
    Loop

    Set LoadManagerListRowsFromAdo = rows

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    Set LoadManagerListRowsFromAdo = Nothing
    Resume Cleanup
End Function

Private Function LoadManagerListRowsFromWorkbook(ByVal sourceFilePath As String) As Collection
    Dim sourceBook As Workbook
    Dim previousDisplayAlerts As Boolean
    previousDisplayAlerts = Application.DisplayAlerts

    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False

    Set sourceBook = Application.Workbooks.Open(Filename:=sourceFilePath, _
                                                UpdateLinks:=False, _
                                                ReadOnly:=True, _
                                                AddToMru:=False)

    Dim sourceSheet As Worksheet
    Set sourceSheet = sourceBook.Worksheets(1)

    Dim rows As Collection
    Set rows = New Collection

    Dim lastRow As Long
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 2).End(xlUp).Row

    Dim rr As Long
    For rr = 2 To lastRow
        rows.Add Array(CommonNzText(sourceSheet.Cells(rr, 2).Value), _
                       CommonNzText(sourceSheet.Cells(rr, 3).Value), _
                       CommonNzText(sourceSheet.Cells(rr, 6).Value))
    Next rr

    Set LoadManagerListRowsFromWorkbook = rows

Cleanup:
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.DisplayAlerts = previousDisplayAlerts
    Exit Function

ErrorHandler:
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
        MsgBox "出張所長リストフォルダが見つかりません。" & vbCrLf & folderPath, vbExclamation
        Exit Function
    End If

    GetManagerListFilePath = FindManagerListFile(folderPath, yearText)
    mod_DebugLog.Log "[FillMgr] GetManagerListFilePath: 結果=[" & GetManagerListFilePath & "]"
    If GetManagerListFilePath = "" Then
        mod_DebugLog.Log "[FillMgr] GetManagerListFilePath: ファイル不存在 -> Exit"
        MsgBox yearText & " 年の出張所長リストファイルが見つかりません。ファイル名に年度が含まれているか確認してください。", vbExclamation
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
    folderPath = fso.BuildPath(documentRootPath, UnitPriceMasterFolderText())
    folderPath = fso.BuildPath(folderPath, UnitPriceReferenceFolderText())
    folderPath = fso.BuildPath(folderPath, ManagerNameFolderText())
    BuildManagerListFolderPath = folderPath & Chr$(92)
    mod_DebugLog.Log "[FillMgr] BuildManagerListFolderPath: root=[" & documentRootPath & "] -> [" & BuildManagerListFolderPath & "]"
End Function

Private Function FirstExistingManagerListFolderPath(ByVal candidates As Collection) As String
    Dim candidate As Variant
    For Each candidate In candidates
        If Len(CStr(candidate)) > 0 Then
            If Dir(CStr(candidate), vbDirectory) <> "" Then
                FirstExistingManagerListFolderPath = CStr(candidate)
                Exit Function
            End If
        End If
    Next candidate
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

'--------------------------------------------------------------------------
'  本モジュール専用の日本語名（Common に共通化していない）
'--------------------------------------------------------------------------

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

Private Function UnitPriceMasterFolderText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H5358) & ChrW$(&H4FA1) & ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF)
    UnitPriceMasterFolderText = cached
End Function

Private Function UnitPriceReferenceFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H4EF6) & ChrW$(&H540D) & _
                 ChrW$(&H5225) & ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF)
    End If
    UnitPriceReferenceFolderText = cached
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
