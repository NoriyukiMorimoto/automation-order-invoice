Option Explicit

'==========================================================================
'  出張所長名 自動入力／支店・出張所バリデーション再構築モジュール
'    改修内容：
'      #7 : WriteValidationLists での Dictionary→セル単位書き込みを
'           Variant 2次元配列＋Range 一括代入へ置換。
'      #9 : NormalizeText / CommonGetBasicInfoWorksheet / 日本語名生成 /
'           ADO 接続生成は mod_Common に集約。重複定義を撤去。
'      #10: CommitOfficeComboBoxSelection で C6 書き込み時に
'           Worksheet_Change の C6 処理が再起動するのを防ぐため
'           mSuppressC6Change フラグを追加。
'      #15: CommitOfficeComboBoxSelection の「現在値と同じならスキップ」
'           ガードを撤廃。LinkedCell が C6 を自動書き換えるため、
'           選択変更時に StrComp が常に False になり FillManagerName が
'           呼ばれない問題を修正。フラグは常に ON にして Change イベントの
'           二重発火を抑制する。
'      #16: 出張所長リストファイル名を
'           「年度_線路出張所長リスト.xlsx」（アンダーバー後スペースなし）優先に変更。
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

Public Function IsSuppressingC6Change() As Boolean
    IsSuppressingC6Change = mSuppressC6Change
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

Public Sub RefreshBranchOfficeValidation(Optional ByVal keepOffice As Boolean = True)
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

    Dim sourceFilePath As String
    sourceFilePath = GetManagerListFilePath(yearText)
    If sourceFilePath = "" Then Exit Sub

    Dim rows As Collection
    Set rows = LoadManagerListRows(sourceFilePath)
    If rows Is Nothing Then
        MsgBox "出張所長リストファイルを参照できませんでした。" & vbCrLf & sourceFilePath, vbExclamation
        Exit Sub
    End If

    Dim branchList As Object, officeList As Object
    Set branchList = CreateObject("Scripting.Dictionary")
    Set officeList = CreateObject("Scripting.Dictionary")
    branchList.CompareMode = vbTextCompare
    officeList.CompareMode = vbTextCompare

    Dim selectedBranch As String
    selectedBranch = CommonNormalizeText(CStr(wsInfo.Range("B6").Value))

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

    WriteValidationLists wsInfo, branchList, officeList, keepOffice
End Sub

'--------------------------------------------------------------------------
'  支店・出張所バリデーションリスト書き込み
'  改修（#7）：Dictionary.Keys() を 2次元配列に詰めて Range 一括代入する。
'              旧版はセル単位 .Cells(...).Value = ... のループだった。
'--------------------------------------------------------------------------
Private Sub WriteValidationLists(ByVal wsInfo As Worksheet, _
                                 ByVal branchList As Object, _
                                 ByVal officeList As Object, _
                                 ByVal keepOffice As Boolean)
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
        ScheduleOfficeComboBoxPrompt
    End If
End Sub

' Dictionary の Keys を 2次元配列に展開し、指定列に Range 一括代入する。
' 件数 0 件のときは何も書かない（呼び出し側で ClearContents 済み）。
Private Sub WriteDictionaryKeysToColumn(ByVal wsInfo As Worksheet, _
                                         ByVal dict As Object, _
                                         ByVal colLetter As String)
    If dict Is Nothing Then Exit Sub
    Dim total As Long
    total = dict.Count
    If total = 0 Then Exit Sub

    Dim keysArr As Variant
    keysArr = dict.Keys ' 1 回だけ呼び出してキャッシュする

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

Private Sub ScheduleOfficeComboBoxPrompt()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!PromptOfficeComboBox"
    If Err.Number <> 0 Then
        Err.Clear
        PromptOfficeComboBox
    End If
    On Error GoTo 0
End Sub

Public Sub PromptOfficeComboBox()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    ShowOfficeComboBox wsInfo
End Sub

'--------------------------------------------------------------------------
'  CommitOfficeComboBoxSelection  (#15 修正)
'    ComboBox で選択確定した出張所名を C6 に反映し、担当者名を更新する。
'
'    【修正前の問題】
'      LinkedCell = "C6" の設定により、ドロップダウンで選択した瞬間に
'      Excel が自動で C6 を書き換える。そのため CommitOfficeComboBoxSelection
'      が呼ばれた時点では「C6 の現在値 = selectedOffice」になっており、
'      旧ガード条件（StrComp <> 0 の場合のみ書き込む）が常に False と
'      判定されて FillManagerNameToBasicInfo が呼ばれなかった。
'      さらに SilentClearBasicInfo は Worksheet_Change 経由で実行済みのため
'      担当者名が空のままになっていた。
'
'    【修正内容】
'      - 「現在値と同じならスキップ」ガードを撤廃し、常に処理を続行する。
'      - LinkedCell による自動書き込みも Worksheet_Change を発火させるため、
'        mSuppressC6Change フラグを ComboBox 非表示の直前（処理全体の開始時）
'        に立てることで、Change イベントの二重発火を防ぐ。
'--------------------------------------------------------------------------
Public Sub CommitOfficeComboBoxSelection()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then Exit Sub

    On Error GoTo ExitHandler

    Dim ole As OLEObject
    Set ole = wsInfo.OLEObjects(OFFICE_COMBO_NAME)
    If ole Is Nothing Then Exit Sub

    Dim selectedOffice As String
    selectedOffice = CommonNormalizeText(CStr(ole.Object.Value))
    If selectedOffice = "" Then Exit Sub

    ' フラグを先に ON にして LinkedCell による C6 自動書き込みの
    ' Change イベントも含めて抑制する
    mSuppressC6Change = True

    ' ComboBox を非表示にしてから C6 に確定値を書き込む
    ' （LinkedCell 経由で既に書き換わっている場合も含め、確実に確定値を設定）
    ole.Visible = False
    wsInfo.Range("C6").Value = selectedOffice
    mSuppressC6Change = False

    FillManagerNameToBasicInfo
    wsInfo.Range("C6").Select

ExitHandler:
    mSuppressC6Change = False
End Sub

Private Sub ShowOfficeComboBox(ByVal wsInfo As Worksheet)
    On Error Resume Next

    Dim ole As OLEObject
    Set ole = GetOfficeComboBox(wsInfo)
    If ole Is Nothing Then
        ShowC6ValidationDropdown wsInfo
        Exit Sub
    End If

    FitOfficeComboBoxToC6 wsInfo, ole
    wsInfo.Activate
    wsInfo.Range("C6").Select
    ole.Visible = True
    ole.Activate
    ole.Object.DropDown
    If Err.Number <> 0 Then
        Err.Clear
        ole.Visible = False
        ShowC6ValidationDropdown wsInfo
    End If

    On Error GoTo 0
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
        .LinkedCell = wsInfo.Range("C6").Address(False, False)
        .ListRows = Application.Max(1, Application.Min(12, officeList.Count))
        .MatchRequired = False
        .Value = CStr(wsInfo.Range("C6").Value)
    End With
    On Error GoTo 0
End Sub

Private Function LoadManagerListRows(ByVal sourceFilePath As String) As Collection
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

    If Not rs.EOF Then rs.MoveNext ' 先頭行（見出し）スキップ
    Do Until rs.EOF
        rows.Add Array(CommonNzText(rs.Fields(0).Value), _
                       CommonNzText(rs.Fields(1).Value), _
                       CommonNzText(rs.Fields(2).Value))
        rs.MoveNext
    Loop

    Set LoadManagerListRows = rows

Cleanup:
    CommonCloseAdoRecordset rs
    CommonCloseAdoConnection cn
    Exit Function

ErrorHandler:
    Set LoadManagerListRows = Nothing
    Resume Cleanup
End Function

' ADO スキーマからワークシート名（先頭シート相当）を取得する。
' Common モジュールのスキーマ一覧と同じロジックを使い、最初に見つけた
' 名前を「[シート名$]」形式で返す。
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
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "出張所長リストフォルダが見つかりません。" & vbCrLf & folderPath, vbExclamation
        Exit Function
    End If

    GetManagerListFilePath = FindManagerListFile(folderPath, yearText)
    If GetManagerListFilePath = "" Then
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

    GetManagerListFolderPath = FirstExistingManagerListFolderPath(candidates)
End Function

Private Function BuildManagerListFolderPath(ByVal documentRootPath As String, _
                                            ByVal fso As Object) As String
    If Len(Trim$(documentRootPath)) = 0 Then Exit Function

    Dim folderPath As String
    folderPath = fso.BuildPath(documentRootPath, UnitPriceMasterFolderText())
    folderPath = fso.BuildPath(folderPath, UnitPriceReferenceFolderText())
    folderPath = fso.BuildPath(folderPath, ManagerNameFolderText())
    BuildManagerListFolderPath = folderPath & Chr$(92)
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
'    線路出張所用_注文書_請求書アクセスサイト - ドキュメント
'    単価マスタ
'    工事件名別マスタ
'    出張所長名
'    線路出張所リスト
'    本社／支店名／出張所名／出張所・支所名
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
