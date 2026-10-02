Option Explicit

' 工事番号選択(基本情報C9)の一覧データを先読み・保持する。
'   工事現況表データは SharePoint 同期フォルダにあり、ADO で開くたびに時間がかかる。
'   1) 読んだ結果をこのモジュールに保持し、フォームを閉じても使い回す
'   2) 結果を %LOCALAPPDATA% に保存し、元ファイルの更新日時が変わるまで ADO を開かない
'   3) ブックを開いた直後と B4/C6 変更後に OnTime で先読みする
'   目印は「ファイルパス + 更新日時 + サイズ + 検索キー(G列)」。どれかが変われば読み直す。
'   ※OnTime は Excel 本体と同じスレッドで動くため、先読み中は Excel を操作できない。

Public Const PROJECT_LIST_OK As Long = 0
Public Const PROJECT_LIST_NO_FILE As Long = 1
Public Const PROJECT_LIST_READ_FAILED As Long = 2

Private Const CACHE_DIR_NAME As String = "OrderInvoiceAutomation"
Private Const CACHE_FILE_NAME As String = "project_list.txt"
Private Const CACHE_FORMAT_VERSION As String = "v1"
Private Const PREFETCH_PROC_NAME As String = "RunProjectListPrefetch"
Private Const PROJECT_DATA_COLS As Long = 6

Private mProjectData As Variant
Private mProjectHitCount As Long
Private mProjectStamp As String
Private mPrefetchAt As Date
Private mPrefetchScheduled As Boolean

' 一覧データを返す。メモリ → 端末キャッシュ → 工事現況表(ADO) の順に探す。
'   outData は (1 To n, 1 To 6): 工事番号 / 件名 / 件名 / 契約日 / 工期開始 / 工期終了
'   該当0件のときは空行1行の配列で outHitCount = 0。
Public Function LoadProjectList(ByVal targetYear As Long, ByVal BranchName As String, ByVal OfficeName As String, _
                                ByRef outData As Variant, ByRef outHitCount As Long, _
                                ByRef sourceFilePath As String, ByRef searchKey As String, _
                                ByRef errText As String) As Long
    Dim branchForFile As String
    Dim stamp As String
    Dim loadedFrom As String

    outData = Empty
    outHitCount = 0
    errText = ""

    searchKey = GetBranchOfficeSearchKey(BranchName, OfficeName)
    branchForFile = GetBranchNameForFile(BranchName)
    sourceFilePath = mod_common.CommonFindProjectStatusSourceFilePath(CStr(targetYear), branchForFile)
    If Len(sourceFilePath) = 0 And targetYear > 0 And Len(branchForFile) > 0 Then
        sourceFilePath = mod_common.CommonGetProjectStatusDataFolderPath() & CStr(targetYear) & "_" & _
            branchForFile & mod_common.CommonProjectStatusFileSuffixText()
    End If
    If Len(sourceFilePath) = 0 Or Not mod_common.CommonProjectStatusFileExists(sourceFilePath) Then
        LoadProjectList = PROJECT_LIST_NO_FILE
        Exit Function
    End If

    stamp = BuildProjectListStamp(sourceFilePath, searchKey)

    If Len(stamp) > 0 And stamp = mProjectStamp And IsArray(mProjectData) Then
        loadedFrom = "memory"
    ElseIf ReadProjectListCache(stamp) Then
        loadedFrom = "cache"
    ElseIf ReadProjectListFromSource(sourceFilePath, searchKey, stamp, errText) Then
        loadedFrom = "ado"
    Else
        LoadProjectList = PROJECT_LIST_READ_FAILED
        Exit Function
    End If

    outData = mProjectData
    outHitCount = mProjectHitCount
    LoadProjectList = PROJECT_LIST_OK
    mod_DebugLog.Log "[ProjSel] source=" & loadedFrom & " hits=" & CStr(mProjectHitCount)
End Function

' ブックを開いた直後・B4/C6 変更後に呼ぶ。1秒後に先読みを予約する。
Public Sub ScheduleProjectListPrefetch()
    CancelScheduledProjectListPrefetch
    On Error Resume Next
    mPrefetchAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime mPrefetchAt, "'" & ThisWorkbook.Name & "'!" & PREFETCH_PROC_NAME
    mPrefetchScheduled = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Sub

' ブックを閉じるときに呼ぶ(予約が残っていると閉じた直後にブックが開き直される)。
Public Sub CancelScheduledProjectListPrefetch()
    If Not mPrefetchScheduled Then Exit Sub
    On Error Resume Next
    Application.OnTime mPrefetchAt, "'" & ThisWorkbook.Name & "'!" & PREFETCH_PROC_NAME, , False
    Err.Clear
    On Error GoTo 0
    mPrefetchScheduled = False
End Sub

' OnTime から呼ばれる先読み本体。エラーやメッセージは出さない(通常操作を妨げない)。
Public Sub RunProjectListPrefetch()
    mPrefetchScheduled = False

    Dim previousStatusBar As Variant
    Dim statusShown As Boolean
    On Error GoTo Done

    Dim ws As Worksheet
    Set ws = mod_common.CommonGetBasicInfoWorksheet(ThisWorkbook)
    If ws Is Nothing Then Exit Sub

    Dim targetYear As Long
    Dim BranchName As String
    Dim OfficeName As String
    targetYear = Val(mod_common.CommonExtractYear4Digits(CStr(ws.Range("B4").value)))
    BranchName = Trim$(CStr(ws.Range("B6").value))
    OfficeName = mod_BasicInfoCalendar.NormalizeProjectSelectionOfficeName(Trim$(CStr(ws.Range("C6").value)))
    If targetYear <= 0 Or Len(BranchName) = 0 Or Len(OfficeName) = 0 Then Exit Sub

    previousStatusBar = Application.StatusBar
    Application.StatusBar = PrefetchStatusText()
    statusShown = True

    Dim t0 As Double
    Dim dataOut As Variant
    Dim hitCount As Long
    Dim sourceFilePath As String
    Dim searchKey As String
    Dim errText As String
    Dim result As Long
    t0 = Timer
    result = LoadProjectList(targetYear, BranchName, OfficeName, dataOut, hitCount, sourceFilePath, searchKey, errText)
    mod_DebugLog.Log "[ProjSel] prefetch result=" & CStr(result) & " sec=" & Format$(Timer - t0, "0.00") & _
                     " " & errText

Done:
    On Error Resume Next
    If statusShown Then
        If VarType(previousStatusBar) = vbBoolean Then
            Application.StatusBar = False
        Else
            Application.StatusBar = previousStatusBar
        End If
    End If
    On Error GoTo 0
End Sub

Private Function BuildProjectListStamp(ByVal sourceFilePath As String, ByVal searchKey As String) As String
    On Error GoTo ErrorHandler
    BuildProjectListStamp = sourceFilePath & "*" & CStr(CDbl(FileDateTime(sourceFilePath))) & "*" & _
                            CStr(FileLen(sourceFilePath)) & "#" & searchKey & "#" & CACHE_FORMAT_VERSION
    Exit Function

ErrorHandler:
    BuildProjectListStamp = ""
End Function

' 工事現況表を ADO で読み、G列(施工課所)が検索キーを含む行だけを残す(下の行から順)。
Private Function ReadProjectListFromSource(ByVal sourceFilePath As String, ByVal searchKey As String, _
                                           ByVal stamp As String, ByRef errText As String) As Boolean
    On Error GoTo ErrorHandler

    Dim sourceArr As Variant
    If Not ReadProjectStatusFileToArray(sourceFilePath, "Sheet1", sourceArr, errText) Then Exit Function

    Dim i As Long
    Dim hitCount As Long
    If IsArray(sourceArr) Then
        For i = UBound(sourceArr, 1) To 1 Step -1
            If IsSearchKeyHit(sourceArr, i, searchKey) Then hitCount = hitCount + 1
        Next i
    End If

    Dim result() As Variant
    ReDim result(1 To Application.Max(1, hitCount), 1 To PROJECT_DATA_COLS)
    Dim c As Long
    For c = 1 To PROJECT_DATA_COLS
        result(1, c) = ""
    Next c

    If hitCount > 0 Then
        Dim writeIndex As Long
        writeIndex = 1
        For i = UBound(sourceArr, 1) To 1 Step -1
            If IsSearchKeyHit(sourceArr, i, searchKey) Then
                result(writeIndex, 1) = GetSourceText(sourceArr, i, 10)
                result(writeIndex, 2) = RemoveSpaces(GetSourceText(sourceArr, i, 29) & GetSourceText(sourceArr, i, 30))
                result(writeIndex, 3) = result(writeIndex, 2)
                result(writeIndex, 4) = GetSourceText(sourceArr, i, 69)
                result(writeIndex, 5) = GetSourceText(sourceArr, i, 34)
                result(writeIndex, 6) = GetSourceText(sourceArr, i, 35)
                writeIndex = writeIndex + 1
            End If
        Next i
    End If

    mProjectData = result
    mProjectHitCount = hitCount
    mProjectStamp = stamp
    WriteProjectListCache stamp
    ReadProjectListFromSource = True
    Exit Function

ErrorHandler:
    errText = Err.Description
    ReadProjectListFromSource = False
End Function

Private Function IsSearchKeyHit(ByVal sourceArr As Variant, ByVal rowIndex As Long, ByVal searchKey As String) As Boolean
    If searchKey = "" Then
        IsSearchKeyHit = True
    Else
        IsSearchKeyHit = (InStr(1, RemoveSpaces(GetSourceText(sourceArr, rowIndex, 7)), searchKey, vbTextCompare) > 0)
    End If
End Function

Private Function ReadProjectStatusFileToArray(ByVal sourcePath As String, ByVal sheetName As String, _
                                              ByRef outArr As Variant, ByRef errText As String) As Boolean
    Dim cn As Object
    Dim rs As Object
    outArr = Empty

    Set cn = mod_common.CommonOpenExcelAdoConnection(sourcePath)
    If cn Is Nothing Then
        errText = "ADO connection failed"
        Exit Function
    End If

    On Error GoTo ErrorHandler

    Dim adoSheetName As String
    adoSheetName = ResolveProjectStatusAdoSheetName(cn, sheetName)
    If adoSheetName = "" Then
        errText = "sheet not found"
        GoTo Cleanup
    End If

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & adoSheetName & "$]", cn, 0, 1
    If Not rs.EOF Then outArr = ConvertAdoRecordsetToRowMajorArray(rs)
    ReadProjectStatusFileToArray = True
    GoTo Cleanup

ErrorHandler:
    errText = Err.Description
    ReadProjectStatusFileToArray = False

Cleanup:
    mod_common.CommonCloseAdoRecordset rs
    mod_common.CommonCloseAdoConnection cn
End Function

Private Function ResolveProjectStatusAdoSheetName(ByVal cn As Object, ByVal preferredSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = mod_common.CommonGetAdoWorksheetNames(cn)
    If sheetNames Is Nothing Then Exit Function
    If sheetNames.Count = 0 Then Exit Function

    Dim i As Long
    For i = 1 To sheetNames.Count
        If StrComp(CStr(sheetNames(i)), preferredSheetName, vbTextCompare) = 0 Then
            ResolveProjectStatusAdoSheetName = CStr(sheetNames(i))
            Exit Function
        End If
    Next i

    ResolveProjectStatusAdoSheetName = CStr(sheetNames(1))
End Function

Private Function ConvertAdoRecordsetToRowMajorArray(ByVal rs As Object) As Variant
    Dim data As Variant
    data = rs.GetRows

    Dim fieldCount As Long
    Dim recordCount As Long
    fieldCount = UBound(data, 1) + 1
    recordCount = UBound(data, 2) + 1
    If fieldCount <= 0 Or recordCount <= 0 Then Exit Function

    Dim result() As Variant
    ReDim result(1 To recordCount, 1 To fieldCount)

    Dim rowIndex As Long
    Dim colIndex As Long
    For rowIndex = 1 To recordCount
        For colIndex = 1 To fieldCount
            result(rowIndex, colIndex) = data(colIndex - 1, rowIndex - 1)
        Next colIndex
    Next rowIndex

    ConvertAdoRecordsetToRowMajorArray = result
End Function

' Null・エラー値は空文字にする(フォーム側は表示時に CStr するため文字列で揃える)
Private Function GetSourceText(ByVal sourceArr As Variant, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    On Error GoTo ErrorHandler
    If colIndex > UBound(sourceArr, 2) Then Exit Function
    If IsNull(sourceArr(rowIndex, colIndex)) Or IsError(sourceArr(rowIndex, colIndex)) Then Exit Function
    GetSourceText = CStr(sourceArr(rowIndex, colIndex))
    Exit Function

ErrorHandler:
    GetSourceText = ""
End Function

'--- 端末キャッシュ(1行目=目印、2行目以降=タブ区切り6列) ---

Private Function ProjectListCacheFilePath() As String
    Dim baseDir As String
    baseDir = Environ$("LOCALAPPDATA")
    If baseDir = "" Then baseDir = Environ$("TEMP")
    If baseDir = "" Then Exit Function

    On Error GoTo ErrorHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    folderPath = baseDir & Chr$(92) & CACHE_DIR_NAME
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    ProjectListCacheFilePath = folderPath & Chr$(92) & CACHE_FILE_NAME
    Exit Function

ErrorHandler:
    ProjectListCacheFilePath = ""
End Function

Private Function ReadProjectListCache(ByVal stamp As String) As Boolean
    If Len(stamp) = 0 Then Exit Function

    Dim filePath As String
    filePath = ProjectListCacheFilePath()
    If Len(filePath) = 0 Then Exit Function

    Dim fso As Object
    Dim ts As Object
    Dim lines As Collection
    Set lines = New Collection

    On Error GoTo Cleanup
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then GoTo Cleanup
    Set ts = fso.OpenTextFile(filePath, 1, False, -1)
    If ts.AtEndOfStream Then GoTo Cleanup
    If ts.ReadLine <> stamp Then GoTo Cleanup
    Do While Not ts.AtEndOfStream
        lines.Add ts.ReadLine
    Loop

    Dim result() As Variant
    ReDim result(1 To Application.Max(1, lines.Count), 1 To PROJECT_DATA_COLS)
    Dim c As Long
    For c = 1 To PROJECT_DATA_COLS
        result(1, c) = ""
    Next c

    Dim i As Long
    Dim parts() As String
    For i = 1 To lines.Count
        parts = Split(CStr(lines(i)), vbTab)
        If UBound(parts) <> PROJECT_DATA_COLS - 1 Then GoTo Cleanup
        For c = 1 To PROJECT_DATA_COLS
            result(i, c) = parts(c - 1)
        Next c
    Next i

    mProjectData = result
    mProjectHitCount = lines.Count
    mProjectStamp = stamp
    ReadProjectListCache = True

Cleanup:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    On Error GoTo 0
End Function

Private Sub WriteProjectListCache(ByVal stamp As String)
    If Len(stamp) = 0 Then Exit Sub

    Dim filePath As String
    filePath = ProjectListCacheFilePath()
    If Len(filePath) = 0 Then Exit Sub

    Dim fso As Object
    Dim ts As Object
    On Error GoTo Cleanup
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.WriteLine stamp

    Dim i As Long
    Dim c As Long
    Dim lineText As String
    For i = 1 To mProjectHitCount
        lineText = ""
        For c = 1 To PROJECT_DATA_COLS
            If c > 1 Then lineText = lineText & vbTab
            lineText = lineText & SanitizeCacheField(CStr(mProjectData(i, c)))
        Next c
        ts.WriteLine lineText
    Next i

Cleanup:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    On Error GoTo 0
End Sub

Private Function SanitizeCacheField(ByVal value As String) As String
    value = Replace$(value, vbTab, " ")
    value = Replace$(value, vbCr, " ")
    value = Replace$(value, vbLf, " ")
    SanitizeCacheField = value
End Function

'--- 検索キー・ファイル名(Project_Number_Selection から移設) ---

Private Function GetBranchNameForFile(ByVal BranchName As String) As String
    BranchName = RemoveSpaces(BranchName)
    If BranchName = "" Then Exit Function

    If Right$(BranchName, 2) = BranchSuffixText() Then
        GetBranchNameForFile = BranchName
    Else
        GetBranchNameForFile = BranchName & BranchSuffixText()
    End If
End Function

Private Function GetBranchOfficeSearchKey(ByVal BranchName As String, ByVal OfficeName As String) As String
    Dim normalizedBranch As String
    Dim normalizedOffice As String

    normalizedBranch = GetBranchNameForFile(BranchName)
    normalizedOffice = RemoveSpaces(OfficeName)

    If StrComp(normalizedBranch, KobeBranchText(), vbTextCompare) = 0 And _
       StrComp(normalizedOffice, SanyoShinkansenTrackMaintenanceOfficeText(), vbTextCompare) = 0 Then
        GetBranchOfficeSearchKey = KobeSanyoShinkansenTrackSearchText()
    Else
        GetBranchOfficeSearchKey = RemoveSpaces(normalizedBranch & normalizedOffice)
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
    Static cached As String
    If cached = "" Then cached = ChrW$(&H652F) & ChrW$(&H5E97)
    BranchSuffixText = cached
End Function

Private Function RemoveSpaces(ByVal value As String) As String
    value = Replace$(value, " ", "")
    value = Replace$(value, ChrW$(&H3000), "")
    value = Replace$(value, vbTab, "")
    value = Replace$(value, vbCr, "")
    value = Replace$(value, vbLf, "")
    RemoveSpaces = value
End Function

' 工事番号一覧を読込中です…
Private Function PrefetchStatusText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H756A) & ChrW$(&H53F7) & ChrW$(&H4E00) & ChrW$(&H89A7) & _
                 ChrW$(&H3092) & ChrW$(&H8AAD) & ChrW$(&H8FBC) & ChrW$(&H4E2D) & ChrW$(&H3067) & ChrW$(&H3059) & _
                 ChrW$(&H2026)
    End If
    PrefetchStatusText = cached
End Function
