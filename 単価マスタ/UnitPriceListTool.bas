Option Explicit

Private Const DEFAULT_SOURCE_ROOT As String = "\\dt-ims\公開フォルダ\090_線路本部\008_単価契約関係\2026年度 R8年度単価契約工事"
Private Const PROJECT_LIST_BOOK As String = "出張所別_単価適用線区.xlsx"
Private Const PROJECT_LIST_DIR As String = "工事件名別マスタ"
Private Const PROJECT_LIST_SHEET As String = "単価適用工事件名マスタ"
Private Const PROJECT_MASTER_DIR As String = "工事件名別マスタ"
Private Const OUTPUT_DATA_DIR As String = "単価データ"
Private Const OUTPUT_CONVENTIONAL_LINE_DIR As String = "01_在来線"
Private Const OUTPUT_SHINKANSEN_LINE_DIR As String = "02_新幹線"
Private Const SYNC_COMPANY_FOLDER As String = "大鉄工業株式会社"
Private Const SYNC_LIBRARY_FOLDER As String = "線路出張所用_注文書_請求書アクセスサイト - ドキュメント"
Private Const ROOT_FOLDER_NAME As String = "単価マスタ"
Private Const SOURCE_SEARCH_MAX_DEPTH As Long = 8
Private Const PURCHASE_PROJECT_CODE As String = "10"
Private Const PURCHASE_TITLE_PREFIX_NORMAL As String = "早期発注"
Private Const PURCHASE_TITLE_PREFIX_CHANGE As String = "設計変更"

Private mRootPath As String
Private mLogPath As String
Private mLogBuffer As String
Private mSourceYearText As String
Private mProgDone As Long
Private mProgName As String

Public Sub CreateUnitPriceLists()
    On Error GoTo FatalError

    Dim sourcePath As String
    Dim targetBooks As Object
    Dim branches As Collection
    Dim conventionalBranches As Collection
    Dim shinkansenRoots As Collection
    Dim conventionalProjectNames As Collection
    Dim shinkansenProjectNames As Collection
    Dim totalProjectCount As Long
    Dim branchResolveError As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    sourcePath = SelectSourceFolder()
    If Len(sourcePath) = 0 Then Exit Sub

    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean
    Dim settingsApplied As Boolean

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    settingsApplied = True

    mRootPath = ResolveRootPath()
    mSourceYearText = GetYearFromSourcePath(sourcePath)
    StartLog
    WriteLog "起動しました。"
    WriteLog "単価マスタフォルダ: " & mRootPath
    WriteLog "参照指定: " & sourcePath
    If Len(mSourceYearText) > 0 Then WriteLog "参照年度: " & mSourceYearText

    On Error Resume Next
    Set branches = ResolveBranches(sourcePath)
    branchResolveError = Err.Description
    Err.Clear
    On Error GoTo FatalError
    If branches Is Nothing Then Set branches = New Collection

    Set conventionalBranches = FilterBranchesByLine(branches, OUTPUT_CONVENTIONAL_LINE_DIR)
    Set shinkansenRoots = ResolveShinkansenSourceRoots(sourcePath)
    If conventionalBranches.Count = 0 And shinkansenRoots.Count = 0 Then
        If Len(branchResolveError) > 0 Then
            Err.Raise vbObjectError + 4, , branchResolveError
        Else
            Err.Raise vbObjectError + 4, , "処理対象フォルダが見つかりません。"
        End If
    End If

    WriteLog "在来線処理対象フォルダ数: " & CStr(conventionalBranches.Count)
    WriteLog "新幹線起点フォルダ数: " & CStr(shinkansenRoots.Count)

    Dim projectName As Variant
    Set conventionalProjectNames = ReadProjectNames(PROJECT_LIST_SHEET)
    Set shinkansenProjectNames = ReadProjectNames(ShinkansenProjectListSheetName())
    WriteLog "在来線工事件名数: " & CStr(conventionalProjectNames.Count)
    WriteLog "新幹線工事件名数: " & CStr(shinkansenProjectNames.Count)

    If conventionalBranches.Count > 0 Then totalProjectCount = totalProjectCount + conventionalProjectNames.Count
    If shinkansenRoots.Count > 0 Then totalProjectCount = totalProjectCount + shinkansenProjectNames.Count

    mProgDone = 0
    If totalProjectCount < 1 Then
        frmProgress.ShowProgress 1
    Else
        frmProgress.ShowProgress totalProjectCount
    End If

    FlushLog

    If shinkansenRoots.Count > 0 Then
        WriteLog "新幹線処理開始"
        FlushLog
        For Each projectName In shinkansenProjectNames
            mProgName = CStr(projectName)
            WriteLog "新幹線工事件名処理開始: " & mProgName
            FlushLog
            frmProgress.UpdateProgress mProgDone, mProgName
            Set targetBooks = CreateObject("Scripting.Dictionary")
            ProcessShinkansenProject mProgName, shinkansenRoots, targetBooks
            If targetBooks.Count > 0 Then
                frmProgress.UpdateProgress mProgDone, "保存処理中: " & mProgName
                SaveTargetBooks targetBooks
            End If
            WriteLog "新幹線工事件名処理終了: " & mProgName & " / 作成候補 " & CStr(targetBooks.Count)
            FlushLog
            Set targetBooks = Nothing
            mProgDone = mProgDone + 1
            frmProgress.UpdateProgress mProgDone, mProgName
        Next projectName
        WriteLog "新幹線処理終了"
        FlushLog
    End If

    If conventionalBranches.Count > 0 Then
        WriteLog "在来線処理開始"
        FlushLog
        For Each projectName In conventionalProjectNames
            mProgName = CStr(projectName)
            WriteLog "在来線工事件名処理開始: " & mProgName
            FlushLog
            frmProgress.UpdateProgress mProgDone, mProgName
            Set targetBooks = CreateObject("Scripting.Dictionary")
            ProcessProjectAllBranches mProgName, conventionalBranches, targetBooks
            If targetBooks.Count > 0 Then
                frmProgress.UpdateProgress mProgDone, "保存処理中: " & mProgName
                SaveTargetBooks targetBooks
            End If
            WriteLog "在来線工事件名処理終了: " & mProgName & " / 作成候補 " & CStr(targetBooks.Count)
            FlushLog
            Set targetBooks = Nothing
            mProgDone = mProgDone + 1
            frmProgress.UpdateProgress mProgDone, mProgName
        Next projectName
        WriteLog "在来線処理終了"
        FlushLog
    End If

    frmProgress.CloseProgress
    WriteLog "正常終了しました。"
    FlushLog
    RestoreAppSettings settingsApplied, oldCalc, oldScreenUpdating, oldDisplayAlerts, oldEnableEvents
    MsgBox "作成が完了しました。" & vbCrLf & mLogPath, vbInformation
    Exit Sub

FatalError:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description

    On Error Resume Next
    frmProgress.CloseProgress
    CloseUnsavedTargetBooks targetBooks
    On Error GoTo 0
    FlushLog
    RestoreAppSettings settingsApplied, oldCalc, oldScreenUpdating, oldDisplayAlerts, oldEnableEvents
    If Len(mLogPath) > 0 Then
        WriteErrorLogValues "アプリ全体", savedErrNumber, savedErrSource, savedErrDescription
        FlushLog
        MsgBox "エラーが発生しました。ログを確認してください。" & vbCrLf & mLogPath, vbExclamation
    Else
        MsgBox "エラーが発生しました。" & vbCrLf & savedErrDescription, vbExclamation
    End If
End Sub

Private Sub RestoreAppSettings(ByVal applied As Boolean, ByVal oldCalc As XlCalculation, ByVal oldScreenUpdating As Boolean, ByVal oldDisplayAlerts As Boolean, ByVal oldEnableEvents As Boolean)
    If Not applied Then Exit Sub
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
End Sub

' 中断時: 未保存の作成中ブックを閉じる(開きっぱなし防止)
Private Sub CloseUnsavedTargetBooks(ByVal targetBooks As Object)
    Dim key As Variant
    Dim book As Workbook
    If targetBooks Is Nothing Then Exit Sub
    On Error Resume Next
    For Each key In targetBooks.Keys
        Set book = targetBooks(key)
        If Not book Is Nothing Then book.Close SaveChanges:=False
        Set book = Nothing
    Next key
    On Error GoTo 0
End Sub

Private Function ResolveRootPath() As String
    Dim candidate As String

    candidate = ThisWorkbook.Path
    If IsValidRootPath(candidate) Then
        ResolveRootPath = candidate
        Exit Function
    End If

    candidate = CombinePath(CombinePath(Environ$("OneDriveCommercial"), SYNC_LIBRARY_FOLDER), ROOT_FOLDER_NAME)
    If IsValidRootPath(candidate) Then
        ResolveRootPath = candidate
        Exit Function
    End If

    candidate = CombinePath(CombinePath(CombinePath(Environ$("USERPROFILE"), SYNC_COMPANY_FOLDER), SYNC_LIBRARY_FOLDER), ROOT_FOLDER_NAME)
    If IsValidRootPath(candidate) Then
        ResolveRootPath = candidate
        Exit Function
    End If

    candidate = CombinePath(CombinePath(Environ$("OneDrive"), SYNC_LIBRARY_FOLDER), ROOT_FOLDER_NAME)
    If IsValidRootPath(candidate) Then
        ResolveRootPath = candidate
        Exit Function
    End If

    Err.Raise vbObjectError + 1, , "単価マスタフォルダを自動検出できませんでした。SharePointをローカル同期し、次のファイルが存在することを確認してください。" & vbCrLf & _
        CombinePath(PROJECT_LIST_DIR, PROJECT_LIST_BOOK) & vbCrLf & _
        CombinePath(PROJECT_MASTER_DIR, ConventionalLineName())
End Function

Private Function GetProjectListPath(ByVal rootPath As String) As String
    GetProjectListPath = CombinePath(CombinePath(rootPath, PROJECT_LIST_DIR), PROJECT_LIST_BOOK)
End Function

Private Function IsValidRootPath(ByVal candidate As String) As Boolean
    If Len(candidate) = 0 Then Exit Function
    If Left$(LCase$(candidate), 4) = "http" Then Exit Function

    candidate = TrimTrailingSlash(candidate)
    IsValidRootPath = FileExists(GetProjectListPath(candidate)) And FolderExists(CombinePath(CombinePath(candidate, PROJECT_MASTER_DIR), ConventionalLineName()))
End Function

Private Function OpenReadOnlyWorkbook(ByVal path As String, Optional ByRef openedByTool As Boolean = True) As Workbook
    Dim book As Workbook
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    openedByTool = False

    Set book = FindOpenWorkbook(path, False)
    If Not book Is Nothing Then
        Set OpenReadOnlyWorkbook = book
        Exit Function
    End If

    On Error Resume Next
    Set book = Workbooks.Open( _
        Filename:=path, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, _
        AddToMru:=False, _
        Notify:=False)
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error GoTo 0

    If savedErrNumber = 0 And Not book Is Nothing Then
        openedByTool = True
        Set OpenReadOnlyWorkbook = book
        Exit Function
    End If

    Set book = FindOpenWorkbook(path, True)
    If Not book Is Nothing Then
        WriteLog "既に開いているブックを参照します: " & path
        Set OpenReadOnlyWorkbook = book
        Exit Function
    End If

    If savedErrNumber = 0 Then
        savedErrNumber = vbObjectError + 23
        savedErrSource = "OpenReadOnlyWorkbook"
        savedErrDescription = "ブックを開けませんでした: " & path
    End If
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Function

Private Function FindOpenWorkbook(ByVal path As String, ByVal allowNameMatch As Boolean) As Workbook
    Dim book As Workbook
    Dim targetFullName As String
    Dim targetName As String

    targetFullName = NormalizeWorkbookPath(path)
    targetName = GetFileName(Replace(path, "/", "\"))

    For Each book In Application.Workbooks
        If NormalizeWorkbookPath(book.FullName) = targetFullName Then
            Set FindOpenWorkbook = book
            Exit Function
        End If
    Next book

    If Not allowNameMatch Then Exit Function
    For Each book In Application.Workbooks
        If StrComp(book.Name, targetName, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = book
            Exit Function
        End If
    Next book
End Function

Private Function NormalizeWorkbookPath(ByVal path As String) As String
    NormalizeWorkbookPath = LCase$(TrimTrailingSlash(Replace(path, "/", "\")))
End Function

Private Function SelectSourceFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "参照フォルダを選択してください"
        .InitialFileName = DEFAULT_SOURCE_ROOT & "\"
        If .Show <> -1 Then
            SelectSourceFolder = vbNullString
        Else
            SelectSourceFolder = .SelectedItems(1)
        End If
    End With
End Function

Private Sub StartLog()
    Dim logDir As String
    logDir = CombinePath(mRootPath, "logs")
    EnsureFolder logDir
    mLogPath = CombinePath(logDir, "unit-price-list-vba_" & Format(Now, "yyyymmdd_hhnnss") & ".log")

    mLogBuffer = vbNullString
    Dim fileNo As Integer
    fileNo = FreeFile
    Open mLogPath For Output As #fileNo
    Print #fileNo, "=== Unit Price List VBA Log ==="
    Print #fileNo, "Started    : " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNo, "User       : " & Environ$("USERNAME")
    Print #fileNo, "Machine    : " & Environ$("COMPUTERNAME")
    Print #fileNo, "Workbook   : " & ThisWorkbook.FullName
    Print #fileNo, ""
    Close #fileNo
End Sub

Private Sub WriteLog(ByVal message As String)
    mLogBuffer = mLogBuffer & Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & message & vbCrLf
    If Len(mLogBuffer) > 16384 Then FlushLog
End Sub

Private Sub FlushLog()
    If Len(mLogBuffer) = 0 Then Exit Sub
    If Len(mLogPath) = 0 Then Exit Sub
    Dim fileNo As Integer
    fileNo = FreeFile
    Open mLogPath For Append As #fileNo
    Print #fileNo, mLogBuffer;
    Close #fileNo
    mLogBuffer = vbNullString
End Sub

Private Sub WriteErrorLog(ByVal context As String, ByVal errObject As ErrObject)
    WriteErrorLogValues context, errObject.Number, errObject.Source, errObject.Description
End Sub

Private Sub WriteErrorLogValues(ByVal context As String, ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String)
    WriteLog "ERROR: " & context
    WriteLog "  Number : " & CStr(errNumber)
    WriteLog "  Source : " & errSource
    WriteLog "  Message: " & errDescription
End Sub

Private Function ResolveBranches(ByVal selectedPath As String) As Collection
    Dim result As Collection
    Dim seen As Object
    Dim outputLineDir As String

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    selectedPath = TrimTrailingSlash(selectedPath)
    If Not FolderExists(selectedPath) Then Err.Raise vbObjectError + 2, , "参照フォルダが存在しません: " & selectedPath

    If Not AddBranchesFromLineRoots(selectedPath, result, seen) Then
        outputLineDir = ResolveOutputLineDir(selectedPath)
        If Len(outputLineDir) = 0 Then outputLineDir = OUTPUT_CONVENTIONAL_LINE_DIR
        If Not AddBranchesFromSourceRoot(selectedPath, result, seen, outputLineDir) Then
            WriteLog "選択フォルダ直下に対象形式なし。配下フォルダを検索します: " & selectedPath
            ScanSourceRoots selectedPath, 0, result, seen, outputLineDir
        End If
    End If

    If result.Count = 0 Then
        Err.Raise vbObjectError + 4, , "01_在来線/02_新幹線、近畿統括本部または金沢支社のフォルダ、保線区フォルダが並ぶフォルダ、またはそれらを配下に含むフォルダを選択してください。"
    End If

    Set ResolveBranches = result
End Function

Private Function AddBranchesFromLineRoots(ByVal sourcePath As String, ByVal result As Collection, ByVal seen As Object) As Boolean
    Dim found As Boolean

    found = AddBranchesFromLineRoot(sourcePath, OUTPUT_CONVENTIONAL_LINE_DIR, result, seen)

    AddBranchesFromLineRoots = found
End Function

Private Function AddBranchesFromLineRoot(ByVal sourcePath As String, ByVal outputLineDir As String, ByVal result As Collection, ByVal seen As Object) As Boolean
    Dim linePath As String
    Dim beforeCount As Long

    linePath = CombinePath(sourcePath, outputLineDir)
    If Not FolderExists(linePath) Then Exit Function

    beforeCount = result.Count
    WriteLog "線区フォルダ起点: " & linePath
    If Not AddBranchesFromSourceRoot(linePath, result, seen, outputLineDir) Then
        ScanSourceRoots linePath, 0, result, seen, outputLineDir
    End If

    AddBranchesFromLineRoot = (result.Count > beforeCount)
End Function

Private Function ResolveOutputLineDir(ByVal sourcePath As String) As String
    If InStr(sourcePath, OUTPUT_SHINKANSEN_LINE_DIR) > 0 Or InStr(sourcePath, "新幹線") > 0 Then
        ResolveOutputLineDir = OUTPUT_SHINKANSEN_LINE_DIR
    ElseIf InStr(sourcePath, OUTPUT_CONVENTIONAL_LINE_DIR) > 0 Or InStr(sourcePath, "在来線") > 0 Then
        ResolveOutputLineDir = OUTPUT_CONVENTIONAL_LINE_DIR
    End If
End Function

Private Function AddBranchesFromSourceRoot(ByVal sourcePath As String, ByVal result As Collection, ByVal seen As Object, ByVal outputLineDir As String) As Boolean
    Dim leaf As String
    Dim kinkiPath As String
    Dim kanazawaPath As String
    Dim foundBranchFolder As Boolean
    Dim detectedLineDir As String

    sourcePath = TrimTrailingSlash(sourcePath)
    detectedLineDir = ResolveOutputLineDir(sourcePath)
    If Len(detectedLineDir) > 0 Then outputLineDir = detectedLineDir
    If Len(outputLineDir) = 0 Then outputLineDir = OUTPUT_CONVENTIONAL_LINE_DIR

    leaf = GetFileName(sourcePath)
    If leaf = "近畿統括本部" Or leaf = "金沢支社" Then
        AddBranchInfo result, seen, leaf, sourcePath, outputLineDir
        WriteLog "対象フォルダ検出: " & outputLineDir & " / " & sourcePath
        AddBranchesFromSourceRoot = True
        Exit Function
    End If

    kinkiPath = CombinePath(sourcePath, "近畿統括本部")
    kanazawaPath = CombinePath(sourcePath, "金沢支社")

    If FolderExists(kinkiPath) Then
        AddBranchInfo result, seen, "近畿統括本部", kinkiPath, outputLineDir
        foundBranchFolder = True
    End If
    If FolderExists(kanazawaPath) Then
        AddBranchInfo result, seen, "金沢支社", kanazawaPath, outputLineDir
        foundBranchFolder = True
    End If
    If foundBranchFolder Then
        WriteLog "対象フォルダ検出: " & outputLineDir & " / " & sourcePath
        AddBranchesFromSourceRoot = True
        Exit Function
    End If

    If HasAreaFolders(sourcePath) Then
        AddBranchInfo result, seen, "近畿統括本部", sourcePath, outputLineDir
        AddBranchInfo result, seen, "金沢支社", sourcePath, outputLineDir
        WriteLog "本部支社フォルダなし。保線区フォルダ直下形式として処理します: " & outputLineDir & " / " & sourcePath
        AddBranchesFromSourceRoot = True
    End If
End Function

Private Sub AddBranchInfo(ByVal result As Collection, ByVal seen As Object, ByVal branchName As String, ByVal branchPath As String, ByVal outputLineDir As String)
    Dim key As String
    Dim detectedLineDir As String

    branchPath = TrimTrailingSlash(branchPath)
    detectedLineDir = ResolveOutputLineDir(branchPath)
    If Len(detectedLineDir) > 0 Then outputLineDir = detectedLineDir
    If Len(outputLineDir) = 0 Then outputLineDir = OUTPUT_CONVENTIONAL_LINE_DIR
    key = outputLineDir & "|" & branchName & "|" & LCase$(branchPath)
    If seen.Exists(key) Then Exit Sub

    seen.Add key, True
    result.Add Array(branchName, branchPath, outputLineDir)
End Sub

Private Function FilterBranchesByLine(ByVal branches As Collection, ByVal outputLineDir As String) As Collection
    Dim result As Collection
    Dim branchInfo As Variant

    Set result = New Collection
    If branches Is Nothing Then
        Set FilterBranchesByLine = result
        Exit Function
    End If

    For Each branchInfo In branches
        If CStr(branchInfo(2)) = outputLineDir Then result.Add branchInfo
    Next branchInfo

    Set FilterBranchesByLine = result
End Function

Private Function ConventionalLineName() As String
    Dim pos As Long

    pos = InStr(OUTPUT_CONVENTIONAL_LINE_DIR, "_")
    If pos > 0 Then
        ConventionalLineName = Mid$(OUTPUT_CONVENTIONAL_LINE_DIR, pos + 1)
    Else
        ConventionalLineName = OUTPUT_CONVENTIONAL_LINE_DIR
    End If
End Function

Private Function GetConventionalMasterDir() As String
    GetConventionalMasterDir = CombinePath(CombinePath(mRootPath, PROJECT_MASTER_DIR), ConventionalLineName())
End Function

Private Function ShinkansenLineName() As String
    Dim pos As Long

    pos = InStr(OUTPUT_SHINKANSEN_LINE_DIR, "_")
    If pos > 0 Then
        ShinkansenLineName = Mid$(OUTPUT_SHINKANSEN_LINE_DIR, pos + 1)
    Else
        ShinkansenLineName = OUTPUT_SHINKANSEN_LINE_DIR
    End If
End Function

Private Function ShinkansenProjectListSheetName() As String
    ShinkansenProjectListSheetName = ShinkansenLineName() & "_" & PROJECT_LIST_SHEET
End Function

Private Function GetShinkansenMasterDir() As String
    GetShinkansenMasterDir = CombinePath(CombinePath(mRootPath, PROJECT_MASTER_DIR), ShinkansenLineName())
End Function

Private Function ResolveShinkansenSourceRoots(ByVal selectedPath As String) As Collection
    Dim result As Collection
    Dim seen As Object
    Dim linePath As String

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    selectedPath = TrimTrailingSlash(selectedPath)
    If Not FolderExists(selectedPath) Then
        Set ResolveShinkansenSourceRoots = result
        Exit Function
    End If

    linePath = CombinePath(selectedPath, OUTPUT_SHINKANSEN_LINE_DIR)
    If FolderExists(linePath) Then
        AddUniquePath result, seen, linePath
        Set ResolveShinkansenSourceRoots = result
        Exit Function
    End If

    If GetFileName(selectedPath) = OUTPUT_SHINKANSEN_LINE_DIR Or ResolveOutputLineDir(selectedPath) = OUTPUT_SHINKANSEN_LINE_DIR Then
        AddUniquePath result, seen, selectedPath
        Set ResolveShinkansenSourceRoots = result
        Exit Function
    End If

    ScanShinkansenLineRoots selectedPath, 0, result, seen
    Set ResolveShinkansenSourceRoots = result
End Function

Private Sub ScanShinkansenLineRoots(ByVal parentPath As String, ByVal depth As Long, ByVal result As Collection, ByVal seen As Object)
    Dim childPath As Variant

    If depth > SOURCE_SEARCH_MAX_DEPTH Then Exit Sub
    If GetFileName(parentPath) = OUTPUT_SHINKANSEN_LINE_DIR Then
        AddUniquePath result, seen, parentPath
        Exit Sub
    End If

    For Each childPath In ListChildFolderPaths(parentPath)
        ScanShinkansenLineRoots CStr(childPath), depth + 1, result, seen
    Next childPath
End Sub

Private Sub AddUniquePath(ByVal result As Collection, ByVal seen As Object, ByVal path As String)
    Dim key As String

    path = TrimTrailingSlash(path)
    key = LCase$(path)
    If seen.Exists(key) Then Exit Sub
    seen.Add key, True
    result.Add path
End Sub
Private Sub ScanSourceRoots(ByVal parentPath As String, ByVal depth As Long, ByVal result As Collection, ByVal seen As Object, ByVal outputLineDir As String)
    Dim childFolders As Collection
    Dim childPath As Variant

    If depth > SOURCE_SEARCH_MAX_DEPTH Then Exit Sub
    If AddBranchesFromSourceRoot(parentPath, result, seen, outputLineDir) Then Exit Sub

    Set childFolders = ListChildFolderPaths(parentPath)
    For Each childPath In childFolders
        ScanSourceRoots CStr(childPath), depth + 1, result, seen, outputLineDir
    Next childPath
End Sub
Private Function ListChildFolderPaths(ByVal parentPath As String) As Collection
    Dim result As Collection
    Dim fso As Object
    Dim folder As Object
    Dim childFolder As Object

    Set result = New Collection
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(parentPath)
    For Each childFolder In folder.SubFolders
        result.Add CStr(childFolder.Path)
    Next childFolder

Done:
    Set ListChildFolderPaths = result
End Function

Private Function HasAreaFolders(ByVal parentPath As String) As Boolean
    Dim childPath As Variant
    Dim name As String

    For Each childPath In ListChildFolderPaths(parentPath)
        name = GetFileName(CStr(childPath))
        If Left$(name, 4) Like "####" And HasProjectFolders(CStr(childPath)) Then
            HasAreaFolders = True
            Exit Function
        End If
    Next childPath
End Function

Private Function HasProjectFolders(ByVal areaPath As String) As Boolean
    Dim childPath As Variant
    Dim name As String

    For Each childPath In ListChildFolderPaths(areaPath)
        name = GetFileName(CStr(childPath))
        If Len(GetProjectCode(name)) > 0 Then
            HasProjectFolders = True
            Exit Function
        End If
    Next childPath
End Function

Private Function GetRowProjectCode(ByVal sheet As Worksheet, ByVal rowIndex As Long, ByVal fallbackCode As String) As String
    GetRowProjectCode = NormalizeProjectCode(sheet.Cells(rowIndex, 2).Value)
    If Len(GetRowProjectCode) = 0 Then GetRowProjectCode = fallbackCode
End Function

Private Function NormalizeProjectCode(ByVal value As Variant) As String
    Dim text As String
    Dim digits As String

    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function
    If IsNumeric(text) Then
        NormalizeProjectCode = CStr(CLng(CDbl(text)))
        Exit Function
    End If

    digits = OnlyDigits(text)
    If Len(digits) > 0 Then
        NormalizeProjectCode = CStr(CLng(CDbl(digits)))
    Else
        NormalizeProjectCode = GetProjectCode(text)
    End If
End Function

Private Function FindProjectFolder(ByVal areaDir As String, ByVal projectName As String, ByVal projectCode As String) As String
    Dim name As String
    Dim fullPath As String

    fullPath = CombinePath(areaDir, projectName)
    If FolderExists(fullPath) Then
        FindProjectFolder = fullPath
        Exit Function
    End If

    name = Dir$(CombinePath(areaDir, "*"), vbDirectory)
    Do While Len(name) > 0
        If name <> "." And name <> ".." Then
            fullPath = CombinePath(areaDir, name)
            If FolderExists(fullPath) Then
                If GetProjectCode(name) = projectCode Then
                    FindProjectFolder = fullPath
                    Exit Function
                End If
            End If
        End If
        name = Dir$()
    Loop
End Function

Private Function ResolveConventionalMasterPath(ByVal projectName As String) As String
    Dim primaryPath As String
    Dim legacyPath As String

    primaryPath = CombinePath(GetConventionalMasterDir(), projectName & ".xlsx")
    If FileExists(primaryPath) Then
        ResolveConventionalMasterPath = primaryPath
        Exit Function
    End If

    legacyPath = CombinePath(CombinePath(mRootPath, PROJECT_MASTER_DIR), projectName & ".xlsx")
    If FileExists(legacyPath) Then
        ResolveConventionalMasterPath = legacyPath
    Else
        ResolveConventionalMasterPath = primaryPath
    End If
End Function

Private Sub ProcessProjectAllBranches(ByVal projectName As String, ByVal branches As Collection, ByVal targetBooks As Object)
    On Error GoTo ProjectError

    Dim projectCode As String
    Dim masterPath As String
    Dim masterBook As Workbook
    Dim masterSheet As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim branchInfo As Variant
    Dim branchName As String
    Dim branchPath As String
    Dim outputLineDir As String
    Dim areaCache As Object
    Dim projectFolderCache As Object
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim masterOpenedByTool As Boolean

    projectCode = GetProjectCode(projectName)
    If Len(projectCode) = 0 Then
        WriteLog "スキップ: 工事件名コードを取得できません: " & projectName
        Exit Sub
    End If

    masterPath = ResolveConventionalMasterPath(projectName)
    If Not FileExists(masterPath) Then
        WriteLog "スキップ: 工事件名別マスタが見つかりません: " & masterPath
        Exit Sub
    End If

    WriteLog "工事件名別マスタ: " & masterPath & " / コード " & projectCode
    Err.Clear
    On Error Resume Next
    Set masterBook = OpenReadOnlyWorkbook(masterPath, masterOpenedByTool)
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error GoTo ProjectError
    If savedErrNumber <> 0 Or masterBook Is Nothing Then
        WriteErrorLogValues "工事件名別マスタを開けないためスキップ: " & projectName, savedErrNumber, savedErrSource, savedErrDescription
        Exit Sub
    End If
    Set areaCache = CreateObject("Scripting.Dictionary")
    Set projectFolderCache = CreateObject("Scripting.Dictionary")

    For Each branchInfo In branches
        branchName = CStr(branchInfo(0))
        branchPath = CStr(branchInfo(1))
        outputLineDir = CStr(branchInfo(2))
        If Not SheetExists(masterBook, branchName) Then
            WriteLog "スキップ: マスタに支社シートがありません: " & projectName & " / " & branchName
        Else
            Set masterSheet = masterBook.Worksheets(branchName)
            lastRow = masterSheet.Cells(masterSheet.Rows.Count, 3).End(xlUp).Row
            For rowIndex = 2 To lastRow
                ProcessMasterRow masterSheet, rowIndex, projectName, projectCode, branchName, branchPath, outputLineDir, targetBooks, areaCache, projectFolderCache
            Next rowIndex
        End If
    Next branchInfo

    If masterOpenedByTool Then masterBook.Close SaveChanges:=False
    Exit Sub

ProjectError:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    WriteErrorLogValues "工事件名処理: " & projectName, savedErrNumber, savedErrSource, savedErrDescription
    On Error Resume Next
    If masterOpenedByTool Then
        If Not masterBook Is Nothing Then masterBook.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Function ReadProjectNames(ByVal sheetName As String) As Collection
    On Error GoTo ReadError

    Dim result As Collection
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim value As String
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim bookOpenedByTool As Boolean

    Set result = New Collection
    Set book = OpenReadOnlyWorkbook(GetProjectListPath(mRootPath), bookOpenedByTool)
    Set sheet = book.Worksheets(sheetName)
    lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row

    For rowIndex = 2 To lastRow
        value = Trim$(CStr(sheet.Cells(rowIndex, 1).Value))
        If Len(value) > 0 Then AddUnique result, RemoveExtension(value)
    Next rowIndex

    If bookOpenedByTool Then book.Close SaveChanges:=False
    Set ReadProjectNames = result
    Exit Function

ReadError:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error Resume Next
    If bookOpenedByTool Then
        If Not book Is Nothing Then book.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Function

Private Sub ProcessShinkansenProject(ByVal projectName As String, ByVal sourceRoots As Collection, ByVal targetBooks As Object)
    On Error GoTo ProjectError

    Dim masterPath As String
    Dim masterBook As Workbook
    Dim masterOpenedByTool As Boolean
    Dim masterBaseName As String
    Dim sourceRoot As Variant
    Dim projectDirs As Collection
    Dim projectDir As Variant
    Dim processedDirs As Object
    Dim dirKey As String
    Dim fileName As String
    Dim matchedFileCount As Long
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    If sourceRoots Is Nothing Then Exit Sub
    If sourceRoots.Count = 0 Then Exit Sub

    masterPath = ResolveShinkansenMasterPath(projectName)
    If Len(masterPath) = 0 Then
        WriteLog "スキップ: 新幹線工事件名別マスタが見つかりません: " & projectName
        Exit Sub
    End If
    masterBaseName = RemoveExtension(GetFileName(masterPath))

    WriteLog "新幹線工事件名別マスタ: " & masterPath
    Err.Clear
    On Error Resume Next
    Set masterBook = OpenReadOnlyWorkbook(masterPath, masterOpenedByTool)
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    On Error GoTo ProjectError
    If savedErrNumber <> 0 Or masterBook Is Nothing Then
        WriteErrorLogValues "新幹線工事件名別マスタを開けないためスキップ: " & projectName, savedErrNumber, savedErrSource, savedErrDescription
        Exit Sub
    End If

    Set processedDirs = CreateObject("Scripting.Dictionary")
    For Each sourceRoot In sourceRoots
        Set projectDirs = FindShinkansenProjectFolders(CStr(sourceRoot), projectName, masterBaseName)
        If projectDirs.Count = 0 Then
            WriteLog "未検出: 新幹線フォルダ " & projectName & " / " & CStr(sourceRoot)
        Else
            For Each projectDir In projectDirs
                dirKey = LCase$(TrimTrailingSlash(CStr(projectDir)))
                If Not processedDirs.Exists(dirKey) Then
                    processedDirs.Add dirKey, True
                    matchedFileCount = 0
                    fileName = Dir$(CombinePath(CStr(projectDir), "*.xls*"))
                    Do While Len(fileName) > 0
                        If ShouldUseShinkansenSourceFile(fileName) Then
                            matchedFileCount = matchedFileCount + 1
                            CopyShinkansenSourceToTarget CombinePath(CStr(projectDir), fileName), fileName, projectName, masterBook, targetBooks
                        End If
                        fileName = Dir$()
                    Loop
                    If matchedFileCount = 0 Then WriteLog "対象ファイルなし: " & CStr(projectDir)
                End If
            Next projectDir
        End If
    Next sourceRoot

    If masterOpenedByTool Then masterBook.Close SaveChanges:=False
    Exit Sub

ProjectError:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    WriteErrorLogValues "新幹線工事件名処理: " & projectName, savedErrNumber, savedErrSource, savedErrDescription
    On Error Resume Next
    If masterOpenedByTool Then
        If Not masterBook Is Nothing Then masterBook.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Function ResolveShinkansenMasterPath(ByVal projectName As String) As String
    Dim masterDir As String
    Dim candidatePath As String
    Dim fileName As String

    masterDir = GetShinkansenMasterDir()
    candidatePath = CombinePath(masterDir, projectName & ".xlsx")
    If FileExists(candidatePath) Then
        ResolveShinkansenMasterPath = candidatePath
        Exit Function
    End If

    fileName = Dir$(CombinePath(masterDir, "*.xlsx"))
    Do While Len(fileName) > 0
        candidatePath = CombinePath(masterDir, fileName)
        If ShinkansenMasterContainsProjectName(candidatePath, projectName) Then
            ResolveShinkansenMasterPath = candidatePath
            Exit Function
        End If
        fileName = Dir$()
    Loop
End Function

Private Function ShinkansenMasterContainsProjectName(ByVal masterPath As String, ByVal projectName As String) As Boolean
    On Error GoTo Done

    Dim book As Workbook
    Dim openedByTool As Boolean
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim lastRow As Long

    Set book = OpenReadOnlyWorkbook(masterPath, openedByTool)
    For Each ws In book.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
        For rowIndex = 2 To lastRow
            If TextMatches(CStr(ws.Cells(rowIndex, 3).Value), projectName) Then
                ShinkansenMasterContainsProjectName = True
                GoTo Done
            End If
        Next rowIndex
    Next ws

Done:
    On Error Resume Next
    If openedByTool Then
        If Not book Is Nothing Then book.Close SaveChanges:=False
    End If
    On Error GoTo 0
End Function

Private Function FindShinkansenProjectFolders(ByVal rootPath As String, ByVal projectName As String, ByVal alternateName As String) As Collection
    Dim result As Collection
    Dim names As Collection
    Dim seen As Object

    Set result = New Collection
    Set names = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    AddUnique names, projectName
    If Len(alternateName) > 0 Then AddUnique names, alternateName

    ScanShinkansenProjectFolders TrimTrailingSlash(rootPath), names, 0, result, seen
    Set FindShinkansenProjectFolders = result
End Function

Private Sub ScanShinkansenProjectFolders(ByVal parentPath As String, ByVal names As Collection, ByVal depth As Long, ByVal result As Collection, ByVal seen As Object)
    Dim childPath As Variant

    If depth > SOURCE_SEARCH_MAX_DEPTH Then Exit Sub
    If FolderNameMatches(GetFileName(parentPath), names) Then
        AddUniquePath result, seen, parentPath
        Exit Sub
    End If

    For Each childPath In ListChildFolderPaths(parentPath)
        ScanShinkansenProjectFolders CStr(childPath), names, depth + 1, result, seen
    Next childPath
End Sub

Private Function FolderNameMatches(ByVal folderName As String, ByVal names As Collection) As Boolean
    Dim item As Variant

    For Each item In names
        If folderName = CStr(item) Then
            FolderNameMatches = True
            Exit Function
        End If
    Next item
End Function

Private Function ShouldUseShinkansenSourceFile(ByVal fileName As String) As Boolean
    Dim ext As String

    If Left$(fileName, 2) = "~$" Then Exit Function
    If InStr(fileName, "ｼｽﾃﾑ") > 0 Or InStr(fileName, "システム") > 0 Then Exit Function
    ext = LCase$(Mid$(fileName, InStrRev(fileName, ".") + 1))
    ShouldUseShinkansenSourceFile = (ext = "xlsx" Or ext = "xlsm" Or ext = "xls")
End Function

Private Sub CopyShinkansenSourceToTarget(ByVal sourcePath As String, ByVal sourceFileName As String, ByVal projectName As String, ByVal masterBook As Workbook, ByVal targetBooks As Object)
    On Error GoTo CopyError

    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetBook As Workbook
    Dim targetPath As String
    Dim yearText As String
    Dim sheetName As String
    Dim isChange As Boolean
    Dim isPurchaseProject As Boolean
    Dim branchName As String
    Dim areaName As String
    Dim projectTitle As String
    Dim lineName As String
    Dim targetProjectName As String
    Dim copiedSheet As Worksheet
    Dim createdTargetBook As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim sourceOpenedByTool As Boolean

    yearText = GetYearFromFileName(RemoveExtension(sourceFileName))
    If Len(yearText) = 0 Then yearText = GetYearFromFileName(sourcePath)
    If Len(yearText) = 0 Then yearText = mSourceYearText
    If Len(yearText) = 0 Then
        WriteLog "スキップ: 年度を取得できません: " & sourcePath
        Exit Sub
    End If

    isChange = (InStr(sourceFileName, "(変)") > 0 Or InStr(sourcePath, "(変)") > 0 Or _
                InStr(sourceFileName, PURCHASE_TITLE_PREFIX_CHANGE) > 0 Or InStr(sourcePath, PURCHASE_TITLE_PREFIX_CHANGE) > 0)

    Set sourceBook = OpenReadOnlyWorkbook(sourcePath, sourceOpenedByTool)
    Set sourceSheet = sourceBook.Worksheets(1)
    isPurchaseProject = IsShinkansenPurchaseSource(sourceSheet, projectName)

    If Not ResolveShinkansenMasterRow(masterBook, sourceSheet, projectName, isPurchaseProject, branchName, areaName, projectTitle, lineName) Then
        WriteLog "スキップ: 新幹線マスタ行に一致しません: " & sourceFileName
        If sourceOpenedByTool Then sourceBook.Close SaveChanges:=False
        Set sourceBook = Nothing
        Exit Sub
    End If

    targetProjectName = BuildShinkansenProjectNamePart(projectTitle, lineName)
    targetPath = BuildTargetPath(OUTPUT_SHINKANSEN_LINE_DIR, branchName, areaName, yearText, isChange, targetProjectName)

    If FileExists(targetPath) And Not targetBooks.Exists(targetPath) Then
        WriteLog "スキップ: 出力ファイルが既に存在します: " & targetPath
        If sourceOpenedByTool Then sourceBook.Close SaveChanges:=False
        Set sourceBook = Nothing
        Exit Sub
    End If

    sheetName = RemoveExtension(sourceFileName)
    If Len(sheetName) = 0 Then sheetName = targetProjectName

    If Not targetBooks.Exists(targetPath) Then
        Set targetBook = Workbooks.Add(xlWBATWorksheet)
        createdTargetBook = True
        sourceSheet.Copy After:=targetBook.Worksheets(1)
        Set copiedSheet = targetBook.Worksheets(targetBook.Worksheets.Count)
        copiedSheet.Name = SafeSheetName(sheetName)
        If isPurchaseProject Then FormatShinkansenPurchaseUnitPriceSheet copiedSheet, sourceFileName, isChange, areaName, yearText
        Application.DisplayAlerts = False
        targetBook.Worksheets(1).Delete
        targetBooks.Add targetPath, targetBook
    Else
        Set targetBook = targetBooks(targetPath)
        sourceSheet.Copy After:=targetBook.Worksheets(targetBook.Worksheets.Count)
        Set copiedSheet = targetBook.Worksheets(targetBook.Worksheets.Count)
        copiedSheet.Name = UniqueSheetName(targetBook, sheetName)
        If isPurchaseProject Then FormatShinkansenPurchaseUnitPriceSheet copiedSheet, sourceFileName, isChange, areaName, yearText
    End If

    WriteLog "  新幹線追加: " & sourceFileName & " -> " & targetPath
    FlushLog
    If sourceOpenedByTool Then sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing
    frmProgress.UpdateProgress mProgDone, mProgName & " / " & sourceFileName
    Exit Sub

CopyError:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    WriteErrorLogValues "新幹線単価一覧コピー: " & sourcePath, savedErrNumber, savedErrSource, savedErrDescription
    On Error Resume Next
    If Not copiedSheet Is Nothing Then
        Application.DisplayAlerts = False
        copiedSheet.Delete
    End If
    If createdTargetBook Then
        If Not targetBooks Is Nothing Then
            If targetBooks.Exists(targetPath) Then targetBooks.Remove targetPath
        End If
        If Not targetBook Is Nothing Then targetBook.Close SaveChanges:=False
    End If
    If sourceOpenedByTool Then
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Function ResolveShinkansenMasterRow(ByVal masterBook As Workbook, ByVal sourceSheet As Worksheet, ByVal projectName As String, ByVal isPurchaseProject As Boolean, ByRef branchName As String, ByRef areaName As String, ByRef projectTitle As String, ByRef lineName As String) As Boolean
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim sourceProjectTitle As String
    Dim sourceLineName As String
    Dim purchaseKey As String
    Dim rowProjectTitle As String
    Dim rowLineName As String
    Dim rowShortName As String

    If isPurchaseProject Then
        purchaseKey = GetShinkansenPurchaseOfficeKey(sourceSheet)
        If Len(purchaseKey) = 0 Then Exit Function
    ElseIf IsShinkansenInspectionProject(projectName) Then
        sourceProjectTitle = GetShinkansenInspectionProjectTitle(sourceSheet)
        sourceLineName = vbNullString
        If Len(sourceProjectTitle) = 0 Then Exit Function
    Else
        sourceProjectTitle = Trim$(CStr(sourceSheet.Range("B3").Value))
        sourceLineName = Trim$(CStr(sourceSheet.Range("B4").Value))
        If Len(sourceProjectTitle) = 0 Then Exit Function
    End If

    For Each ws In masterBook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For rowIndex = 2 To lastRow
            rowProjectTitle = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
            rowLineName = Trim$(CStr(ws.Cells(rowIndex, 5).Value))
            rowShortName = Trim$(CStr(ws.Cells(rowIndex, 4).Value))
            If isPurchaseProject Then
                If TextMatches(rowShortName, purchaseKey) Then
                    ResolveShinkansenMasterRow = ApplyShinkansenMasterRow(ws, rowIndex, branchName, areaName, projectTitle, lineName)
                    Exit Function
                End If
            ElseIf TextMatches(rowProjectTitle, sourceProjectTitle) Then
                If Len(sourceLineName) = 0 Or Len(rowLineName) = 0 Or TextMatches(rowLineName, sourceLineName) Then
                    ResolveShinkansenMasterRow = ApplyShinkansenMasterRow(ws, rowIndex, branchName, areaName, projectTitle, lineName)
                    Exit Function
                End If
            End If
        Next rowIndex
    Next ws
End Function

Private Function ApplyShinkansenMasterRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByRef branchName As String, ByRef areaName As String, ByRef projectTitle As String, ByRef lineName As String) As Boolean
    branchName = ws.Name
    areaName = Trim$(CStr(ws.Cells(rowIndex, 1).Value))
    projectTitle = Trim$(CStr(ws.Cells(rowIndex, 3).Value))
    lineName = Trim$(CStr(ws.Cells(rowIndex, 5).Value))
    ApplyShinkansenMasterRow = (Len(branchName) > 0 And Len(areaName) > 0 And Len(projectTitle) > 0)
End Function

Private Function IsShinkansenPurchaseSource(ByVal sourceSheet As Worksheet, ByVal projectName As String) As Boolean
    If InStr(projectName, PurchaseProjectKeyword()) > 0 Then
        IsShinkansenPurchaseSource = True
    Else
        IsShinkansenPurchaseSource = (InStr(CStr(sourceSheet.Range("A1").Value), PurchaseProjectKeyword()) > 0)
    End If
End Function

Private Function PurchaseProjectKeyword() As String
    PurchaseProjectKeyword = "購入充当"
End Function

Private Function IsShinkansenInspectionProject(ByVal projectName As String) As Boolean
    IsShinkansenInspectionProject = (InStr(projectName, "検査") > 0)
End Function

Private Function GetShinkansenInspectionProjectTitle(ByVal sourceSheet As Worksheet) As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim lastCol As Long
    Dim value As String
    Dim rowText As String
    Dim extractedTitle As String

    For rowIndex = 1 To 12
        lastCol = sourceSheet.Cells(rowIndex, sourceSheet.Columns.Count).End(xlToLeft).Column
        rowText = vbNullString
        For colIndex = 1 To lastCol
            value = Trim$(CStr(sourceSheet.Cells(rowIndex, colIndex).Value))
            If Len(value) > 0 Then rowText = rowText & value
        Next colIndex

        extractedTitle = ExtractShinkansenInspectionProjectTitle(rowText)
        If Len(extractedTitle) > 0 Then
            GetShinkansenInspectionProjectTitle = extractedTitle
            Exit Function
        End If
    Next rowIndex
End Function

Private Function ExtractShinkansenInspectionProjectTitle(ByVal text As String) As String
    Dim normalizedText As String
    Dim pos As Long
    Dim result As String

    normalizedText = NormalizeCompareText(text)
    pos = InStr(normalizedText, "件名")
    If pos = 0 Then Exit Function

    result = Mid$(normalizedText, pos + Len("件名"))
    Do While Len(result) > 0 And (Left$(result, 1) = ":" Or Left$(result, 1) = "：" Or Left$(result, 1) = "・")
        result = Mid$(result, 2)
    Loop

    ExtractShinkansenInspectionProjectTitle = result
End Function

Private Function GetShinkansenPurchaseOfficeKey(ByVal sourceSheet As Worksheet) As String
    Dim address As Variant
    Dim value As String

    For Each address In Array("H2", "H3", "H1", "H4")
        value = Trim$(CStr(sourceSheet.Range(CStr(address)).Value))
        If Len(value) > 0 And Not IsNumeric(value) And InStr(value, "単価") = 0 And InStr(value, "円") = 0 Then
            GetShinkansenPurchaseOfficeKey = value
            Exit Function
        End If
    Next address
End Function

Private Function BuildShinkansenProjectNamePart(ByVal projectTitle As String, ByVal lineName As String) As String
    BuildShinkansenProjectNamePart = projectTitle
    If Len(lineName) > 0 Then BuildShinkansenProjectNamePart = BuildShinkansenProjectNamePart & "_" & lineName
End Function

Private Function TextMatches(ByVal leftText As String, ByVal rightText As String) As Boolean
    leftText = NormalizeCompareText(leftText)
    rightText = NormalizeCompareText(rightText)
    If Len(leftText) = 0 Or Len(rightText) = 0 Then Exit Function
    TextMatches = (leftText = rightText Or InStr(leftText, rightText) > 0 Or InStr(rightText, leftText) > 0)
End Function

Private Function NormalizeCompareText(ByVal value As String) As String
    value = StrConv(Trim$(CStr(value)), vbNarrow)
    value = Replace$(value, vbCr, "")
    value = Replace$(value, vbLf, "")
    value = Replace$(value, vbTab, "")
    value = Replace$(value, " ", "")
    value = Replace$(value, ChrW$(&H3000), "")
    NormalizeCompareText = value
End Function

Private Sub FormatShinkansenPurchaseUnitPriceSheet(ByVal ws As Worksheet, ByVal sourceFileName As String, ByVal isChange As Boolean, ByVal areaText As String, Optional ByVal sourceYearText As String = "")
    Const SOURCE_HEADER_ROW As Long = 3
    Const SOURCE_FIRST_ROW As Long = 4
    Const TABLE_HEADER_ROW As Long = 1

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim a4 As String
    Dim c2 As String
    Dim e5 As String
    Dim code11 As String
    Dim newLastCol As Long
    Dim newLastRow As Long
    Dim tableRange As Range
    Dim tbl As ListObject
    Dim lo As ListObject
    Dim tableName As String
    Dim titleRange As Range
    Dim titleText As String
    Dim yearText As String
    Dim prefixText As String

    tableName = CleanListObjectName(areaText)
    If Len(tableName) = 0 Then tableName = "品目コードTable"

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < SOURCE_FIRST_ROW Then Err.Raise vbObjectError + 31, , "新幹線購入充当単価表のA列4行目以降にデータがありません: " & sourceFileName

    ws.Range("A:A").NumberFormat = "@"
    For rowIndex = SOURCE_FIRST_ROW To lastRow
        If IsNumeric(ws.Cells(rowIndex, "A").Value) And Trim$(CStr(ws.Cells(rowIndex, "A").Value)) <> "" Then
            a4 = Format$(CLng(CDbl(ws.Cells(rowIndex, "A").Value)) + 1000, "0000")
        Else
            a4 = "0000"
        End If
        c2 = PadZeroText(OnlyDigitsNarrow(CStr(ws.Cells(rowIndex, "C").Value)), 2)
        e5 = PadZeroText(OnlyDigitsNarrow(CStr(ws.Cells(rowIndex, "E").Value)), 5)
        code11 = a4 & c2 & e5
        ws.Cells(rowIndex, "A").Value = code11
    Next rowIndex

    ws.Range("A:A").NumberFormat = "0"
    For rowIndex = SOURCE_FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(rowIndex, "A").Value))) > 0 Then ws.Cells(rowIndex, "A").Value = CDec(ws.Cells(rowIndex, "A").Value)
    Next rowIndex

    ws.Columns("J").Delete Shift:=xlShiftToLeft
    ws.Columns("I").Delete Shift:=xlShiftToLeft
    ws.Columns("E").Delete Shift:=xlShiftToLeft
    ws.Columns("C").Delete Shift:=xlShiftToLeft

    ws.Cells(SOURCE_HEADER_ROW, "G").Value = "保線区名"
    For rowIndex = SOURCE_FIRST_ROW To lastRow
        ws.Cells(rowIndex, "G").Value = areaText
    Next rowIndex

    ws.Rows("1:2").UnMerge
    ws.Rows("1:2").Delete Shift:=xlShiftUp

    ws.Cells.VerticalAlignment = xlVAlignCenter

    newLastCol = ws.Cells(TABLE_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    newLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set tableRange = ws.Range(ws.Cells(TABLE_HEADER_ROW, 1), ws.Cells(newLastRow, newLastCol))

    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo

    Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    tbl.Name = UniqueListObjectName(ws.Parent, tableName)
    tbl.TableStyle = "TableStyleMedium7"

    ws.Rows("1:4").Insert Shift:=xlDown
    ws.Rows(5).HorizontalAlignment = xlHAlignCenter
    ws.Columns("A").HorizontalAlignment = xlHAlignCenter
    ws.Columns("E").HorizontalAlignment = xlHAlignCenter
    ws.Columns("G").HorizontalAlignment = xlHAlignCenter
    ws.Columns("F").NumberFormat = "#,##0"
    ws.Cells(5, "E").Value = "単位"
    ws.Cells(5, "F").Value = "単価"

    yearText = sourceYearText
    If Len(yearText) = 0 Then yearText = GetYearFromFileName(RemoveExtension(sourceFileName))
    If Len(yearText) = 0 Then
        yearText = OnlyDigitsNarrow(sourceFileName)
        If Len(yearText) >= 4 Then yearText = Left$(yearText, 4)
    End If
    If isChange Then
        prefixText = PURCHASE_TITLE_PREFIX_CHANGE
    Else
        prefixText = "通常"
    End If
    titleText = yearText & "年度_購入充当単価表_" & prefixText & "_" & areaText

    Set titleRange = ws.Range("A1:G1")
    With titleRange
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Font.Name = "BIZ UDGothic"
        .Font.Size = 14
    End With

    ws.Cells.Font.Name = "BIZ UDGothic"
    On Error Resume Next
    ws.Cells.Font.NameFarEast = "BIZ UDGothic"
    On Error GoTo 0
    ws.Range("A1").Font.Size = 14
    ws.Cells.EntireColumn.AutoFit
End Sub
Private Sub ProcessMasterRow(ByVal sheet As Worksheet, ByVal rowIndex As Long, ByVal projectName As String, ByVal projectCode As String, ByVal branchName As String, ByVal branchPath As String, ByVal outputLineDir As String, ByVal targetBooks As Object, ByVal areaCache As Object, ByVal projectFolderCache As Object)
    Dim areaName As String
    Dim areaCode As String
    Dim lineCode As String
    Dim lineName As String
    Dim areaFolderName As String
    Dim areaDir As String
    Dim projectDir As String
    Dim filePrefix As String
    Dim rowProjectCode As String
    Dim fileName As String
    Dim areaKey As String
    Dim projectKey As String
    Dim areaWasCached As Boolean
    Dim projectWasCached As Boolean
    Dim isPurchaseProject As Boolean

    areaName = Trim$(CStr(sheet.Cells(rowIndex, 1).Value))
    areaCode = FormatCode(sheet.Cells(rowIndex, 3).Value, 4)
    lineName = Trim$(CStr(sheet.Cells(rowIndex, 6).Value))
    isPurchaseProject = IsPurchaseProjectName(projectName, projectCode)

    If isPurchaseProject Then
        lineCode = NormalizePurchaseBranchSuffix(sheet.Cells(rowIndex, 4).Value)
        rowProjectCode = projectCode
        If Len(areaCode) = 0 Or Len(rowProjectCode) = 0 Then Exit Sub
    Else
        lineCode = FormatCode(sheet.Cells(rowIndex, 4).Value, 2)
        rowProjectCode = GetRowProjectCode(sheet, rowIndex, projectCode)
        If Len(areaCode) = 0 Or Len(lineCode) = 0 Or Len(rowProjectCode) = 0 Then Exit Sub
    End If

    If Left$(areaName, 4) Like "####" Then
        areaFolderName = areaName
    Else
        areaFolderName = areaCode & areaName
    End If

    If isPurchaseProject Then
        filePrefix = areaCode & lineCode
    Else
        filePrefix = rowProjectCode & areaCode & lineCode
    End If

    areaKey = branchPath & "|" & areaFolderName & "|" & areaCode
    areaWasCached = areaCache.Exists(areaKey)
    If areaWasCached Then
        areaDir = CStr(areaCache(areaKey))
    Else
        areaDir = FindChildFolder(branchPath, areaFolderName, areaCode)
        areaCache.Add areaKey, areaDir
    End If
    If Len(areaDir) = 0 Then
        If Not areaWasCached Then WriteLog "未検出: 保線区フォルダ " & areaFolderName
        Exit Sub
    End If

    projectKey = areaDir & "|" & projectName & "|" & rowProjectCode
    projectWasCached = projectFolderCache.Exists(projectKey)
    If projectWasCached Then
        projectDir = CStr(projectFolderCache(projectKey))
    Else
        projectDir = FindProjectFolder(areaDir, projectName, rowProjectCode)
        projectFolderCache.Add projectKey, projectDir
    End If
    If Len(projectDir) = 0 Then
        If Not projectWasCached Then WriteLog "未検出: " & GetFileName(areaDir) & " 内の " & projectName
        Exit Sub
    End If

    Dim matchedFileCount As Long
    fileName = Dir$(CombinePath(projectDir, filePrefix & "*.xlsx"))
    Do While Len(fileName) > 0
        If (isPurchaseProject And ShouldUsePurchaseSourceFile(fileName, filePrefix)) Or _
           (Not isPurchaseProject And ShouldUseSourceFile(fileName, filePrefix)) Then
            matchedFileCount = matchedFileCount + 1
            CopySourceToTarget CombinePath(projectDir, fileName), fileName, branchName, outputLineDir, areaFolderName, projectName, lineName, filePrefix, targetBooks, isPurchaseProject
        End If
        fileName = Dir$()
    Loop
    If matchedFileCount = 0 Then WriteLog "対象ファイルなし: " & projectDir & " / 期待先頭=" & filePrefix
End Sub

Private Function ShouldUseSourceFile(ByVal fileName As String, ByVal filePrefix As String) As Boolean
    If Left$(fileName, 2) = "~$" Then Exit Function
    If InStr(fileName, "ｼｽﾃﾑ") > 0 Or InStr(fileName, "システム") > 0 Then Exit Function
    ShouldUseSourceFile = (Left$(RemoveExtension(fileName), Len(filePrefix)) = filePrefix)
End Function

Private Function ShouldUsePurchaseSourceFile(ByVal fileName As String, ByVal filePrefix As String) As Boolean
    Dim baseName As String
    Dim nextChar As String

    If Left$(fileName, 2) = "~$" Then Exit Function
    If InStr(fileName, "ｼｽﾃﾑ") > 0 Or InStr(fileName, "システム") > 0 Then Exit Function

    baseName = RemoveExtension(fileName)
    If Left$(baseName, Len(filePrefix)) <> filePrefix Then Exit Function
    If Len(baseName) = Len(filePrefix) Then
        ShouldUsePurchaseSourceFile = True
        Exit Function
    End If

    nextChar = Mid$(baseName, Len(filePrefix) + 1, 1)
    If nextChar = "-" Or nextChar Like "#" Then Exit Function
    ShouldUsePurchaseSourceFile = True
End Function

Private Sub CopySourceToTarget(ByVal sourcePath As String, ByVal sourceFileName As String, ByVal branchName As String, ByVal outputLineDir As String, ByVal areaFolderName As String, ByVal projectName As String, ByVal lineName As String, ByVal filePrefix As String, ByVal targetBooks As Object, Optional ByVal isPurchaseProject As Boolean = False)
    On Error GoTo CopyError

    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetBook As Workbook
    Dim targetPath As String
    Dim yearText As String
    Dim sheetName As String
    Dim isChange As Boolean
    Dim copiedSheet As Worksheet
    Dim createdTargetBook As Boolean
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String
    Dim sourceOpenedByTool As Boolean

    yearText = GetYearFromFileName(RemoveExtension(sourceFileName))
    If Len(yearText) = 0 Then yearText = GetYearFromFileName(sourcePath)
    If Len(yearText) = 0 Then yearText = mSourceYearText
    If Len(yearText) = 0 Then
        WriteLog "スキップ: 年度を取得できません: " & sourcePath
        Exit Sub
    End If

    isChange = (InStr(sourceFileName, "(変)") > 0 Or InStr(sourcePath, "(変)") > 0 Or _
                InStr(sourceFileName, PURCHASE_TITLE_PREFIX_CHANGE) > 0 Or InStr(sourcePath, PURCHASE_TITLE_PREFIX_CHANGE) > 0)
    targetPath = BuildTargetPath(outputLineDir, branchName, areaFolderName, yearText, isChange, projectName)

    If FileExists(targetPath) And Not targetBooks.Exists(targetPath) Then
        WriteLog "スキップ: 出力ファイルが既に存在します: " & targetPath
        Exit Sub
    End If

    Set sourceBook = OpenReadOnlyWorkbook(sourcePath, sourceOpenedByTool)
    Set sourceSheet = sourceBook.Worksheets(1)
    If isPurchaseProject Then
        sheetName = filePrefix
    Else
        sheetName = Trim$(CStr(sourceSheet.Range("B4").Value))
        If Len(sheetName) = 0 Then
            If Len(lineName) > 0 Then
                sheetName = lineName
            Else
                sheetName = filePrefix
            End If
        End If
    End If
    If Not targetBooks.Exists(targetPath) Then
        Set targetBook = Workbooks.Add(xlWBATWorksheet)
        createdTargetBook = True
        sourceSheet.Copy After:=targetBook.Worksheets(1)
        Set copiedSheet = targetBook.Worksheets(targetBook.Worksheets.Count)
        copiedSheet.Name = SafeSheetName(sheetName)
        If isPurchaseProject Then FormatPurchaseUnitPriceSheet copiedSheet, sourceFileName, isChange
        Application.DisplayAlerts = False
        targetBook.Worksheets(1).Delete
        targetBooks.Add targetPath, targetBook
    Else
        Set targetBook = targetBooks(targetPath)
        sourceSheet.Copy After:=targetBook.Worksheets(targetBook.Worksheets.Count)
        Set copiedSheet = targetBook.Worksheets(targetBook.Worksheets.Count)
        copiedSheet.Name = UniqueSheetName(targetBook, sheetName)
        If isPurchaseProject Then FormatPurchaseUnitPriceSheet copiedSheet, sourceFileName, isChange
    End If

    WriteLog "  追加: " & sourceFileName & " -> " & targetPath
    FlushLog
    If sourceOpenedByTool Then sourceBook.Close SaveChanges:=False
    Set sourceBook = Nothing
    frmProgress.UpdateProgress mProgDone, mProgName & " / " & sourceFileName
    Exit Sub

CopyError:
    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description
    WriteErrorLogValues "単価一覧コピー: " & sourcePath, savedErrNumber, savedErrSource, savedErrDescription
    On Error Resume Next
    If Not copiedSheet Is Nothing Then
        Application.DisplayAlerts = False
        copiedSheet.Delete
    End If
    If createdTargetBook Then
        If Not targetBooks Is Nothing Then
            If targetBooks.Exists(targetPath) Then targetBooks.Remove targetPath
        End If
        If Not targetBook Is Nothing Then targetBook.Close SaveChanges:=False
    End If
    If sourceOpenedByTool Then
        If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Sub FormatPurchaseUnitPriceSheet(ByVal ws As Worksheet, ByVal sourceFileName As String, ByVal isChange As Boolean)
    Const FIRST_ROW As Long = 2
    Const HEADER_ROW As Long = 1

    Dim lastRow As Long
    Dim rowIndex As Long
    Dim d4 As String
    Dim f2 As String
    Dim h5 As String
    Dim code11 As String
    Dim lastCol As Long
    Dim newLastCol As Long
    Dim newLastRow As Long
    Dim tableRange As Range
    Dim tbl As ListObject
    Dim lo As ListObject
    Dim tableName As String
    Dim titleRange As Range
    Dim titleText As String
    Dim yearText As String
    Dim prefixText As String
    Dim areaText As String

    tableName = CleanListObjectName(CStr(ws.Cells(2, "C").Value))
    If Len(tableName) = 0 Then tableName = "品目コードTable"

    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If lastRow < FIRST_ROW Then Err.Raise vbObjectError + 30, , "購入充当単価表のD列2行目以降にデータがありません: " & sourceFileName

    ws.Range("A:A").NumberFormat = "@"
    For rowIndex = FIRST_ROW To lastRow
        If IsNumeric(ws.Cells(rowIndex, "D").Value) And Trim$(CStr(ws.Cells(rowIndex, "D").Value)) <> "" Then
            d4 = Format$(CLng(CDbl(ws.Cells(rowIndex, "D").Value)) + 1000, "0000")
        Else
            d4 = "0000"
        End If
        f2 = PadZeroText(OnlyDigitsNarrow(CStr(ws.Cells(rowIndex, "F").Value)), 2)
        h5 = PadZeroText(OnlyDigitsNarrow(CStr(ws.Cells(rowIndex, "H").Value)), 5)
        code11 = d4 & f2 & h5
        ws.Cells(rowIndex, "A").Value = code11
    Next rowIndex

    ws.Range("A:A").NumberFormat = "0"
    For rowIndex = FIRST_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(rowIndex, "A").Value))) > 0 Then ws.Cells(rowIndex, "A").Value = CDec(ws.Cells(rowIndex, "A").Value)
    Next rowIndex

    lastCol = ws.Cells(FIRST_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column Then lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    ws.Columns("C").Copy Destination:=ws.Columns(lastCol + 1)
    ws.Columns("C").UnMerge
    ws.Columns("C").ClearContents

    ws.Columns("M").Delete Shift:=xlShiftToLeft
    ws.Columns("L").Delete Shift:=xlShiftToLeft
    ws.Columns("H").Delete Shift:=xlShiftToLeft
    ws.Columns("F").Delete Shift:=xlShiftToLeft
    ws.Columns("D").Delete Shift:=xlShiftToLeft
    ws.Columns("C").Delete Shift:=xlShiftToLeft
    ws.Columns("B").Delete Shift:=xlShiftToLeft

    ws.Cells.VerticalAlignment = xlVAlignCenter

    newLastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    newLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set tableRange = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(newLastRow, newLastCol))

    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo

    Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    tbl.Name = UniqueListObjectName(ws.Parent, tableName)
    tbl.TableStyle = "TableStyleMedium7"

    ws.Rows("1:4").Insert Shift:=xlDown
    ws.Rows(5).HorizontalAlignment = xlHAlignCenter
    ws.Columns("A").HorizontalAlignment = xlHAlignCenter
    ws.Columns("E").HorizontalAlignment = xlHAlignCenter
    ws.Columns("G").HorizontalAlignment = xlHAlignCenter
    ws.Columns("H").HorizontalAlignment = xlHAlignCenter
    ws.Columns("F").NumberFormat = "#,##0"
    ws.Cells(5, "E").Value = "単位"
    ws.Cells(5, "F").Value = "単価"

    yearText = GetYearFromFileName(RemoveExtension(sourceFileName))
    If Len(yearText) = 0 Then
        yearText = OnlyDigitsNarrow(sourceFileName)
        If Len(yearText) >= 4 Then yearText = Left$(yearText, 4)
    End If
    prefixText = PurchaseTitlePrefix(sourceFileName, isChange)
    areaText = CStr(ws.Cells(6, "G").Value)
    titleText = yearText & "年度_購入充当単価表_" & prefixText & "_" & areaText

    Set titleRange = ws.Range("A1:G1")
    With titleRange
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Font.Name = "BIZ UDGothic"
        .Font.Size = 14
    End With

    ws.Cells.Font.Name = "BIZ UDGothic"
    On Error Resume Next
    ws.Cells.Font.NameFarEast = "BIZ UDGothic"
    On Error GoTo 0
    ws.Range("A1").Font.Size = 14
    ws.Cells.EntireColumn.AutoFit
End Sub

Private Function PurchaseTitlePrefix(ByVal sourceFileName As String, ByVal isChange As Boolean) As String
    If isChange Or InStr(sourceFileName, PURCHASE_TITLE_PREFIX_CHANGE) > 0 Then
        PurchaseTitlePrefix = PURCHASE_TITLE_PREFIX_CHANGE
    ElseIf InStr(sourceFileName, PURCHASE_TITLE_PREFIX_NORMAL) > 0 Then
        PurchaseTitlePrefix = PURCHASE_TITLE_PREFIX_NORMAL
    Else
        PurchaseTitlePrefix = PURCHASE_TITLE_PREFIX_NORMAL
    End If
End Function

Private Function IsPurchaseProjectName(ByVal projectName As String, ByVal projectCode As String) As Boolean
    IsPurchaseProjectName = (projectCode = PURCHASE_PROJECT_CODE Or InStr(projectName, "軌道材料購入充当") > 0)
End Function

Private Function NormalizePurchaseBranchSuffix(ByVal value As Variant) As String
    Dim text As String
    Dim digits As String
    Dim numericValue As Long

    text = Trim$(StrConv(CStr(value), vbNarrow))
    text = Replace$(text, "－", "-")
    text = Replace$(text, " ", "")
    If Len(text) = 0 Then Exit Function

    If IsNumeric(text) Then
        numericValue = CLng(CDbl(text))
        If numericValue = 0 Then Exit Function
        If numericValue < 0 Then
            NormalizePurchaseBranchSuffix = CStr(numericValue)
        Else
            NormalizePurchaseBranchSuffix = "-" & CStr(numericValue)
        End If
        Exit Function
    End If

    digits = OnlyDigitsNarrow(text)
    If Len(digits) = 0 Then Exit Function
    NormalizePurchaseBranchSuffix = "-" & CStr(CLng(CDbl(digits)))
End Function

Private Function OnlyDigitsNarrow(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    text = StrConv(text, vbNarrow)
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch Like "#" Then OnlyDigitsNarrow = OnlyDigitsNarrow & ch
    Next i
End Function

Private Function PadZeroText(ByVal value As String, ByVal width As Long) As String
    Dim text As String
    text = OnlyDigitsNarrow(value)
    If Len(text) = 0 Then text = "0"
    If Len(text) >= width Then
        PadZeroText = Right$(text, width)
    Else
        PadZeroText = String$(width - Len(text), "0") & text
    End If
End Function

Private Function CleanListObjectName(ByVal value As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String

    value = Trim$(value)
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch Like "[A-Za-z0-9_]" Or AscW(ch) > 255 Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i
    If Len(result) > 0 Then
        If Left$(result, 1) Like "#" Then result = "_" & result
    End If
    If Len(result) > 200 Then result = Left$(result, 200)
    CleanListObjectName = result
End Function

Private Function UniqueListObjectName(ByVal book As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As String
    Dim i As Long

    candidate = CleanListObjectName(baseName)
    If Len(candidate) = 0 Then candidate = "品目コードTable"
    If Not ListObjectNameExists(book, candidate) Then
        UniqueListObjectName = candidate
        Exit Function
    End If

    For i = 2 To 999
        suffix = "_" & CStr(i)
        candidate = Left$(CleanListObjectName(baseName), 200 - Len(suffix)) & suffix
        If Not ListObjectNameExists(book, candidate) Then
            UniqueListObjectName = candidate
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 31, , "テーブル名を作成できません: " & baseName
End Function

Private Function ListObjectNameExists(ByVal book As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In book.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                ListObjectNameExists = True
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Sub SaveTargetBooks(ByVal targetBooks As Object)
    Dim key As Variant
    Dim book As Workbook
    Dim targetPath As String

    WriteLog "作成対象ファイル数: " & CStr(targetBooks.Count)

    For Each key In targetBooks.Keys
        targetPath = CStr(key)
        frmProgress.UpdateProgress mProgDone, "保存: " & GetFileName(targetPath)
        Set book = targetBooks(key)
        SaveOneBook book, targetPath
        frmProgress.UpdateProgress mProgDone, "保存完了: " & GetFileName(targetPath)
    Next key
End Sub

' 1ブックを一時ファイル名で保存し、成功後に正式名へ変更する
Private Sub SaveOneBook(ByVal book As Workbook, ByVal targetPath As String)
    On Error GoTo SaveErr
    Dim targetDir As String
    Dim tempPath As String

    targetDir = Left$(targetPath, InStrRev(targetPath, "\") - 1)
    EnsureFolder targetDir
    If IsWorkbookPathOpen(targetPath) Then Err.Raise vbObjectError + 20, , "出力ブックが開いています。閉じてから再実行してください: " & targetPath

    If FileExists(targetPath) Then
        WriteLog "スキップ: 保存時点で出力ファイルが既に存在します: " & targetPath
        book.Close SaveChanges:=False
        Exit Sub
    End If

    tempPath = BuildTempPath(targetPath)
    If IsWorkbookPathOpen(tempPath) Then Err.Raise vbObjectError + 22, , "一時保存ブックが開いています。閉じてから再実行してください: " & tempPath

    ApplyWorkbookFont book, "BIZ UDゴシック"
    WriteLog "一時保存開始: " & tempPath
    FlushLog
    book.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook
    book.Close SaveChanges:=False
    Set book = Nothing

    If FileExists(targetPath) Then
        On Error Resume Next
        Kill tempPath
        On Error GoTo SaveErr
        WriteLog "スキップ: 一時保存後に出力ファイルが作成済みです: " & targetPath
        Exit Sub
    End If

    FileCopy tempPath, targetPath
    Kill tempPath
    WriteLog "保存完了: " & targetPath
    FlushLog
    Exit Sub

SaveErr:
    Dim savedNumber As Long
    Dim savedSource As String
    Dim savedDescription As String

    savedNumber = Err.Number
    savedSource = Err.Source
    savedDescription = Err.Description
    WriteErrorLogValues "保存失敗: " & targetPath, savedNumber, savedSource, savedDescription
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    If Len(tempPath) > 0 Then
        If FileExists(tempPath) Then Kill tempPath
    End If
    On Error GoTo 0
    Err.Raise savedNumber, savedSource, savedDescription
End Sub

Private Function BuildTempPath(ByVal targetPath As String) As String
    Dim tempDir As String
    Dim tempPath As String
    Dim i As Long

    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then tempDir = Environ$("TMP")
    If Len(tempDir) = 0 Then Err.Raise vbObjectError + 24, , "一時保存フォルダを取得できません。"
    tempDir = CombinePath(tempDir, "UnitPriceListTool")
    EnsureFolder tempDir

    For i = 1 To 999
        tempPath = CombinePath(tempDir, "unit-price-list_" & Format(Now, "yyyymmdd_hhnnss") & "_" & Format$(i, "000") & ".xlsx")
        If Not FileExists(tempPath) Then
            BuildTempPath = tempPath
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 23, , "一時保存ファイル名を作成できません: " & targetPath
End Function
' 出力先ブックが開いているか確認する
Private Function IsWorkbookPathOpen(ByVal targetPath As String) As Boolean
    Dim wb As Workbook

    For Each wb In Workbooks
        If LCase$(wb.FullName) = LCase$(targetPath) Then
            IsWorkbookPathOpen = True
            Exit Function
        End If
    Next wb
End Function

Private Sub ApplyWorkbookFont(ByVal book As Workbook, ByVal fontName As String)
    Dim sheet As Worksheet

    For Each sheet In book.Worksheets
        sheet.UsedRange.Font.Name = fontName
        On Error Resume Next
        sheet.UsedRange.Font.NameFarEast = fontName
        On Error GoTo 0
    Next sheet
End Sub

Private Function BuildTargetPath(ByVal outputLineDir As String, ByVal branchName As String, ByVal areaFolderName As String, ByVal yearText As String, ByVal isChange As Boolean, ByVal projectName As String) As String
    Dim kindFolder As String
    Dim filePrefix As String
    Dim targetDir As String

    If isChange Then
        kindFolder = "設計変更単価"
        filePrefix = "設計変更"
    Else
        kindFolder = "通常単価"
        filePrefix = "通常"
    End If

    If Len(outputLineDir) = 0 Then outputLineDir = OUTPUT_CONVENTIONAL_LINE_DIR
    targetDir = CombinePath(CombinePath(CombinePath(mRootPath, OUTPUT_DATA_DIR), outputLineDir), branchName)
    targetDir = CombinePath(targetDir, SafeFileName(areaFolderName))
    targetDir = CombinePath(targetDir, yearText)
    targetDir = CombinePath(targetDir, kindFolder)

    BuildTargetPath = CombinePath(targetDir, filePrefix & "_" & yearText & "_" & SafeFileName(areaFolderName) & "_" & SafeFileName(projectName) & ".xlsx")
End Function

Private Function GetProjectCode(ByVal projectName As String) As String
    Dim firstChar As String
    firstChar = Left$(Trim$(projectName), 1)

    Select Case firstChar
        Case "①": GetProjectCode = "1"
        Case "②": GetProjectCode = "2"
        Case "③": GetProjectCode = "3"
        Case "④": GetProjectCode = "4"
        Case "⑤": GetProjectCode = "5"
        Case "⑥": GetProjectCode = "6"
        Case "⑦": GetProjectCode = "7"
        Case "⑧": GetProjectCode = "8"
        Case "⑨": GetProjectCode = "9"
        Case "⑩": GetProjectCode = "10"
        Case "⑪": GetProjectCode = "11"
        Case "⑫": GetProjectCode = "12"
        Case "⑬": GetProjectCode = "13"
        Case "⑭": GetProjectCode = "14"
        Case "⑮": GetProjectCode = "15"
        Case "⑯": GetProjectCode = "16"
        Case "⑰": GetProjectCode = "17"
        Case "⑱": GetProjectCode = "18"
        Case "⑲": GetProjectCode = "19"
        Case "⑳": GetProjectCode = "20"
        Case Else
            If firstChar Like "#" Then GetProjectCode = firstChar
    End Select
End Function

Private Function GetYearFromSourcePath(ByVal sourcePath As String) As String
    Dim normalizedPath As String
    Dim parts() As String
    Dim i As Long
    Dim candidate As String

    normalizedPath = TrimTrailingSlash(Replace(sourcePath, "/", "\"))
    parts = Split(normalizedPath, "\")

    For i = UBound(parts) To LBound(parts) Step -1
        candidate = GetYearFromSourceFolderName(CStr(parts(i)))
        If Len(candidate) > 0 Then
            GetYearFromSourcePath = candidate
            Exit Function
        End If
    Next i
End Function

Private Function GetYearFromSourceFolderName(ByVal folderName As String) As String
    Dim candidate As String

    candidate = FindYearTextFromLeft(folderName)
    If Len(candidate) > 0 Then
        GetYearFromSourceFolderName = candidate
        Exit Function
    End If

    GetYearFromSourceFolderName = GetWesternYearFromJapaneseEra(folderName)
End Function

Private Function GetWesternYearFromJapaneseEra(ByVal text As String) As String
    Dim normalizedText As String
    Dim pos As Long
    Dim startPos As Long
    Dim i As Long
    Dim ch As String
    Dim digits As String
    Dim eraYear As Long

    normalizedText = UCase$(StrConv(text, vbNarrow))
    pos = InStr(normalizedText, "R")
    If pos > 0 Then
        startPos = pos + Len("R")
    Else
        pos = InStr(text, "令和")
        If pos = 0 Then Exit Function
        startPos = pos + Len("令和")
    End If

    For i = startPos To Len(normalizedText)
        ch = Mid$(normalizedText, i, 1)
        If ch Like "#" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        ElseIf ch = "元" Then
            digits = "1"
            Exit For
        End If
    Next i

    If Len(digits) = 0 Then Exit Function
    eraYear = CLng(digits)
    If eraYear <= 0 Or eraYear > 99 Then Exit Function
    GetWesternYearFromJapaneseEra = CStr(2018 + eraYear)
End Function
Private Function GetYearFromFileName(ByVal baseName As String) As String
    Dim pos As Long
    Dim changePos As Long
    Dim candidate As String
    Dim beforeListName As String
    Dim afterListName As String

    pos = InStr(baseName, "単価一覧表")
    If pos > 0 Then
        beforeListName = Left$(baseName, pos - 1)
        afterListName = Mid$(baseName, pos + Len("単価一覧表"))
    Else
        afterListName = baseName
    End If

    changePos = InStr(afterListName, "(変)")
    If changePos > 0 Then
        candidate = FindYearTextFromLeft(Mid$(afterListName, changePos + Len("(変)")))
        If Len(candidate) > 0 Then
            GetYearFromFileName = candidate
            Exit Function
        End If
    End If

    candidate = FindYearTextFromRight(afterListName)
    If Len(candidate) > 0 Then
        GetYearFromFileName = candidate
        Exit Function
    End If

    candidate = FindYearTextFromRight(beforeListName)
    If Len(candidate) > 0 Then GetYearFromFileName = candidate
End Function

Private Function FindYearTextFromLeft(ByVal text As String) As String
    Dim i As Long
    Dim candidate As String

    text = StrConv(text, vbNarrow)
    For i = 1 To Len(text) - 3
        candidate = Mid$(text, i, 4)
        If candidate Like "####" Then
            FindYearTextFromLeft = candidate
            Exit Function
        End If
    Next i
End Function

Private Function FindYearTextFromRight(ByVal text As String) As String
    Dim i As Long
    Dim candidate As String

    text = StrConv(text, vbNarrow)
    For i = Len(text) - 3 To 1 Step -1
        candidate = Mid$(text, i, 4)
        If candidate Like "####" Then
            FindYearTextFromRight = candidate
            Exit Function
        End If
    Next i
End Function
Private Function FormatCode(ByVal value As Variant, ByVal width As Long) As String
    Dim text As String
    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function
    If IsNumeric(text) Then
        FormatCode = Right$(String$(width, "0") & CStr(CLng(CDbl(text))), width)
    Else
        FormatCode = Right$(String$(width, "0") & OnlyDigits(text), width)
    End If
End Function

Private Function OnlyDigits(ByVal text As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch Like "#" Then OnlyDigits = OnlyDigits & ch
    Next i
End Function

Private Function FindChildFolder(ByVal parentPath As String, ByVal exactName As String, ByVal prefix As String) As String
    Dim name As String
    Dim fullPath As String

    fullPath = CombinePath(parentPath, exactName)
    If FolderExists(fullPath) Then
        FindChildFolder = fullPath
        Exit Function
    End If

    name = Dir$(CombinePath(parentPath, "*"), vbDirectory)
    Do While Len(name) > 0
        If name <> "." And name <> ".." Then
            fullPath = CombinePath(parentPath, name)
            If FolderExists(fullPath) Then
                If name = exactName Then
                    FindChildFolder = fullPath
                    Exit Function
                End If
                If Len(prefix) > 0 Then
                    If Left$(name, Len(prefix)) = prefix Then
                        FindChildFolder = fullPath
                        Exit Function
                    End If
                End If
            End If
        End If
        name = Dir$()
    Loop
End Function

Private Function UniqueSheetName(ByVal book As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim i As Long
    Dim suffix As String

    candidate = SafeSheetName(baseName)
    If Not SheetExists(book, candidate) Then
        UniqueSheetName = candidate
        Exit Function
    End If

    For i = 2 To 999
        suffix = " (" & CStr(i) & ")"
        candidate = Left$(SafeSheetName(baseName), 31 - Len(suffix)) & suffix
        If Not SheetExists(book, candidate) Then
            UniqueSheetName = candidate
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 10, , "シート名を作成できません: " & baseName
End Function

Private Function SheetExists(ByVal book As Workbook, ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet
    For Each sheet In book.Worksheets
        If sheet.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next sheet
End Function

Private Function SafeSheetName(ByVal value As String) As String
    Dim result As String
    result = Trim$(value)
    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")
    result = Replace(result, ":", "_")
    If Len(result) = 0 Then result = "Sheet"
    If Len(result) > 31 Then result = Left$(result, 31)
    SafeSheetName = result
End Function

Private Function SafeFileName(ByVal value As String) As String
    Dim result As String
    result = Trim$(value)
    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    SafeFileName = result
End Function


Private Sub AddUnique(ByVal collection As Collection, ByVal value As String)
    Dim item As Variant
    For Each item In collection
        If CStr(item) = value Then Exit Sub
    Next item
    collection.Add value
End Sub

Private Function CombinePath(ByVal leftPath As String, ByVal rightPath As String) As String
    If Right$(leftPath, 1) = "\" Then
        CombinePath = leftPath & rightPath
    Else
        CombinePath = leftPath & "\" & rightPath
    End If
End Function

Private Function TrimTrailingSlash(ByVal path As String) As String
    Do While Right$(path, 1) = "\" And Len(path) > 3
        path = Left$(path, Len(path) - 1)
    Loop
    TrimTrailingSlash = path
End Function

Private Function GetFileName(ByVal path As String) As String
    path = TrimTrailingSlash(path)
    GetFileName = Mid$(path, InStrRev(path, "\") + 1)
End Function

Private Function RemoveExtension(ByVal fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        RemoveExtension = Left$(fileName, pos - 1)
    Else
        RemoveExtension = fileName
    End If
End Function

Private Function FileExists(ByVal path As String) As Boolean
    Dim attr As Long

    On Error Resume Next
    attr = GetAttr(path)
    If Err.Number = 0 Then FileExists = ((attr And vbDirectory) = 0)
    On Error GoTo 0
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim parentPath As String

    If FolderExists(path) Then Exit Sub
    parentPath = Left$(path, InStrRev(path, "\") - 1)
    If Len(parentPath) > 0 And Not FolderExists(parentPath) Then EnsureFolder parentPath
    MkDir path
End Sub
