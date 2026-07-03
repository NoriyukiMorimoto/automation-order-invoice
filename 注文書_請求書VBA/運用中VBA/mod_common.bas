Option Explicit

Public Function CommonNormalizeText(ByVal value As String) As String
    Dim s As String
    s = Trim$(Replace$(Replace$(Replace$(value, vbCr, ""), vbLf, ""), vbTab, " "))
    s = Replace$(s, ChrW$(&H3000), " ")
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    CommonNormalizeText = Trim$(s)
End Function

Public Function CommonRemoveAllSpaces(ByVal value As String) As String
    Dim s As String
    s = Replace$(value, " ", "")
    s = Replace$(s, ChrW$(&H3000), "")
    s = Replace$(s, vbTab, "")
    s = Replace$(s, vbCr, "")
    s = Replace$(s, vbLf, "")
    CommonRemoveAllSpaces = s
End Function

Public Function CommonNzText(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        CommonNzText = ""
    Else
        CommonNzText = CStr(value)
    End If
End Function

Public Function CommonExtractYear4Digits(ByVal sourceText As String) As String
    Dim result As String
    Dim i As Long, ch As String

    result = ""
    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If ch >= "0" And ch <= "9" Then
            result = result & ch
            If Len(result) >= 4 Then Exit For
        End If
    Next i

    If Len(result) = 4 Then
        CommonExtractYear4Digits = result
    Else
        CommonExtractYear4Digits = ""
    End If
End Function

Public Function CommonBasicInfoSheetNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H57FA) & ChrW$(&H672C) & ChrW$(&H60C5) & ChrW$(&H5831)
    End If
    CommonBasicInfoSheetNameText = cached
End Function

Public Function CommonCoverSheetNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H8868) & ChrW$(&H7D19)
    End If
    CommonCoverSheetNameText = cached
End Function

Public Function CommonPurchaseOrderOutputSheetName() As String
    Static cached As String
    If Len(cached) = 0 Then
        cached = ChrW$(&H8CFC) & ChrW$(&H5165) & ChrW$(&H5145) & ChrW$(&H5F53) & _
                 ChrW$(&H6307) & ChrW$(&H793A)
    End If
    CommonPurchaseOrderOutputSheetName = cached
End Function

Public Function CommonPurchaseNoticeOutputSheetName() As String
    Static cached As String
    If Len(cached) = 0 Then
        cached = ChrW$(&H8CFC) & ChrW$(&H5165) & ChrW$(&H5145) & ChrW$(&H5F53) & _
                 ChrW$(&H901A) & ChrW$(&H77E5)
    End If
    CommonPurchaseNoticeOutputSheetName = cached
End Function

Public Function CommonConstructionOrderWorksSheetName() As String
    Static cached As String
    If Len(cached) = 0 Then
        cached = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H6307) & ChrW$(&H793A) & ChrW$(&H66F8) & _
                 "(" & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ")"
    End If
    CommonConstructionOrderWorksSheetName = cached
End Function

Public Function CommonConstructionOrderWeldingSheetName() As String
    Static cached As String
    If Len(cached) = 0 Then
        cached = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H6307) & ChrW$(&H793A) & ChrW$(&H66F8) & _
                 "(" & ChrW$(&H6EB6) & ChrW$(&H63A5) & ")"
    End If
    CommonConstructionOrderWeldingSheetName = cached
End Function

Public Function CommonConstructionNoticeWorksSheetName() As String
    Static cached As String
    If Len(cached) = 0 Then
        cached = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H901A) & ChrW$(&H77E5) & ChrW$(&H66F8) & _
                 "(" & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ")"
    End If
    CommonConstructionNoticeWorksSheetName = cached
End Function

Public Function CommonConstructionNoticeWeldingSheetName() As String
    Static cached As String
    If Len(cached) = 0 Then
        cached = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H901A) & ChrW$(&H77E5) & ChrW$(&H66F8) & _
                 "(" & ChrW$(&H6EB6) & ChrW$(&H63A5) & ")"
    End If
    CommonConstructionNoticeWeldingSheetName = cached
End Function

Public Function CommonCompanyNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5927) & ChrW$(&H9244) & ChrW$(&H5DE5) & ChrW$(&H696D) & _
                 ChrW$(&H682A) & ChrW$(&H5F0F) & ChrW$(&H4F1A) & ChrW$(&H793E)
    End If
    CommonCompanyNameText = cached
End Function

Public Function CommonBranchSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H652F) & ChrW$(&H5E97)
    End If
    CommonBranchSuffixText = cached
End Function

Public Function CommonUnitPriceProjectNameMasterSheetNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5358) & ChrW$(&H4FA1) & ChrW$(&H9069) & ChrW$(&H7528) & _
                 ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H4EF6) & ChrW$(&H540D) & _
                 ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF)
    End If
    CommonUnitPriceProjectNameMasterSheetNameText = cached
End Function

Public Function CommonGetBasicInfoWorksheet(Optional ByVal wb As Workbook = Nothing) As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim candidateNames As Variant
    candidateNames = Array(CommonBasicInfoSheetNameText(), CommonCoverSheetNameText(), "Cover")

    Dim i As Long, ws As Worksheet
    For i = LBound(candidateNames) To UBound(candidateNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.worksheets(CStr(candidateNames(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set CommonGetBasicInfoWorksheet = ws
            Exit Function
        End If
    Next i
End Function

Public Function CommonGetBasicInfoSheetName(Optional ByVal wb As Workbook = Nothing) As String
    Dim ws As Worksheet
    Set ws = CommonGetBasicInfoWorksheet(wb)
    If Not ws Is Nothing Then CommonGetBasicInfoSheetName = ws.Name
End Function

Public Function CommonGetExcelAdoConnectionString(ByVal sourceFilePath As String) As String
    Dim ext As String
    ext = LCase$(Mid$(sourceFilePath, InStrRev(sourceFilePath, ".") + 1))

    Dim excelProps As String
    If ext = "xls" Then
        excelProps = "Excel 8.0;HDR=NO;IMEX=1"
    ElseIf ext = "xlsm" Then
        excelProps = "Excel 12.0 Macro;HDR=NO;IMEX=1"
    Else
        excelProps = "Excel 12.0 Xml;HDR=NO;IMEX=1"
    End If

    CommonGetExcelAdoConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                                        sourceFilePath & ";Extended Properties=""" & excelProps & """;"
End Function

Public Function CommonOpenExcelAdoConnection(ByVal sourceFilePath As String) As Object
    On Error GoTo ErrorHandler

    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open CommonGetExcelAdoConnectionString(sourceFilePath)

    Set CommonOpenExcelAdoConnection = cn
    Exit Function

ErrorHandler:
    Set CommonOpenExcelAdoConnection = Nothing
    Err.Clear
End Function

Public Sub CommonCloseAdoConnection(ByVal cn As Object)
    On Error Resume Next
    If Not cn Is Nothing Then
        If cn.State <> 0 Then cn.Close
    End If
    On Error GoTo 0
End Sub

Public Sub CommonCloseAdoRecordset(ByVal rs As Object)
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State <> 0 Then rs.Close
    End If
    On Error GoTo 0
End Sub

Public Function CommonGetAdoWorksheetNames(ByVal cn As Object) As Collection
    Const AD_SCHEMA_TABLES As Long = 20

    Dim sheetNames As Collection
    Set sheetNames = New Collection

    Dim schemaRows As Object
    On Error GoTo Cleanup

    Set schemaRows = cn.OpenSchema(AD_SCHEMA_TABLES)

    Do Until schemaRows.EOF
        Dim tableName As String
        tableName = CommonNzText(schemaRows.Fields("TABLE_NAME").value)
        If CommonIsAdoWorksheetTableName(tableName) Then
            sheetNames.Add CommonCleanAdoWorksheetName(tableName)
        End If
        schemaRows.MoveNext
    Loop

Cleanup:
    CommonCloseAdoRecordset schemaRows
    Set CommonGetAdoWorksheetNames = sheetNames
End Function

Public Function CommonIsAdoWorksheetTableName(ByVal tableName As String) As Boolean
    Dim t As String
    t = Replace$(tableName, "'", "")
    CommonIsAdoWorksheetTableName = (Right$(t, 1) = "$" And _
                                     InStr(1, t, "_xlnm#", vbTextCompare) = 0)
End Function

Public Function CommonCleanAdoWorksheetName(ByVal tableName As String) As String
    Dim t As String
    t = Replace$(tableName, "'", "")
    If Right$(t, 1) = "$" Then t = Left$(t, Len(t) - 1)
    CommonCleanAdoWorksheetName = t
End Function

Public Function CommonGetAdoFieldValue(ByVal rs As Object, ByVal fieldIndex As Long) As Variant
    On Error Resume Next
    If fieldIndex >= 1 And fieldIndex <= rs.Fields.Count Then
        CommonGetAdoFieldValue = rs.Fields(fieldIndex).value
    ElseIf fieldIndex = 0 And rs.Fields.Count > 0 Then
        CommonGetAdoFieldValue = rs.Fields(0).value
    End If
    On Error GoTo 0
End Function

Public Function CommonTextFromChars(ParamArray charCodes() As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(charCodes) To UBound(charCodes)
        result = result & ChrW$(CLng(charCodes(i)))
    Next i
    CommonTextFromChars = result
End Function

' 工事現況表マスタ(【各支店工事番号データ】)のフォルダパスを解決する。
' 正本: ドキュメント\マスタデータ\【各支店工事番号データ】\
Public Function CommonGetProjectStatusDataFolderPath() As String
    Dim candidates As Collection
    Set candidates = CommonCollectProjectStatusFolderCandidates()

    CommonGetProjectStatusDataFolderPath = CommonFirstExistingProjectStatusFolderPath(candidates)
    If Len(CommonGetProjectStatusDataFolderPath) = 0 And candidates.Count > 0 Then
        CommonGetProjectStatusDataFolderPath = CStr(candidates(1))
    End If
End Function

Public Function CommonFindProjectStatusSourceFilePath(ByVal targetYear As String, _
                                                      ByVal branchNameForFile As String) As String
    If Len(targetYear) = 0 Or Len(branchNameForFile) = 0 Then Exit Function

    Dim fileName As String
    fileName = targetYear & "_" & branchNameForFile & CommonProjectStatusFileSuffixText()

    Dim candidates As Collection
    Set candidates = CommonCollectProjectStatusFolderCandidates()

    Dim candidate As Variant
    For Each candidate In candidates
        Dim sourcePath As String
        sourcePath = CStr(candidate) & fileName
        If CommonProjectStatusFileExists(sourcePath) Then
            CommonFindProjectStatusSourceFilePath = sourcePath
            Exit Function
        End If
    Next candidate
End Function

Public Function CommonProjectStatusFolderExists(ByVal folderPath As String) As Boolean
    If Len(folderPath) = 0 Then Exit Function

    Dim checkPath As String
    checkPath = CommonNormalizeProjectStatusPath(folderPath, False)

    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(checkPath)
    If Err.Number = 0 Then
        CommonProjectStatusFolderExists = ((attr And vbDirectory) = vbDirectory)
        Exit Function
    End If
    Err.Clear

    CommonProjectStatusFolderExists = (Len(Dir(checkPath & Chr$(92) & "*.xlsx", vbNormal)) > 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Function CommonProjectStatusFileExists(ByVal filePath As String) As Boolean
    If Len(filePath) = 0 Then Exit Function

    On Error Resume Next
    Dim attr As Long
    attr = GetAttr(filePath)
    If Err.Number = 0 Then
        CommonProjectStatusFileExists = ((attr And vbDirectory) = 0)
        Exit Function
    End If
    Err.Clear

    CommonProjectStatusFileExists = (Len(Dir(filePath, vbNormal)) > 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Function CommonProjectStatusFileSuffixText() As String
    Static cached As String
    If cached = "" Then
        cached = "_" & ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H73FE) & ChrW$(&H6CC1) & _
                 ChrW$(&H8868) & ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF) & ".xlsx"
    End If
    CommonProjectStatusFileSuffixText = cached
End Function

Private Function CommonCollectProjectStatusFolderCandidates() As Collection
    Set CommonCollectProjectStatusFolderCandidates = New Collection

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then
        userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    End If

    If Len(Trim$(userProfilePath)) > 0 Then
        CommonAppendProjectStatusFolderCandidate CommonCollectProjectStatusFolderCandidates, _
            CommonBuildProjectStatusDataFolderPathFromRoot(userProfilePath & Chr$(92) & CommonCompanyNameText() & Chr$(92) & _
                                                           CommonOrderInvoiceDocumentFolderText())
    End If

    Dim walkedPath As String
    walkedPath = CommonFindProjectStatusFolderFromWorkbookPath()
    If Len(walkedPath) > 0 Then
        CommonAppendProjectStatusFolderCandidate CommonCollectProjectStatusFolderCandidates, walkedPath
    End If

    If Len(ThisWorkbook.Path) > 0 And Left$(LCase$(ThisWorkbook.Path), 8) <> "https://" Then
        Dim currentPath As String
        Dim depth As Long
        currentPath = ThisWorkbook.Path
        For depth = 1 To 8
            If Len(currentPath) = 0 Then Exit For
            CommonAppendProjectStatusFolderCandidate CommonCollectProjectStatusFolderCandidates, _
                CommonBuildProjectStatusDataFolderPathFromRoot(currentPath)
            currentPath = CommonGetParentFolderPath(currentPath)
        Next depth
    End If
End Function

Private Function CommonBuildProjectStatusDataFolderPathFromRoot(ByVal documentRootPath As String) As String
    If Len(Trim$(documentRootPath)) = 0 Then Exit Function

    Dim folderPath As String
    folderPath = documentRootPath
    If Right$(folderPath, 1) <> Chr$(92) Then folderPath = folderPath & Chr$(92)
    folderPath = folderPath & CommonMasterDataFolderText() & Chr$(92) & CommonProjectStatusDataSubFolderText() & Chr$(92)
    CommonBuildProjectStatusDataFolderPathFromRoot = folderPath
End Function

Private Function CommonNormalizeProjectStatusPath(ByVal pathValue As String, ByVal keepTrailingSlash As Boolean) As String
    Dim normalizedPath As String
    normalizedPath = Trim$(pathValue)
    If Len(normalizedPath) = 0 Then Exit Function

    If keepTrailingSlash Then
        If Right$(normalizedPath, 1) <> Chr$(92) Then normalizedPath = normalizedPath & Chr$(92)
    ElseIf Right$(normalizedPath, 1) = Chr$(92) Then
        normalizedPath = Left$(normalizedPath, Len(normalizedPath) - 1)
    End If
    CommonNormalizeProjectStatusPath = normalizedPath
End Function

Private Function CommonGetParentFolderPath(ByVal folderPath As String) As String
    Dim normalizedPath As String
    normalizedPath = CommonNormalizeProjectStatusPath(folderPath, False)
    If Len(normalizedPath) = 0 Then Exit Function

    Dim slashPos As Long
    slashPos = InStrRev(normalizedPath, Chr$(92))
    If slashPos <= 1 Then Exit Function
    CommonGetParentFolderPath = Left$(normalizedPath, slashPos - 1)
End Function

Private Sub CommonAppendProjectStatusFolderCandidate(ByVal candidates As Collection, _
                                                     ByVal folderPath As String)
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

Private Function CommonFirstExistingProjectStatusFolderPath(ByVal candidates As Collection) As String
    Dim candidate As Variant
    For Each candidate In candidates
        If Len(CStr(candidate)) > 0 Then
            If Left$(LCase$(CStr(candidate)), 8) <> "https://" Then
                If CommonProjectStatusFolderExists(CStr(candidate)) Then
                    CommonFirstExistingProjectStatusFolderPath = CStr(candidate)
                    Exit Function
                End If
            End If
        End If
    Next candidate
End Function

Private Function CommonFindProjectStatusFolderFromWorkbookPath() As String
    Dim currentPath As String
    currentPath = ThisWorkbook.Path
    If Len(currentPath) = 0 Then Exit Function
    If Left$(LCase$(currentPath), 8) = "https://" Then Exit Function

    Dim depth As Long
    For depth = 1 To 8
        If Len(currentPath) = 0 Then Exit For
        Dim testPath As String
        testPath = CommonBuildProjectStatusDataFolderPathFromRoot(currentPath)
        If Len(testPath) > 0 Then
            If CommonProjectStatusFolderExists(testPath) Then
                CommonFindProjectStatusFolderFromWorkbookPath = testPath
                Exit Function
            End If
        End If
        currentPath = CommonGetParentFolderPath(currentPath)
    Next depth
End Function

Public Function CommonMasterDataFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H30DE) & ChrW$(&H30B9) & ChrW$(&H30BF) & ChrW$(&H30C7) & _
                 ChrW$(&H30FC) & ChrW$(&H30BF)
    End If
    CommonMasterDataFolderText = cached
End Function

Private Function CommonProjectStatusDataSubFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H3010) & ChrW$(&H5404) & ChrW$(&H652F) & ChrW$(&H5E97) & _
                 ChrW$(&H5DE5) & ChrW$(&H4E8B) & ChrW$(&H756A) & ChrW$(&H53F7) & _
                 ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF) & ChrW$(&H3011)
    End If
    CommonProjectStatusDataSubFolderText = cached
End Function

Public Function CommonOrderInvoiceDocumentFolderText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7DDA) & ChrW$(&H8DEF) & ChrW$(&H51FA) & ChrW$(&H5F35) & ChrW$(&H6240) & _
                 ChrW$(&H7528) & ChrW$(&H5F) & ChrW$(&H6CE8) & ChrW$(&H6587) & ChrW$(&H66F8) & ChrW$(&H5F) & _
                 ChrW$(&H8ACB) & ChrW$(&H6C42) & ChrW$(&H66F8) & ChrW$(&H30A2) & ChrW$(&H30AF) & ChrW$(&H30BB) & _
                 ChrW$(&H30B9) & ChrW$(&H30B5) & ChrW$(&H30A4) & ChrW$(&H30C8) & " - " & _
                 ChrW$(&H30C9) & ChrW$(&H30AD) & ChrW$(&H30E5) & ChrW$(&H30E1) & ChrW$(&H30F3) & ChrW$(&H30C8)
    End If
    CommonOrderInvoiceDocumentFolderText = cached
End Function
