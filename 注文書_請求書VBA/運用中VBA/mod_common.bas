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

Public Sub CommonSortLongArrayAsc(ByRef arr() As Long, ByVal itemCount As Long)
    Dim i As Long, j As Long, key As Long
    For i = 1 To itemCount - 1
        key = arr(i)
        j = i - 1
        Do While j >= 0
            If arr(j) <= key Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = key
    Next i
End Sub

Public Function CommonTextFromChars(ParamArray charCodes() As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(charCodes) To UBound(charCodes)
        result = result & ChrW$(CLng(charCodes(i)))
    Next i
    CommonTextFromChars = result
End Function
