Option Explicit

Public Function FindAdoSheetNameWUP(ByVal cn As Object, ByVal targetSheetName As String) As String
    Dim sheetNames As Collection
    Set sheetNames = CommonGetAdoWorksheetNames(cn)

    Dim normalizedTarget As String
    normalizedTarget = NormalizeMatchTextWUP(targetSheetName)

    Dim sheetName As Variant
    For Each sheetName In sheetNames
        If StrComp(NormalizeMatchTextWUP(CStr(sheetName)), normalizedTarget, vbTextCompare) = 0 Then
            FindAdoSheetNameWUP = CStr(sheetName)
            Exit Function
        End If
    Next sheetName

    ' ?????????????????
    For Each sheetName In sheetNames
        If InStr(1, NormalizeMatchTextWUP(CStr(sheetName)), normalizedTarget, vbTextCompare) > 0 Then
            FindAdoSheetNameWUP = CStr(sheetName)
            Exit Function
        End If
    Next sheetName
End Function

' ????

' ????????????mod_WeldingUnitPrice ???????
' ???????????Private ???????????????
Private Function NormalizeMatchTextWUP(ByVal value As String) As String
    NormalizeMatchTextWUP = CommonRemoveAllSpaces(CommonNormalizeText(value))
End Function


Public Function MasterDataFolderText() As String
    MasterDataFolderText = mod_common.CommonMasterDataFolderText()
End Function

Public Function OrderInvoiceDocumentFolderTextWUP() As String
    OrderInvoiceDocumentFolderTextWUP = mod_common.CommonOrderInvoiceDocumentFolderText()
End Function

Public Function ResolveTemotoMasterFilePath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folderCandidates As Collection
    Set folderCandidates = New Collection

    Dim userProfilePath As String
    userProfilePath = Environ$("USERPROFILE")
    If Len(Trim$(userProfilePath)) = 0 Then userProfilePath = Environ$("HOMEDRIVE") & Environ$("HOMEPATH")
    If Len(Trim$(userProfilePath)) > 0 Then
        folderCandidates.Add userProfilePath & "\" & CommonCompanyNameText() & "\" & _
                             OrderInvoiceDocumentFolderTextWUP() & "\" & MasterDataFolderText()
    End If
    If Len(ThisWorkbook.Path) > 0 Then
        folderCandidates.Add fso.GetParentFolderName(ThisWorkbook.Path) & "\" & MasterDataFolderText()
        folderCandidates.Add ThisWorkbook.Path & "\" & MasterDataFolderText()
    End If

    Dim folderPath As Variant
    For Each folderPath In folderCandidates
        Dim foundName As String
        foundName = ""
        On Error Resume Next
        foundName = Dir(CStr(folderPath) & "\" & TemotoMasterFilePatternText(), vbNormal)
        On Error GoTo 0
        If Len(foundName) > 0 Then
            ResolveTemotoMasterFilePath = CStr(folderPath) & "\" & foundName
            Exit Function
        End If
    Next folderPath
End Function

Public Function TemotoMasterFilePatternText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H30EC) & ChrW$(&H30FC) & ChrW$(&H30EB) & _
                 ChrW$(&H6EB6) & ChrW$(&H63A5) & "_" & _
                 ChrW$(&H8ECC) & ChrW$(&H9053) & ChrW$(&H4F1A) & ChrW$(&H793E) & _
                 ChrW$(&H5916) & ChrW$(&H6CE8) & ChrW$(&H8CBB) & ChrW$(&H7387) & _
                 ChrW$(&H4E00) & ChrW$(&H89A7) & "*.xlsx"
    End If
    TemotoMasterFilePatternText = cached
End Function

Public Function TemotoMasterSheetNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H6EB6) & ChrW$(&H63A5) & ChrW$(&H624B) & _
                 ChrW$(&H5143) & ChrW$(&H5272) & ChrW$(&H5408)
    End If
    TemotoMasterSheetNameText = cached
End Function
