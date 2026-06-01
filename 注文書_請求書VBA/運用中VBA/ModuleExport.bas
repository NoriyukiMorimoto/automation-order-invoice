Option Explicit

' Requires: Trust access to the VBA project object model
' No reference required: VBIDE objects are late-bound.

Private Const VBEXT_CT_STD_MODULE As Long = 1
Private Const VBEXT_CT_CLASS_MODULE As Long = 2
Private Const VBEXT_CT_MS_FORM As Long = 3
Private Const VBEXT_CT_DOCUMENT As Long = 100
Private Const EXPORT_ROOT_PATH As String = "C:\Users\n-morimoto\大鉄工業株式会社\線路出張所用_注文書_請求書アクセスサイト - ドキュメント\注文書_請求書VBA"

Public Sub Export_VBA_Code()
    Export_All_VBA_Components
End Sub

Public Sub Export_All_VBA_Components(Optional ByVal targetBook As Workbook)
    Dim wb As Workbook
    Dim vbc As Object
    Dim currentComponentName As String
    Dim baseRoot As String
    Dim exportRoot As String
    Dim timeStamp As String
    Dim ext As String
    Dim fileName As String
    Dim fso As Object

    On Error GoTo EH
    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")

    If targetBook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If

    baseRoot = GetExportRoot(wb)
    If Len(baseRoot) = 0 Then
        Err.Raise vbObjectError + 513, , "Please save the workbook to a local folder before exporting VBA code."
    End If

    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    exportRoot = fso.BuildPath(baseRoot, "VBA_Export_" & timeStamp)

    EnsureFolderExists fso, baseRoot
    EnsureFolderExists fso, exportRoot

    For Each vbc In wb.VBProject.VBComponents
        currentComponentName = CStr(vbc.Name)
        ext = GetComponentExtension(vbc)
        If Len(ext) > 0 Then
            fileName = GetSafeFileName(currentComponentName) & ext
            ExportComponent fso, vbc, fso.BuildPath(exportRoot, fileName)
        End If
    Next vbc

    MsgBox "VBA export completed." & vbCrLf & _
           "Export: " & exportRoot, vbInformation
    GoTo FinallyExit

EH:
    MsgBox "VBA export failed." & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description & vbCrLf & _
           "BookPath: " & GetWorkbookPathForMessage(wb) & vbCrLf & _
           "ExportRoot: " & exportRoot & vbCrLf & _
           "Component: " & currentComponentName, vbCritical

FinallyExit:
    Application.ScreenUpdating = True
End Sub

Private Function GetExportRoot(ByVal wb As Workbook) As String
    GetExportRoot = EXPORT_ROOT_PATH
    Exit Function

    If IsUsableLocalFolder(wb.Path) Then
        GetExportRoot = wb.Path
        Exit Function
    End If

    If IsUsableLocalFolder(ThisWorkbook.Path) Then
        GetExportRoot = ThisWorkbook.Path
        Exit Function
    End If

    If IsUsableLocalFolder(CurDir$) Then
        GetExportRoot = CurDir$
    End If
End Function

Private Function IsUsableLocalFolder(ByVal folderPath As String) As Boolean
    If Len(folderPath) = 0 Then Exit Function
    If InStr(1, folderPath, "://", vbTextCompare) > 0 Then Exit Function

    On Error Resume Next
    IsUsableLocalFolder = CreateObject("Scripting.FileSystemObject").FolderExists(folderPath)
    On Error GoTo 0
End Function

Private Function GetWorkbookPathForMessage(ByVal wb As Workbook) As String
    On Error Resume Next
    If wb Is Nothing Then
        GetWorkbookPathForMessage = ""
    Else
        GetWorkbookPathForMessage = wb.Path
    End If
    On Error GoTo 0
End Function

Private Function GetComponentExtension(ByVal vbc As Object) As String
    Select Case vbc.Type
        Case VBEXT_CT_STD_MODULE
            GetComponentExtension = ".bas"
        Case VBEXT_CT_CLASS_MODULE
            GetComponentExtension = ".cls"
        Case VBEXT_CT_MS_FORM
            GetComponentExtension = ".frm"
        Case VBEXT_CT_DOCUMENT
            GetComponentExtension = ".cls"
    End Select
End Function

Private Sub ExportComponent(ByVal fso As Object, ByVal vbc As Object, ByVal filePath As String)
    If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
    vbc.Export filePath
    RemoveAttributeLines fso, filePath
End Sub

Private Sub EnsureFolderExists(ByVal fso As Object, ByVal folderPath As String)
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Function GetSafeFileName(ByVal sourceName As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    Dim result As String
    result = sourceName

    Dim i As Long
    For i = LBound(invalidChars) To UBound(invalidChars)
        result = Replace$(result, CStr(invalidChars(i)), "_")
    Next i

    If Len(result) = 0 Then result = "VBAComponent"
    GetSafeFileName = result
End Function

Private Sub RemoveAttributeLines(ByVal fso As Object, ByVal filePath As String)
    Dim inputStream As Object
    Dim outputStream As Object
    Dim tempPath As String
    Dim lineText As String

    tempPath = filePath & ".tmp"

    Set inputStream = fso.OpenTextFile(filePath, 1, False, 0)
    Set outputStream = fso.CreateTextFile(tempPath, True, False)

    Do Until inputStream.AtEndOfStream
        lineText = inputStream.ReadLine
        If Left$(Trim$(lineText), 10) <> "Attribute " Then
            outputStream.WriteLine lineText
        End If
    Loop

    inputStream.Close
    outputStream.Close

    If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
    fso.MoveFile tempPath, filePath
End Sub


