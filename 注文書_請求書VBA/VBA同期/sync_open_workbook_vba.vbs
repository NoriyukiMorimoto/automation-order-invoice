Option Explicit

Const ForReading = 1
Const TristateFalse = 0

Dim workbookPath, sourceFolder
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: sync_open_workbook_vba.vbs <workbook.xlsm> <source-folder>"
    WScript.Quit 1
End If

workbookPath = WScript.Arguments(0)
sourceFolder = WScript.Arguments(1)

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(workbookPath) Then
    Fail "Workbook was not found: " & workbookPath, 2
End If
If Not fso.FolderExists(sourceFolder) Then
    Fail "Source folder was not found: " & sourceFolder, 2
End If

Dim lockFilePath
lockFilePath = fso.BuildPath(fso.GetParentFolderName(workbookPath), "~$" & fso.GetFileName(workbookPath))
If Not fso.FileExists(lockFilePath) Then
    Fail "The workbook does not appear to be open: " & workbookPath, 3
End If

Dim componentNames, sourceFiles
componentNames = Array( _
    "ThisWorkbook", _
    "Sheet1", _
    "ModuleExport", _
    "mod_BasicInfoCalendar", _
    "mod_BasicInfoUpdate", _
    "mod_common", _
    "mod_DebugLog", _
    "mod_FillManagerName", _
    "mod_MaterialPriceImport", _
    "mod_UiMessages", _
    "mod_VendorMaster", _
    "AllVenderSelection", _
    "Project_Number_Selection", _
    "SelectLineName")

sourceFiles = Array( _
    "ThisWorkbook.cls", _
    "Sheet1.cls", _
    "ModuleExport.bas", _
    "mod_BasicInfoCalendar.bas", _
    "mod_BasicInfoUpdate.bas", _
    "mod_common.bas", _
    "mod_DebugLog.bas", _
    "mod_FillManagerName.bas", _
    "mod_MaterialPriceImport.bas", _
    "mod_UiMessages.bas", _
    "mod_VendorMaster.bas", _
    "AllVenderSelection.frm", _
    "Project_Number_Selection.frm", _
    "SelectLineName.frm")

Dim wb
Set wb = AttachToOpenWorkbook(workbookPath)
If wb Is Nothing Then
    Fail "The workbook is not open in the current Windows session: " & workbookPath, 3
End If

If wb.ReadOnly Then
    Fail "The workbook is open read-only: " & workbookPath, 4
End If

Dim excelApp
Set excelApp = wb.Application

Dim previousDisplayAlerts, previousEnableEvents, previousScreenUpdating
previousDisplayAlerts = excelApp.DisplayAlerts
previousEnableEvents = excelApp.EnableEvents
previousScreenUpdating = excelApp.ScreenUpdating

On Error Resume Next
excelApp.DisplayAlerts = False
excelApp.EnableEvents = False
excelApp.ScreenUpdating = False
On Error GoTo 0

Dim i, sourcePath, sourceText
Dim component, codeModule

For i = 0 To UBound(componentNames)
    sourcePath = fso.BuildPath(sourceFolder, sourceFiles(i))
    If Not fso.FileExists(sourcePath) Then
        RestoreApplicationState excelApp, previousDisplayAlerts, previousEnableEvents, previousScreenUpdating
        Fail "Source file was not found: " & sourcePath, 5
    End If

    sourceText = ReadAnsiText(sourcePath)

    On Error Resume Next
    Err.Clear
    Set component = wb.VBProject.VBComponents(componentNames(i))
    If Err.Number <> 0 Or component Is Nothing Then
        Dim componentError
        componentError = Err.Description
        On Error GoTo 0
        RestoreApplicationState excelApp, previousDisplayAlerts, previousEnableEvents, previousScreenUpdating
        Fail "VBA component was not found: " & componentNames(i) & " / " & componentError, 6
    End If

    Set codeModule = component.CodeModule
    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If
    codeModule.AddFromString sourceText
    If Err.Number <> 0 Then
        Dim updateError
        updateError = Err.Number & ": " & Err.Description
        On Error GoTo 0
        RestoreApplicationState excelApp, previousDisplayAlerts, previousEnableEvents, previousScreenUpdating
        Fail "Failed to update " & componentNames(i) & " / " & updateError, 7
    End If
    On Error GoTo 0

    WScript.Echo "UPDATED=" & componentNames(i) & "|" & codeModule.CountOfLines
    Set codeModule = Nothing
    Set component = Nothing
Next

On Error Resume Next
Err.Clear
wb.Save
If Err.Number <> 0 Then
    Dim saveError
    saveError = Err.Number & ": " & Err.Description
    On Error GoTo 0
    RestoreApplicationState excelApp, previousDisplayAlerts, previousEnableEvents, previousScreenUpdating
    Fail "Failed to save workbook / " & saveError, 8
End If
On Error GoTo 0

RestoreApplicationState excelApp, previousDisplayAlerts, previousEnableEvents, previousScreenUpdating

WScript.Echo "SAVED=" & wb.FullName
WScript.Echo "WORKBOOK_KEPT_OPEN=TRUE"

Set wb = Nothing
Set excelApp = Nothing
Set fso = Nothing

Function AttachToOpenWorkbook(ByVal targetPath)
    Dim excelObject, candidate, retryCount

    For retryCount = 1 To 5
        On Error Resume Next
        Set AttachToOpenWorkbook = GetObject(targetPath)
        On Error GoTo 0
        If Not AttachToOpenWorkbook Is Nothing Then Exit Function

        Set excelObject = Nothing
        On Error Resume Next
        Set excelObject = GetObject(, "Excel.Application")
        On Error GoTo 0

        If Not excelObject Is Nothing Then
            For Each candidate In excelObject.Workbooks
                If StrComp(candidate.FullName, targetPath, vbTextCompare) = 0 Then
                    Set AttachToOpenWorkbook = candidate
                    Exit Function
                End If
            Next
        End If

        WScript.Sleep 500
    Next

    Set AttachToOpenWorkbook = Nothing
End Function

Function ReadAnsiText(ByVal filePath)
    Dim stream
    Set stream = fso.OpenTextFile(filePath, ForReading, False, TristateFalse)
    ReadAnsiText = stream.ReadAll
    stream.Close
    Set stream = Nothing
End Function

Sub RestoreApplicationState(ByVal app, ByVal displayAlerts, ByVal enableEvents, ByVal screenUpdating)
    On Error Resume Next
    app.DisplayAlerts = displayAlerts
    app.EnableEvents = enableEvents
    app.ScreenUpdating = screenUpdating
    On Error GoTo 0
End Sub

Sub Fail(ByVal message, ByVal exitCode)
    WScript.Echo "ERROR=" & message
    WScript.Quit exitCode
End Sub
