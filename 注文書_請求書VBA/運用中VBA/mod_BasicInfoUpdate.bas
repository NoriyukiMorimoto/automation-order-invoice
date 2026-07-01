Option Explicit

Private Const BASIC_INFO_START_DATE_CELL As String = "F2"
Private Const BASIC_INFO_END_DATE_CELL As String = "F3"
Private Const BASIC_INFO_BILLING_COUNT_CELL As String = "F4"
Private Const BASIC_INFO_WORK_START_DATE_CELL As String = "C15"
Private Const BASIC_INFO_WORK_END_DATE_CELL As String = "C16"
Private Const BASIC_INFO_WORK_DAYS_CELL As String = "C17"
Private Const OFFICE_COMBO_NAME As String = "ComboBox1"
Private Const BASIC_INFO_CLEAR_RANGES As String = "C2,C9:C13,C15:C17,C20:C23,C24:C28,C31:C35,F2:F4,F10:F16,F18:F23,F25,F27,F29,F30,F31,F33,I30,I33,L30,O30,R30,U30,X30,AA30,AD30,AG30"
Private Const BASIC_INFO_BILLING_SEQUENCE_CELL As String = "F9"

Public Sub UpdateBasicInfoPeriod()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox "Basic information sheet was not found.", vbExclamation
        Exit Sub
    End If

    Dim savedErrNum As Long
    Dim savedErrDesc As String
    Dim previousScreenUpdating As Boolean
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    If Not IsDate(wsInfo.Range(BASIC_INFO_START_DATE_CELL).value) Then
        Err.Raise vbObjectError + 515, , "Start date is not a valid date."
    End If
    If Not IsDate(wsInfo.Range(BASIC_INFO_END_DATE_CELL).value) Then
        Err.Raise vbObjectError + 516, , "End date is not a valid date."
    End If

    Dim oldStartDate As Date
    oldStartDate = CDate(wsInfo.Range(BASIC_INFO_START_DATE_CELL).value)

    Dim oldEndDate As Date
    oldEndDate = CDate(wsInfo.Range(BASIC_INFO_END_DATE_CELL).value)

    Dim nextBillingCount As Long
    nextBillingCount = GetNextBillingCount(wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value)

    wsInfo.Range(BASIC_INFO_START_DATE_CELL).value = GetNextStartDate(oldStartDate)
    wsInfo.Range(BASIC_INFO_END_DATE_CELL).value = GetNextEndDate(oldEndDate)
    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_START_DATE_CELL)
    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_END_DATE_CELL)

    wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value = nextBillingCount

    GoTo FinallyExit

ErrorHandler:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description

FinallyExit:
    If Not wsInfo Is Nothing Then HideOfficeComboBoxForUpdate wsInfo
    Application.ScreenUpdating = previousScreenUpdating
    If savedErrNum <> 0 Then
        MsgBox "Basic information update failed." & vbCrLf & savedErrDesc, vbExclamation
    End If
End Sub

Public Sub ApplyInitialBillingPeriodFromStartDate(ByVal wsInfo As Worksheet)
    On Error GoTo ErrorHandler

    If wsInfo Is Nothing Then Exit Sub
    If Not IsDate(wsInfo.Range(BASIC_INFO_START_DATE_CELL).value) Then Exit Sub

    Dim startDate As Date
    startDate = CDate(wsInfo.Range(BASIC_INFO_START_DATE_CELL).value)

    wsInfo.Range(BASIC_INFO_END_DATE_CELL).value = DateSerial(Year(startDate), Month(startDate) + 1, 16)
    wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value = 1

    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_START_DATE_CELL)
    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_END_DATE_CELL)
    Exit Sub

ErrorHandler:
    MsgBox "Initial billing period setup failed." & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub ApplyWorkDaysFromWorkDates(ByVal wsInfo As Worksheet)
    On Error GoTo ErrorHandler

    If wsInfo Is Nothing Then Exit Sub

    If Not IsDate(wsInfo.Range(BASIC_INFO_WORK_START_DATE_CELL).value) Or _
       Not IsDate(wsInfo.Range(BASIC_INFO_WORK_END_DATE_CELL).value) Then
        wsInfo.Range(BASIC_INFO_WORK_DAYS_CELL).ClearContents
        Exit Sub
    End If

    Dim startDate As Date
    startDate = CDate(wsInfo.Range(BASIC_INFO_WORK_START_DATE_CELL).value)

    Dim endDate As Date
    endDate = CDate(wsInfo.Range(BASIC_INFO_WORK_END_DATE_CELL).value)

    With wsInfo.Range(BASIC_INFO_WORK_DAYS_CELL)
        .NumberFormatLocal = "0" & ChrW$(&H65E5)
        .value = DateDiff("d", startDate, endDate) + 1
    End With
    Exit Sub

ErrorHandler:
    wsInfo.Range(BASIC_INFO_WORK_DAYS_CELL).ClearContents
End Sub

'  ClearBasicInfo
Public Sub ClearBasicInfo()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox "Basic information sheet was not found.", vbExclamation
        Exit Sub
    End If

    Dim savedErrNum As Long
    Dim savedErrDesc As String
    Dim prevEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    prevEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    HideOfficeComboBoxForUpdate wsInfo
    mod_BasicInfoGuide.ClearAllGuides wsInfo
    mod_MaterialPriceImport.ConfirmAndClearUnitPriceForBasicInfo wsInfo
    DeleteConstructionImportSheets False
    wsInfo.Range(BASIC_INFO_CLEAR_RANGES).ClearContents
    ClearVendorTotalCells wsInfo
    ClearBasicInfoYenTotalCells wsInfo
    wsInfo.Range(BASIC_INFO_BILLING_SEQUENCE_CELL).value = 1
    mod_VendorMaster.SyncVendorBlocksFromCount wsInfo
    ClearBasicInfoYenTotalCells wsInfo

    GoTo AfterClear

ErrorHandler:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description

AfterClear:
    On Error Resume Next
    If Not wsInfo Is Nothing Then mod_BasicInfoGuide.InitBasicInfoGuide wsInfo
    On Error GoTo 0

    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = True

FinallyExit:
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    If savedErrNum <> 0 Then
        MsgBox "Basic information clear failed." & vbCrLf & savedErrDesc, vbExclamation
    End If
End Sub

'  SilentClearBasicInfo
Public Sub SilentClearBasicInfo(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim savedErrNum As Long
    Dim savedErrDesc As String
    Dim prevEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    prevEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    HideOfficeComboBoxForUpdate wsInfo
    mod_BasicInfoGuide.ClearAllGuides wsInfo
    mod_MaterialPriceImport.SilentClearUnitPriceForBasicInfo wsInfo
    DeleteConstructionImportSheets False
    wsInfo.Range(BASIC_INFO_CLEAR_RANGES).ClearContents
    ClearVendorTotalCells wsInfo
    ClearBasicInfoYenTotalCells wsInfo
    wsInfo.Range(BASIC_INFO_BILLING_SEQUENCE_CELL).value = 1
    mod_VendorMaster.SyncVendorBlocksFromCount wsInfo
    ClearBasicInfoYenTotalCells wsInfo

    GoTo AfterClear

ErrorHandler:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description

AfterClear:
    On Error Resume Next
    If Not wsInfo Is Nothing Then mod_BasicInfoGuide.InitBasicInfoGuide wsInfo
    On Error GoTo 0

    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = True

FinallyExit:
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = previousScreenUpdating
    If savedErrNum <> 0 Then
        MsgBox "Basic information silent clear failed." & vbCrLf & savedErrDesc, vbExclamation
    End If
End Sub

'  DeleteConstructionImportSheets
Private Sub DeleteConstructionImportSheets(Optional ByVal showConfirm As Boolean = True)
    Dim wsInfo As Worksheet
    Dim targetNames As Collection
    Dim ws As Worksheet
    Dim i As Long
    Dim prevAlerts As Boolean
    Dim prevScreenUpdating As Boolean

    Set wsInfo = CommonGetBasicInfoWorksheet()

    Set targetNames = New Collection

    For Each ws In ThisWorkbook.Worksheets
        If Not wsInfo Is Nothing Then
            If ws Is wsInfo Then GoTo NextSheet
        End If

        If IsConstructionImportSheetForClear(ws) Then
            AppendUniqueSheetName targetNames, ws.Name
        End If

NextSheet:
    Next ws

    AppendUniqueSheetNameIfExists targetNames, CommonPurchaseOrderOutputSheetName()
    AppendUniqueSheetNameIfExists targetNames, CommonPurchaseNoticeOutputSheetName()

    If targetNames.Count = 0 Then Exit Sub

    If showConfirm Then
        Dim nameList As String
        For i = 1 To targetNames.Count
            nameList = nameList & "  " & ChrW$(&H30FB) & CStr(targetNames(i)) & vbCrLf
        Next i

        Dim ans As VbMsgBoxResult
        ans = MsgBox(ChrW$(&H4EE5) & ChrW$(&H4E0B) & ChrW$(&H306E) & ChrW$(&H65BD) & ChrW$(&H5DE5) & _
                     ChrW$(&H6307) & ChrW$(&H793A) & ChrW$(&H66F8) & ChrW$(&H30FB) & ChrW$(&H901A) & _
                     ChrW$(&H77E5) & ChrW$(&H66F8) & ChrW$(&H30B7) & ChrW$(&H30FC) & ChrW$(&H30C8) & _
                     ChrW$(&H3092) & ChrW$(&H524A) & ChrW$(&H9664) & ChrW$(&H3057) & ChrW$(&H307E) & _
                     ChrW$(&H3059) & ChrW$(&H3002) & ChrW$(&H3088) & ChrW$(&H308D) & ChrW$(&H3057) & _
                     ChrW$(&H3044) & ChrW$(&H3067) & ChrW$(&H3059) & ChrW$(&H304B) & ChrW$(&HFF1F) & _
                     vbCrLf & vbCrLf & nameList, _
                     vbYesNo + vbQuestion + vbDefaultButton2, _
                     ChrW$(&H30B7) & ChrW$(&H30FC) & ChrW$(&H30C8) & ChrW$(&H524A) & ChrW$(&H9664) & _
                     ChrW$(&H306E) & ChrW$(&H78BA) & ChrW$(&H8A8D))
        If ans <> vbYes Then Exit Sub
    End If

    If Not wsInfo Is Nothing Then wsInfo.Activate

    prevAlerts = Application.DisplayAlerts
    prevScreenUpdating = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    On Error Resume Next
    For i = targetNames.Count To 1 Step -1
        ThisWorkbook.Worksheets(CStr(targetNames(i))).Delete
    Next i
    On Error GoTo 0
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevAlerts
End Sub

Private Function IsConstructionImportSheetForClear(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function

    If IsKnownConstructionOutputTemplateSheet(ws.Name) Then
        IsConstructionImportSheetForClear = True
        Exit Function
    End If

    If mod_Construction_Order_Import.IsManagedConstructionImportOutputSheet(ws) Then
        IsConstructionImportSheetForClear = True
        Exit Function
    End If

    If StrComp(ws.Name, CommonPurchaseOrderOutputSheetName(), vbTextCompare) = 0 Then
        IsConstructionImportSheetForClear = True
        Exit Function
    End If

    If StrComp(ws.Name, CommonPurchaseNoticeOutputSheetName(), vbTextCompare) = 0 Then
        IsConstructionImportSheetForClear = True
        Exit Function
    End If

    On Error Resume Next
    If ws.Tab.ColorIndex = xlColorIndexNone Then Exit Function
    Dim tabColor As Long
    tabColor = ws.Tab.Color
    IsConstructionImportSheetForClear = _
        (tabColor = RGB(255, 255, 0)) Or _
        (tabColor = RGB(233, 241, 123))
    On Error GoTo 0
End Function

Private Function IsKnownConstructionOutputTemplateSheet(ByVal sheetName As String) As Boolean
    Dim key As String
    key = NormalizeOutputSheetNameKey(sheetName)
    If Len(key) = 0 Then Exit Function

    If StrComp(key, NormalizeOutputSheetNameKey(CommonConstructionOrderWorksSheetName()), vbTextCompare) = 0 Then
        IsKnownConstructionOutputTemplateSheet = True
        Exit Function
    End If
    If StrComp(key, NormalizeOutputSheetNameKey(CommonConstructionOrderWeldingSheetName()), vbTextCompare) = 0 Then
        IsKnownConstructionOutputTemplateSheet = True
        Exit Function
    End If
    If StrComp(key, NormalizeOutputSheetNameKey(CommonConstructionNoticeWorksSheetName()), vbTextCompare) = 0 Then
        IsKnownConstructionOutputTemplateSheet = True
        Exit Function
    End If
    If StrComp(key, NormalizeOutputSheetNameKey(CommonConstructionNoticeWeldingSheetName()), vbTextCompare) = 0 Then
        IsKnownConstructionOutputTemplateSheet = True
        Exit Function
    End If
    If StrComp(key, NormalizeOutputSheetNameKey(CommonPurchaseOrderOutputSheetName()), vbTextCompare) = 0 Then
        IsKnownConstructionOutputTemplateSheet = True
        Exit Function
    End If
    If StrComp(key, NormalizeOutputSheetNameKey(CommonPurchaseNoticeOutputSheetName()), vbTextCompare) = 0 Then
        IsKnownConstructionOutputTemplateSheet = True
    End If
End Function

Private Function NormalizeOutputSheetNameKey(ByVal sheetName As String) As String
    Dim t As String
    t = CommonNormalizeText(sheetName)
    t = Replace$(t, ChrW$(&HFF08), "(")
    t = Replace$(t, ChrW$(&HFF09), ")")
    NormalizeOutputSheetNameKey = t
End Function

Private Sub AppendUniqueSheetName(ByVal targetNames As Collection, ByVal sheetName As String)
    Dim i As Long
    For i = 1 To targetNames.Count
        If StrComp(CStr(targetNames(i)), sheetName, vbTextCompare) = 0 Then Exit Sub
    Next i
    targetNames.Add sheetName
End Sub

Private Sub AppendUniqueSheetNameIfExists(ByVal targetNames As Collection, ByVal sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    AppendUniqueSheetName targetNames, sheetName
End Sub

Private Sub HideOfficeComboBoxForUpdate(ByVal wsInfo As Worksheet)
    On Error Resume Next

    Dim ole As OLEObject
    Set ole = wsInfo.OLEObjects(OFFICE_COMBO_NAME)
    If ole Is Nothing Then Exit Sub

    ole.Object.value = CStr(wsInfo.Range("C6").value)
    ole.Visible = False

    On Error GoTo 0
End Sub

Private Function GetNextStartDate(ByVal oldDate As Date) As Date
    GetNextStartDate = DateSerial(Year(oldDate), Month(oldDate) + 1, 15)
End Function

Private Function GetNextEndDate(ByVal oldDate As Date) As Date
    GetNextEndDate = DateSerial(Year(oldDate), Month(oldDate) + 1, 16)
End Function

Private Function GetNextBillingCount(ByVal currentValue As Variant) As Long
    Dim currentText As String
    currentText = Trim$(CStr(currentValue))

    If currentText = "" Then
        GetNextBillingCount = 1
        Exit Function
    End If

    Dim digitText As String
    digitText = ExtractBillingCountDigits(currentText)
    If digitText = "" Then
        Err.Raise vbObjectError + 517, , "Billing count is not a valid number."
    End If

    GetNextBillingCount = CLng(digitText) + 1
End Function

Private Function ExtractBillingCountDigits(ByVal value As String) As String
    value = StrConv(value, vbNarrow)

    Dim i As Long
    Dim ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch >= "0" And ch <= "9" Then
            ExtractBillingCountDigits = ExtractBillingCountDigits & ch
        ElseIf ExtractBillingCountDigits <> "" Then
            Exit Function
        End If
    Next i
End Function

Private Sub ApplyJapaneseDateFormat(ByVal targetCell As Range)
    targetCell.NumberFormatLocal = "yyyy" & ChrW$(&H5E74) & "m" & ChrW$(&H6708) & "d" & ChrW$(&H65E5)
End Sub

Private Sub ClearVendorTotalCells(ByVal wsInfo As Worksheet)
    Const VENDOR_TOTAL_ROW As Long = 33
    Const VENDOR_TOTAL_FIRST_COL As Long = 6
    Const VENDOR_TOTAL_STEP_COLS As Long = 3
    Const VENDOR_TOTAL_MAX_BLOCKS As Long = 20

    Dim i As Long
    Dim targetCell As Range
    For i = 1 To VENDOR_TOTAL_MAX_BLOCKS
        Set targetCell = wsInfo.Cells(VENDOR_TOTAL_ROW, _
                                      VENDOR_TOTAL_FIRST_COL + ((i - 1) * VENDOR_TOTAL_STEP_COLS))
        If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)
        targetCell.ClearContents
    Next i
End Sub

' C31:C35(工事合計?税込合計)を結合セル対応で消去する。
' SyncVendorBlocksFromCount 内の合計再計算より後にも呼び、0 への再書込みを防ぐ。
Private Sub ClearBasicInfoYenTotalCells(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = 31 To 35
        Dim targetCell As Range
        Set targetCell = wsInfo.Cells(rowIndex, 3)
        If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)
        targetCell.ClearContents
    Next rowIndex
End Sub
