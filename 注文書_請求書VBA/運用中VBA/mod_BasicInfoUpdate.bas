Option Explicit

Private Const BASIC_INFO_START_DATE_CELL As String = "F2"
Private Const BASIC_INFO_END_DATE_CELL As String = "F3"
Private Const BASIC_INFO_BILLING_COUNT_CELL As String = "F4"
Private Const BASIC_INFO_WORK_START_DATE_CELL As String = "C15"
Private Const BASIC_INFO_WORK_END_DATE_CELL As String = "C16"
Private Const BASIC_INFO_WORK_DAYS_CELL As String = "C17"
Private Const OFFICE_COMBO_NAME As String = "ComboBox1"
Private Const BASIC_INFO_CLEAR_RANGES As String = "C2,C9:C12,C15:C17,C20:C23,C24:C28,C31:C32,F2:F4,F11:F16,F18:F23,F25,F27,F29,F31,F33,I33"
Private Const BASIC_INFO_BILLING_SEQUENCE_CELL As String = "F9"

Private Const CONSTRUCTION_TAB_COLOR As Long = 65535
Private Const PURCHASE_ORDER_SHEET_NAME As String = "購入充当指示"
Private Const PURCHASE_NOTICE_SHEET_NAME As String = "購入充当通知"

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
    previousScreenUpdating = Application.screenUpdating

    On Error GoTo ErrorHandler
    Application.screenUpdating = False

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
    Application.screenUpdating = previousScreenUpdating
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
    previousScreenUpdating = Application.screenUpdating

    On Error GoTo ErrorHandler
    Application.screenUpdating = False
    Application.EnableEvents = False

    HideOfficeComboBoxForUpdate wsInfo
    mod_MaterialPriceImport.ConfirmAndClearUnitPriceForBasicInfo wsInfo
    wsInfo.Range(BASIC_INFO_CLEAR_RANGES).ClearContents
    ClearVendorTotalCells wsInfo
    wsInfo.Range(BASIC_INFO_BILLING_SEQUENCE_CELL).value = 1
    mod_VendorMaster.SyncVendorBlocksFromCount wsInfo

    Application.EnableEvents = prevEnableEvents
    DeleteConstructionImportSheets
    Application.EnableEvents = False

    GoTo FinallyExit

ErrorHandler:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description

FinallyExit:
    Application.EnableEvents = prevEnableEvents
    Application.screenUpdating = previousScreenUpdating
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
    previousScreenUpdating = Application.screenUpdating

    On Error GoTo ErrorHandler
    Application.screenUpdating = False
    Application.EnableEvents = False

    HideOfficeComboBoxForUpdate wsInfo
    mod_MaterialPriceImport.SilentClearUnitPriceForBasicInfo wsInfo
    wsInfo.Range(BASIC_INFO_CLEAR_RANGES).ClearContents
    ClearVendorTotalCells wsInfo
    wsInfo.Range(BASIC_INFO_BILLING_SEQUENCE_CELL).value = 1
    mod_VendorMaster.SyncVendorBlocksFromCount wsInfo

    GoTo FinallyExit

ErrorHandler:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description

FinallyExit:
    Application.EnableEvents = prevEnableEvents
    Application.screenUpdating = previousScreenUpdating
    If savedErrNum <> 0 Then
        MsgBox "Basic information silent clear failed." & vbCrLf & savedErrDesc, vbExclamation
    End If
End Sub

'  DeleteConstructionImportSheets
Private Sub DeleteConstructionImportSheets()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()

    Dim targets As Collection
    Set targets = New Collection

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not wsInfo Is Nothing Then
            If ws Is wsInfo Then GoTo NextSheet
        End If

        Dim hit As Boolean
        hit = False

        Dim tabColor As Long
        On Error Resume Next
        tabColor = ws.Tab.Color
        On Error GoTo 0
        If tabColor = CONSTRUCTION_TAB_COLOR Then hit = True

        If ws.Name = PURCHASE_ORDER_SHEET_NAME Then hit = True
        If ws.Name = PURCHASE_NOTICE_SHEET_NAME Then hit = True

        If hit Then targets.Add ws

NextSheet:
    Next ws

    If targets.Count = 0 Then Exit Sub

    Dim nameList As String
    Dim i As Long
    For i = 1 To targets.Count
        nameList = nameList & "  ・" & targets(i).Name & vbCrLf
    Next i

    Dim ans As VbMsgBoxResult
    ans = MsgBox("以下の施工指示書・通知書シートを削除します。よろしいですか？" & vbCrLf & vbCrLf & nameList, _
                 vbYesNo + vbQuestion + vbDefaultButton2, "シート削除の確認")
    If ans <> vbYes Then Exit Sub

    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    For i = 1 To targets.Count
        targets(i).Delete
    Next i
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts
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
