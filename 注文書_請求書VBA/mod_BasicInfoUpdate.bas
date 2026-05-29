Option Explicit

'==========================================================================
'  基本情報シート 期間／請求回数更新モジュール
'    改修内容（#15）：
'      - エラーハンドラ内で Err.Number / Err.Description を退避し、
'        その後の Application 設定復帰を確実に行う構造へ整理。
'    改修内容（#9）：
'      - NormalizeText / GetBasicInfoWorksheet / 日本語シート名生成は
'        mod_Common に集約済み。重複定義を撤去し、共通関数経由で参照。
'==========================================================================

Private Const BASIC_INFO_START_DATE_CELL As String = "C9"
Private Const BASIC_INFO_END_DATE_CELL As String = "C10"
Private Const BASIC_INFO_BILLING_COUNT_CELL As String = "C11"
Private Const BASIC_INFO_WORK_START_DATE_CELL As String = "C14"
Private Const BASIC_INFO_WORK_END_DATE_CELL As String = "C15"
Private Const BASIC_INFO_WORK_DAYS_CELL As String = "C16"
Private Const OFFICE_COMBO_NAME As String = "ComboBox1"
Private Const BASIC_INFO_CLEAR_RANGES As String = "C9:C11,C14:C16,C18:C21,F9:F14,F16:F21"

Public Sub UpdateBasicInfoPeriod()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox "Basic information sheet was not found.", vbExclamation
        Exit Sub
    End If

    Dim savedErrNum As Long
    Dim savedErrDesc As String

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

    wsInfo.Range(BASIC_INFO_START_DATE_CELL).value = GetNextStartDate(oldStartDate)
    wsInfo.Range(BASIC_INFO_END_DATE_CELL).value = GetNextEndDate(oldEndDate)
    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_START_DATE_CELL)
    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_END_DATE_CELL)

    If Len(Trim$(CStr(wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value))) = 0 Then
        wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value = 1
    Else
        wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value = CLng(wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).value) + 1
    End If

    GoTo FinallyExit

ErrorHandler:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description

FinallyExit:
    If Not wsInfo Is Nothing Then HideOfficeComboBoxForUpdate wsInfo
    Application.ScreenUpdating = True
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

Public Sub ClearBasicInfo()
    Dim wsInfo As Worksheet
    Set wsInfo = CommonGetBasicInfoWorksheet()
    If wsInfo Is Nothing Then
        MsgBox "Basic information sheet was not found.", vbExclamation
        Exit Sub
    End If

    Dim savedErrNum As Long
    Dim savedErrDesc As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    HideOfficeComboBoxForUpdate wsInfo
    wsInfo.Range(BASIC_INFO_CLEAR_RANGES).ClearContents

    GoTo FinallyExit

ErrorHandler:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description

FinallyExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If savedErrNum <> 0 Then
        MsgBox "Basic information clear failed." & vbCrLf & savedErrDesc, vbExclamation
    End If
End Sub

Private Sub HideOfficeComboBoxForUpdate(ByVal wsInfo As Worksheet)
    On Error Resume Next

    Dim ole As OLEObject
    Set ole = wsInfo.OLEObjects(OFFICE_COMBO_NAME)
    If ole Is Nothing Then Exit Sub

    If Len(Trim$(CStr(ole.Object.value))) > 0 Then
        wsInfo.Range("C6").value = ole.Object.value
    End If
    ole.Visible = False

    On Error GoTo 0
End Sub

Private Function GetNextStartDate(ByVal oldDate As Date) As Date
    GetNextStartDate = DateSerial(Year(oldDate), Month(oldDate) + 1, 15)
End Function

Private Function GetNextEndDate(ByVal oldDate As Date) As Date
    GetNextEndDate = DateSerial(Year(oldDate), Month(oldDate) + 1, 16)
End Function

Private Sub ApplyJapaneseDateFormat(ByVal targetCell As Range)
    targetCell.NumberFormatLocal = "yyyy" & ChrW$(&H5E74) & "m" & ChrW$(&H6708) & "d" & ChrW$(&H65E5)
End Sub
