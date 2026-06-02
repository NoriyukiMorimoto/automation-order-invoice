Option Explicit

'==========================================================================
'  基本情報シート 期間／請求回数更新モジュール
'    改修内容（#15）：
'      - エラーハンドラ内で Err.Number / Err.Description を退避し、
'        その後の Application 設定復帰を確実に行う構造へ整理。
'    改修内容（#9）：
'      - NormalizeText / GetBasicInfoWorksheet / 日本語シート名生成は
'        mod_Common に集約済み。重複定義を撤去し、共通関数経由で参照。
'    改修内容（#11）：
'      - SilentClearBasicInfo を追加。
'        B6/C6 変更時に確認メッセージなしで基本情報と単価シートをクリア。
'==========================================================================

Private Const BASIC_INFO_START_DATE_CELL As String = "C27"
Private Const BASIC_INFO_END_DATE_CELL As String = "C28"
Private Const BASIC_INFO_BILLING_COUNT_CELL As String = "C29"
Private Const BASIC_INFO_WORK_START_DATE_CELL As String = "C15"
Private Const BASIC_INFO_WORK_END_DATE_CELL As String = "C16"
Private Const BASIC_INFO_WORK_DAYS_CELL As String = "C17"
Private Const OFFICE_COMBO_NAME As String = "ComboBox1"
Private Const BASIC_INFO_CLEAR_RANGES As String = "C9:C12,C15:C17,C20:C24,C27:C29,F9:F14,F16:F21"

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

    If Not IsDate(wsInfo.Range(BASIC_INFO_START_DATE_CELL).Value) Then
        Err.Raise vbObjectError + 515, , "Start date is not a valid date."
    End If
    If Not IsDate(wsInfo.Range(BASIC_INFO_END_DATE_CELL).Value) Then
        Err.Raise vbObjectError + 516, , "End date is not a valid date."
    End If

    Dim oldStartDate As Date
    oldStartDate = CDate(wsInfo.Range(BASIC_INFO_START_DATE_CELL).Value)

    Dim oldEndDate As Date
    oldEndDate = CDate(wsInfo.Range(BASIC_INFO_END_DATE_CELL).Value)

    Dim nextBillingCount As Long
    nextBillingCount = GetNextBillingCount(wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).Value)

    wsInfo.Range(BASIC_INFO_START_DATE_CELL).Value = GetNextStartDate(oldStartDate)
    wsInfo.Range(BASIC_INFO_END_DATE_CELL).Value = GetNextEndDate(oldEndDate)
    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_START_DATE_CELL)
    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_END_DATE_CELL)

    wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).Value = nextBillingCount

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
    If Not IsDate(wsInfo.Range(BASIC_INFO_START_DATE_CELL).Value) Then Exit Sub

    Dim startDate As Date
    startDate = CDate(wsInfo.Range(BASIC_INFO_START_DATE_CELL).Value)

    wsInfo.Range(BASIC_INFO_END_DATE_CELL).Value = DateSerial(Year(startDate), Month(startDate) + 1, 16)
    wsInfo.Range(BASIC_INFO_BILLING_COUNT_CELL).Value = 1

    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_START_DATE_CELL)
    ApplyJapaneseDateFormat wsInfo.Range(BASIC_INFO_END_DATE_CELL)
    Exit Sub

ErrorHandler:
    MsgBox "Initial billing period setup failed." & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub ApplyWorkDaysFromWorkDates(ByVal wsInfo As Worksheet)
    On Error GoTo ErrorHandler

    If wsInfo Is Nothing Then Exit Sub

    If Not IsDate(wsInfo.Range(BASIC_INFO_WORK_START_DATE_CELL).Value) Or _
       Not IsDate(wsInfo.Range(BASIC_INFO_WORK_END_DATE_CELL).Value) Then
        wsInfo.Range(BASIC_INFO_WORK_DAYS_CELL).ClearContents
        Exit Sub
    End If

    Dim startDate As Date
    startDate = CDate(wsInfo.Range(BASIC_INFO_WORK_START_DATE_CELL).Value)

    Dim endDate As Date
    endDate = CDate(wsInfo.Range(BASIC_INFO_WORK_END_DATE_CELL).Value)

    With wsInfo.Range(BASIC_INFO_WORK_DAYS_CELL)
        .NumberFormatLocal = "0" & ChrW$(&H65E5)
        .Value = DateDiff("d", startDate, endDate) + 1
    End With
    Exit Sub

ErrorHandler:
    wsInfo.Range(BASIC_INFO_WORK_DAYS_CELL).ClearContents
End Sub

'--------------------------------------------------------------------------
'  ClearBasicInfo
'    ボタン押下用。確認メッセージあり・単価クリア確認あり。
'--------------------------------------------------------------------------
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
    mod_MaterialPriceImport.ConfirmAndClearUnitPriceForBasicInfo wsInfo
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

'--------------------------------------------------------------------------
'  SilentClearBasicInfo
'    B6/C6 変更時の自動クリア用。確認メッセージなし。
'    基本情報クリア範囲＋単価シート＋C23 を無条件で削除する。
'--------------------------------------------------------------------------
Public Sub SilentClearBasicInfo(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    Dim savedErrNum As Long
    Dim savedErrDesc As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    HideOfficeComboBoxForUpdate wsInfo
    mod_MaterialPriceImport.SilentClearUnitPriceForBasicInfo wsInfo
    wsInfo.Range(BASIC_INFO_CLEAR_RANGES).ClearContents

    GoTo FinallyExit

ErrorHandler:
    savedErrNum = Err.Number
    savedErrDesc = Err.Description

FinallyExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If savedErrNum <> 0 Then
        MsgBox "Basic information silent clear failed." & vbCrLf & savedErrDesc, vbExclamation
    End If
End Sub

Private Sub HideOfficeComboBoxForUpdate(ByVal wsInfo As Worksheet)
    On Error Resume Next

    Dim ole As OLEObject
    Set ole = wsInfo.OLEObjects(OFFICE_COMBO_NAME)
    If ole Is Nothing Then Exit Sub

    ' C6を手入力した直後に、ComboBoxに残った旧値で上書きしないよう
    ' セルの現在値をComboBox側へ同期してから非表示にする。
    ole.Object.Value = CStr(wsInfo.Range("C6").Value)
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
