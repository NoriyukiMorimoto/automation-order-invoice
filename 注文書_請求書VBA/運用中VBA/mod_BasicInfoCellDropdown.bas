Option Explicit

' ========================================================
' 基本情報シートの C22 / C23 をダブルクリックすると、
' そのセルの入力規則(リスト)からドロップダウンを表示する。
'
' 方式: SendKeys は使わず(NumLock保護)、セル上に一時 Forms.ComboBox.1
'       を重ねて .DropDown で開く(office C6 / 業者名 と同方式)。
' 選択値は LinkedCell と手動書込みの双方でセルへ反映する。
'
' リスト内容はセルの入力規則(Validation.Formula1)を実行時に読むため、
' 範囲参照(=$AO$2:$AO$3 等)・インライン("あり,なし" 等)の双方に追従する。
'
' 改修履歴: CHANGELOG.md 参照
' ========================================================

Private Const CELL_DROPDOWN_COMBO_NAME As String = "ComboBoxCellDropdown"
Private Const CELL_DROPDOWN_POLL_PROC As String = "mod_BasicInfoCellDropdown.PollCellDropdownSelection"
Private Const CELL_DROPDOWN_DEFER_COMMIT_PROC As String = "mod_BasicInfoCellDropdown.RunDeferredCommitCellDropdownSelection"
Private Const CELL_DROPDOWN_DEFER_LOSTFOCUS_PROC As String = "mod_BasicInfoCellDropdown.RunDeferredHandleCellDropdownLostFocus"
Private Const CELL_DROPDOWN_POLL_INTERVAL_DAYS As Double = 0.1 / 86400#
Private Const LOG_TAG As String = "[CellDropdown]"

Private mCellDropdownTargetAddress As String
Private mCellDropdownWorksheetName As String
Private mCellDropdownStartValue As String
Private mCellDropdownPollTime As Date
Private mLastComboListIndex As Long
Private mDropdownInteracted As Boolean
Private mInCellDropdownPrompt As Boolean

Public Function IsPromptingCellDropdown() As Boolean
    IsPromptingCellDropdown = mInCellDropdownPrompt
End Function

Public Function IsCellDropdownSessionOpen(ByVal wsInfo As Worksheet) As Boolean
    If wsInfo Is Nothing Then Exit Function
    If Len(mCellDropdownTargetAddress) > 0 Then
        IsCellDropdownSessionOpen = True
        Exit Function
    End If
    On Error Resume Next
    IsCellDropdownSessionOpen = wsInfo.OLEObjects(CELL_DROPDOWN_COMBO_NAME).Visible
    On Error GoTo 0
End Function

Private Function CellDropdownTargetAddresses() As Variant
    CellDropdownTargetAddresses = Array("C22", "C23")
End Function

Public Function IsCellDropdownTarget(ByVal wsInfo As Worksheet, ByVal target As Range) As Boolean
    If wsInfo Is Nothing Or target Is Nothing Then Exit Function

    Dim addrs As Variant
    addrs = CellDropdownTargetAddresses()

    Dim i As Long
    For i = LBound(addrs) To UBound(addrs)
        If Not Intersect(target, wsInfo.Range(CStr(addrs(i))).MergeArea) Is Nothing Then
            IsCellDropdownTarget = True
            Exit Function
        End If
    Next i
End Function

Private Function ResolveCellDropdownAnchor(ByVal wsInfo As Worksheet, ByVal target As Range) As Range
    Dim addrs As Variant
    addrs = CellDropdownTargetAddresses()

    Dim i As Long
    For i = LBound(addrs) To UBound(addrs)
        Dim anchor As Range
        Set anchor = wsInfo.Range(CStr(addrs(i)))
        If Not Intersect(target, anchor.MergeArea) Is Nothing Then
            Set ResolveCellDropdownAnchor = anchor.MergeArea.Cells(1, 1)
            Exit Function
        End If
    Next i
End Function

Private Function GetCellDropdownAnchor(ByVal wsInfo As Worksheet) As Range
    If wsInfo Is Nothing Then Exit Function
    If Len(mCellDropdownTargetAddress) = 0 Then Exit Function
    Set GetCellDropdownAnchor = wsInfo.Range(mCellDropdownTargetAddress).MergeArea.Cells(1, 1)
End Function

Public Sub ShowCellValidationDropdown(ByVal wsInfo As Worksheet, ByVal target As Range)
    If wsInfo Is Nothing Or target Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = ResolveCellDropdownAnchor(wsInfo, target)
    If anchor Is Nothing Then Exit Sub

    mod_DebugLog.Log LOG_TAG & " Show start " & anchor.Address(False, False)

    If IsCellDropdownSessionOpen(wsInfo) Then
        FinalizeOrAbandonCellDropdownSession wsInfo
    End If

    On Error GoTo CleanFail
    DeleteCellDropdownComboBox wsInfo

    Dim ole As OLEObject
    Set ole = GetCellDropdownComboBox(wsInfo, anchor)
    If ole Is Nothing Then
        mod_DebugLog.Log LOG_TAG & " GetCellDropdownComboBox failed"
        Exit Sub
    End If

    LoadComboItemsFromValidation anchor, ole
    mod_DebugLog.Log LOG_TAG & " ListCount=" & ole.Object.ListCount
    If ole.Object.ListCount = 0 Then
        mod_DebugLog.Log LOG_TAG & " empty list -> abort"
        DeleteCellDropdownComboBox wsInfo
        Exit Sub
    End If

    mCellDropdownTargetAddress = anchor.Address(False, False)
    mCellDropdownWorksheetName = wsInfo.Name
    mCellDropdownStartValue = Trim$(CStr(anchor.value))
    mInCellDropdownPrompt = True
    mDropdownInteracted = False
    FitComboToCell anchor, ole
    BindCellDropdownLinkedCell ole, anchor

    wsInfo.Activate
    anchor.Select
    ole.Visible = True
    ole.Activate
    On Error Resume Next
    ole.Object.DropDown
    If Err.Number <> 0 Then
        mod_DebugLog.Log LOG_TAG & " DropDown Err=" & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    StartCellDropdownPolling
    mod_DebugLog.Log LOG_TAG & " Show ready LinkedCell=[" & ole.LinkedCell & "] prompting=" & mInCellDropdownPrompt
    Exit Sub

CleanFail:
    mod_DebugLog.Log LOG_TAG & " Show CleanFail Err=" & Err.Number & " " & Err.Description
    ResetCellDropdownSession
    DeleteCellDropdownComboBox wsInfo
End Sub

Public Sub CompleteCellDropdownFromSheetChange(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    If Len(mCellDropdownTargetAddress) = 0 Then Exit Sub

    Dim anchor As Range
    Set anchor = GetCellDropdownAnchor(wsInfo)
    If anchor Is Nothing Then Exit Sub

    Dim currentValue As String
    currentValue = Trim$(CStr(anchor.value))
    If Len(currentValue) = 0 Then Exit Sub

    mod_DebugLog.Log LOG_TAG & " CompleteFromSheetChange value=[" & currentValue & "]"
    If TryCommitCellDropdownSelection(wsInfo) Then
        CloseCellDropdownSession wsInfo
    End If
End Sub

Public Sub CommitCellDropdownSelection(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    mDropdownInteracted = True

    If TryCommitCellDropdownSelection(wsInfo) Then
        CloseCellDropdownSession wsInfo
    End If
End Sub

Public Sub ScheduleDeferredCommitCellDropdownSelection()
    ScheduleQualifiedOnTime CELL_DROPDOWN_DEFER_COMMIT_PROC
End Sub

Public Sub RunDeferredCommitCellDropdownSelection()
    Dim wsInfo As Worksheet
    Set wsInfo = GetCellDropdownWorksheet()
    If wsInfo Is Nothing Then Exit Sub
    If Len(mCellDropdownTargetAddress) = 0 Then Exit Sub

    mDropdownInteracted = True

    Dim combo As Object
    On Error Resume Next
    Set combo = wsInfo.OLEObjects(CELL_DROPDOWN_COMBO_NAME).Object
    On Error GoTo 0
    If Not combo Is Nothing Then
        On Error Resume Next
        If combo.DroppedDown Then combo.DroppedDown = False
        On Error GoTo 0
    End If

    CommitCellDropdownSelection wsInfo
End Sub

Public Sub HandleCellDropdownLostFocus(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    If Len(mCellDropdownTargetAddress) = 0 Then Exit Sub
    ScheduleQualifiedOnTime CELL_DROPDOWN_DEFER_LOSTFOCUS_PROC
End Sub

Public Sub RunDeferredHandleCellDropdownLostFocus()
    Dim wsInfo As Worksheet
    Set wsInfo = GetCellDropdownWorksheet()
    If wsInfo Is Nothing Then Exit Sub
    If Len(mCellDropdownTargetAddress) = 0 Then Exit Sub

    If ShouldFinalizeCellDropdownSession(wsInfo) Then
        CloseCellDropdownSession wsInfo
    Else
        HideCellDropdown wsInfo
    End If
End Sub

Public Sub PollCellDropdownSelection()
    On Error GoTo CleanExit

    Dim wsInfo As Worksheet
    Set wsInfo = GetCellDropdownWorksheet()
    If wsInfo Is Nothing Then Exit Sub
    If Not IsCellDropdownSessionOpen(wsInfo) Then Exit Sub

    mod_DebugLog.Log LOG_TAG & " Poll active=" & wsInfo.Application.ActiveCell.Address(False, False) & " target=" & mCellDropdownTargetAddress

    If HasAnchorValueChanged(wsInfo) Then
        mod_DebugLog.Log LOG_TAG & " Poll detected anchor value change"
        If TryCommitCellDropdownSelection(wsInfo) Then
            CloseCellDropdownSession wsInfo
        End If
        Exit Sub
    End If

    Dim ole As OLEObject
    On Error Resume Next
    Set ole = wsInfo.OLEObjects(CELL_DROPDOWN_COMBO_NAME)
    On Error GoTo CleanExit
    If ole Is Nothing Or Not ole.Visible Then GoTo CleanExit

    Dim combo As Object
    Set combo = ole.Object

    Dim droppedDown As Boolean
    droppedDown = False
    On Error Resume Next
    droppedDown = combo.DroppedDown
    On Error GoTo CleanExit

    Dim currentIndex As Long
    currentIndex = -1
    On Error Resume Next
    currentIndex = combo.ListIndex
    On Error GoTo CleanExit

    If droppedDown Then
        mDropdownInteracted = True
        If currentIndex >= 0 Then mLastComboListIndex = currentIndex
        ScheduleCellDropdownPoll
        Exit Sub
    End If

    If mDropdownInteracted And currentIndex >= 0 Then
        If TryCommitCellDropdownSelection(wsInfo) Then
            mod_DebugLog.Log LOG_TAG & " Poll commit ListIndex=" & currentIndex
            CloseCellDropdownSession wsInfo
            Exit Sub
        End If
    End If

    If Not IsCellDropdownTarget(wsInfo, wsInfo.Application.ActiveCell) Then
        If ShouldFinalizeCellDropdownSession(wsInfo) Then
            mod_DebugLog.Log LOG_TAG & " Poll finalize on focus leave"
            CloseCellDropdownSession wsInfo
        Else
            mod_DebugLog.Log LOG_TAG & " Poll hide on focus leave without selection"
            HideCellDropdown wsInfo
        End If
        Exit Sub
    End If

    ScheduleCellDropdownPoll
    Exit Sub

CleanExit:
    mod_DebugLog.Log LOG_TAG & " Poll CleanExit Err=" & Err.Number
    CancelCellDropdownPoll
    If IsCellDropdownSessionOpen(wsInfo) Then HideCellDropdown wsInfo
End Sub

Public Sub HideCellDropdown(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    If IsCellDropdownSessionOpen(wsInfo) Then
        If ShouldFinalizeCellDropdownSession(wsInfo) Then
            CloseCellDropdownSession wsInfo
            Exit Sub
        End If
    End If

    mod_DebugLog.Log LOG_TAG & " Hide without commit"
    ResetCellDropdownSession
    DeleteCellDropdownComboBox wsInfo
End Sub

Private Function TryCommitCellDropdownSelection(ByVal wsInfo As Worksheet) As Boolean
    On Error GoTo CleanFail

    If Len(mCellDropdownTargetAddress) = 0 Then Exit Function

    Dim anchor As Range
    Set anchor = GetCellDropdownAnchor(wsInfo)
    If anchor Is Nothing Then Exit Function

    Dim combo As Object
    Set combo = Nothing
    On Error Resume Next
    Set combo = wsInfo.OLEObjects(CELL_DROPDOWN_COMBO_NAME).Object
    On Error GoTo CleanFail

    Dim selectedValue As String
    selectedValue = Trim$(CStr(anchor.value))

    If Not combo Is Nothing Then
        On Error Resume Next
        If combo.ListIndex >= 0 Then mDropdownInteracted = True
        On Error GoTo CleanFail

        Dim comboValue As String
        comboValue = ReadComboSelectedText(combo)
        If Len(comboValue) > 0 Then
            If StrComp(comboValue, selectedValue, vbBinaryCompare) <> 0 Then
                mod_DebugLog.Log LOG_TAG & " TryCommit write value=[" & comboValue & "] to " & anchor.Address(False, False)
                anchor.value = comboValue
                selectedValue = comboValue
            End If
        End If
    End If

    If Len(selectedValue) = 0 Then Exit Function

    If StrComp(selectedValue, mCellDropdownStartValue, vbBinaryCompare) <> 0 Then
        mod_DebugLog.Log LOG_TAG & " TryCommit via value change=[" & selectedValue & "]"
        TryCommitCellDropdownSelection = True
        Exit Function
    End If

    If mDropdownInteracted Then
        mod_DebugLog.Log LOG_TAG & " TryCommit confirm value=[" & selectedValue & "]"
        TryCommitCellDropdownSelection = True
        Exit Function
    End If

    Exit Function

CleanFail:
    mod_DebugLog.Log LOG_TAG & " TryCommit failed Err=" & Err.Number & " " & Err.Description
    TryCommitCellDropdownSelection = False
End Function

Private Sub NotifyCellDropdownValueCommitted(ByVal wsInfo As Worksheet, ByVal anchor As Range)
    If wsInfo Is Nothing Or anchor Is Nothing Then Exit Sub

    mod_DebugLog.Log LOG_TAG & " NotifyCommitted " & anchor.Address(False, False) & "=[" & Trim$(CStr(anchor.value)) & "]"

    On Error Resume Next
    Application.ScreenUpdating = True
    anchor.MergeArea.Cells(1, 1).Calculate
    DoEvents

    If Not Intersect(anchor, wsInfo.Range("C22,C23")) Is Nothing Then
        If Trim$(CStr(wsInfo.Range("C24").MergeArea.Cells(1, 1).value)) <> "" Then
            mod_MaterialPriceImport.SilentClearUnitPriceForBasicInfo wsInfo
        End If
        mod_Construction_Order_Import.RefreshConstructionReferenceUnitPricesOnExistingSheets
    End If

    mod_BasicInfoGuide.OnCellChanged wsInfo, anchor
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Function ShouldFinalizeCellDropdownSession(ByVal wsInfo As Worksheet) As Boolean
    ShouldFinalizeCellDropdownSession = TryCommitCellDropdownSelection(wsInfo)
    If ShouldFinalizeCellDropdownSession Then Exit Function

    Dim anchor As Range
    Set anchor = GetCellDropdownAnchor(wsInfo)
    If anchor Is Nothing Then Exit Function

    ShouldFinalizeCellDropdownSession = (Len(Trim$(CStr(anchor.value))) > 0)
End Function

Private Function HasAnchorValueChanged(ByVal wsInfo As Worksheet) As Boolean
    Dim anchor As Range
    Set anchor = GetCellDropdownAnchor(wsInfo)
    If anchor Is Nothing Then Exit Function

    Dim currentValue As String
    currentValue = Trim$(CStr(anchor.value))
    HasAnchorValueChanged = (Len(currentValue) > 0 And StrComp(currentValue, mCellDropdownStartValue, vbBinaryCompare) <> 0)
End Function

Private Function ReadComboSelectedText(ByVal combo As Object) As String
    Dim selectedValue As String

    On Error Resume Next
    If combo.ListIndex >= 0 Then
        selectedValue = Trim$(CStr(combo.List(combo.ListIndex)))
    End If
    If Len(selectedValue) = 0 Then
        selectedValue = Trim$(CStr(combo.Text))
    End If
    If Len(selectedValue) = 0 Then
        selectedValue = Trim$(CStr(combo.value))
    End If
    On Error GoTo 0

    ReadComboSelectedText = selectedValue
End Function

Private Sub CloseCellDropdownSession(ByVal wsInfo As Worksheet)
    Dim targetAddress As String
    Dim anchor As Range

    targetAddress = mCellDropdownTargetAddress
    Set anchor = GetCellDropdownAnchor(wsInfo)

    ResetCellDropdownSession
    DeleteCellDropdownComboBox wsInfo

    If Not anchor Is Nothing Then
        If Len(Trim$(CStr(anchor.value))) > 0 Then
            NotifyCellDropdownValueCommitted wsInfo, anchor
        End If
    End If

    On Error Resume Next
    If Len(targetAddress) > 0 Then wsInfo.Range(targetAddress).Select
    On Error GoTo 0
    mod_DebugLog.Log LOG_TAG & " Close session"
End Sub

Private Sub ResetCellDropdownSession()
    CancelCellDropdownPoll
    mInCellDropdownPrompt = False
    mCellDropdownTargetAddress = ""
    mCellDropdownWorksheetName = ""
    mCellDropdownStartValue = ""
    mLastComboListIndex = -2
    mDropdownInteracted = False
End Sub

Private Function GetCellDropdownWorksheet() As Worksheet
    On Error GoTo CleanFail

    If Len(mCellDropdownWorksheetName) > 0 Then
        Set GetCellDropdownWorksheet = ThisWorkbook.Worksheets(mCellDropdownWorksheetName)
        Exit Function
    End If

CleanFail:
    Set GetCellDropdownWorksheet = CommonGetBasicInfoWorksheet()
End Function

Private Sub StartCellDropdownPolling()
    mLastComboListIndex = -2
    mDropdownInteracted = False
    ScheduleCellDropdownPoll
End Sub

Private Sub FinalizeOrAbandonCellDropdownSession(ByVal wsInfo As Worksheet)
    If ShouldFinalizeCellDropdownSession(wsInfo) Then
        CloseCellDropdownSession wsInfo
    Else
        HideCellDropdown wsInfo
    End If
End Sub

Private Sub ScheduleCellDropdownPoll()
    If Len(mCellDropdownTargetAddress) = 0 Then Exit Sub

    CancelCellDropdownPoll
    mCellDropdownPollTime = Now + CELL_DROPDOWN_POLL_INTERVAL_DAYS
    ScheduleQualifiedOnTimeAt CELL_DROPDOWN_POLL_PROC, mCellDropdownPollTime
End Sub

Private Sub CancelCellDropdownPoll()
    On Error Resume Next
    If mCellDropdownPollTime <> 0 Then
        Application.OnTime EarliestTime:=mCellDropdownPollTime, _
                           Procedure:=QualifiedOnTimeProcedure(CELL_DROPDOWN_POLL_PROC), _
                           Schedule:=False
    End If
    mCellDropdownPollTime = 0
    On Error GoTo 0
End Sub

Private Sub ScheduleQualifiedOnTime(ByVal moduleProcedure As String)
    ScheduleQualifiedOnTimeAt moduleProcedure, Now + TimeValue("00:00:00")
End Sub

Private Sub ScheduleQualifiedOnTimeAt(ByVal moduleProcedure As String, ByVal runTime As Date)
    On Error Resume Next
    Application.OnTime EarliestTime:=runTime, Procedure:=QualifiedOnTimeProcedure(moduleProcedure)
    If Err.Number <> 0 Then
        mod_DebugLog.Log LOG_TAG & " OnTime schedule failed proc=[" & moduleProcedure & "] Err=" & Err.Number
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function QualifiedOnTimeProcedure(ByVal moduleProcedure As String) As String
    QualifiedOnTimeProcedure = "'" & ThisWorkbook.Name & "'!" & moduleProcedure
End Function

Private Sub BindCellDropdownLinkedCell(ByVal ole As OLEObject, ByVal anchor As Range)
    ClearCellDropdownLinkedCell ole
    On Error Resume Next
    ole.LinkedCell = anchor.Address(False, False)
    If Err.Number <> 0 Then
        mod_DebugLog.Log LOG_TAG & " LinkedCell bind failed Err=" & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub ClearCellDropdownLinkedCell(ByVal ole As OLEObject)
    On Error Resume Next
    ole.LinkedCell = ""
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GetCellDropdownComboBox(ByVal wsInfo As Worksheet, ByVal anchor As Range) As OLEObject
    On Error Resume Next
    Set GetCellDropdownComboBox = wsInfo.OLEObjects(CELL_DROPDOWN_COMBO_NAME)
    On Error GoTo 0
    If Not GetCellDropdownComboBox Is Nothing Then Exit Function

    On Error Resume Next
    Set GetCellDropdownComboBox = wsInfo.OLEObjects.Add(ClassType:="Forms.ComboBox.1", _
                                                        Link:=False, _
                                                        DisplayAsIcon:=False, _
                                                        Left:=anchor.MergeArea.Left, _
                                                        Top:=anchor.MergeArea.Top, _
                                                        Width:=anchor.MergeArea.Width, _
                                                        Height:=anchor.MergeArea.Height)
    If Not GetCellDropdownComboBox Is Nothing Then
        GetCellDropdownComboBox.Name = CELL_DROPDOWN_COMBO_NAME
        GetCellDropdownComboBox.Visible = False
    End If
    On Error GoTo 0
End Function

Private Sub FitComboToCell(ByVal anchor As Range, ByVal ole As OLEObject)
    On Error Resume Next
    Dim area As Range
    Set area = anchor.MergeArea
    With ole
        .Left = area.Left
        .Top = area.Top
        .Width = area.Width
        .Height = area.Height
        .Placement = xlMove
    End With
    On Error GoTo 0
End Sub

Private Sub LoadComboItemsFromValidation(ByVal anchor As Range, ByVal ole As OLEObject)
    On Error Resume Next

    Dim f1 As String
    f1 = ""
    If anchor.Validation.Type = xlValidateList Then f1 = anchor.Validation.Formula1

    With ole.Object
        .Clear

        If Len(f1) > 0 Then
            If Left$(f1, 1) = "=" Then
                Dim listRange As Range
                Set listRange = ResolveListRangeFromValidation(anchor, f1)
                If Not listRange Is Nothing Then
                    Dim c As Range
                    For Each c In listRange.Cells
                        If Len(Trim$(CStr(c.value))) > 0 Then .AddItem CStr(c.value)
                    Next c
                End If
            Else
                Dim parts As Variant
                parts = Split(f1, ",")
                Dim i As Long
                For i = LBound(parts) To UBound(parts)
                    If Len(Trim$(CStr(parts(i)))) > 0 Then .AddItem Trim$(CStr(parts(i)))
                Next i
            End If
        End If

        .Style = fmStyleDropDownList
        .ListRows = Application.Max(1, Application.Min(12, .ListCount))
        .MatchRequired = True
        .ListIndex = -1
        If Len(Trim$(CStr(anchor.value))) > 0 Then
            .value = CStr(anchor.value)
        End If
    End With

    On Error GoTo 0
End Sub

Private Function ResolveListRangeFromValidation(ByVal anchor As Range, ByVal formula1 As String) As Range
    Dim formulaBody As String
    formulaBody = Mid$(formula1, 2)

    On Error Resume Next
    Set ResolveListRangeFromValidation = anchor.Worksheet.Range(formulaBody)
    If ResolveListRangeFromValidation Is Nothing Then
        Set ResolveListRangeFromValidation = Evaluate(formula1)
    End If
    If ResolveListRangeFromValidation Is Nothing Then
        Set ResolveListRangeFromValidation = anchor.Worksheet.Parent.Names(formulaBody).RefersToRange
    End If
    On Error GoTo 0
End Function

Private Sub DeleteCellDropdownComboBox(ByVal wsInfo As Worksheet)
    On Error Resume Next
    Dim ole As OLEObject
    Set ole = wsInfo.OLEObjects(CELL_DROPDOWN_COMBO_NAME)
    If Not ole Is Nothing Then
        ClearCellDropdownLinkedCell ole
        ole.Visible = False
        ole.Delete
    End If
    On Error GoTo 0
End Sub
