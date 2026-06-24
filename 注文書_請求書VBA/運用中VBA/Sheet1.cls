Option Explicit

Private Sub Worksheet_Activate()
    Dim prevCalc As XlCalculation, prevScreen As Boolean
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation

    On Error GoTo ExitHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    mod_VendorMaster.ClearVendorRowsCache
    mod_FillManagerName.HideOfficeComboBox Me
    mod_FillManagerName.RefreshBranchOfficeValidation True
    mod_VendorMaster.InitVendorBlockCountFromSheet Me
    mod_VendorMaster.RefreshVendorListForBasicInfo Me
    mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation Me
    mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals
    mod_BasicInfoGuide.InitBasicInfoGuide Me

ExitHandler:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    Dim vendorNameRange As Range
    Set vendorNameRange = mod_VendorMaster.GetVendorNameRange(Me)

    Dim vendorUnitPriceMonitorRange As Range
    Set vendorUnitPriceMonitorRange = mod_VendorMaster.GetVendorUnitPriceMonitorRange(Me)

    If Intersect(target, Me.Range("B4,B6,C6,C10,C15:C16,F9,C20,C22,C23,C24,F2")) Is Nothing And _
       Intersect(target, vendorNameRange) Is Nothing And _
       Intersect(target, vendorUnitPriceMonitorRange) Is Nothing And _
       Intersect(target, Me.Range("C9,C22,C23,F9,F10,F11,F27,F29,F31,I10,I11,I27,I29,I31,L10,L11,L27,L29,L31,O10,O11,O27,O29,O31,R10,R11,R27,R29,R31,U10,U11,U27,U29,U31,X10,X11,X27,X29,X31,AA10,AA11,AA27,AA29,AA31,AD10,AD11,AD27,AD29,AD31,AG10,AG11,AG27,AG29,AG31")) Is Nothing Then
        Exit Sub
    End If

    If mod_VendorMaster.IsSyncVendorBlocksInProgress() Then Exit Sub

    mod_DebugLog.Log "[Sheet1] Worksheet_Change fired: " & target.Address & " Value=[" & GetTargetValueText(target) & "] EnableEvents=" & Application.EnableEvents

    Dim prevCalc As XlCalculation, prevScreen As Boolean
    prevScreen = Application.ScreenUpdating
    prevCalc = Application.Calculation

    On Error GoTo ExitHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim needComboPrompt As Boolean
    needComboPrompt = False

    If Not Intersect(target, Me.Range("B4,B6")) Is Nothing Then
        mod_DebugLog.Log "[Sheet1] B4/B6changed -> RefreshBranchOfficeValidation(False) call"
        needComboPrompt = mod_FillManagerName.RefreshBranchOfficeValidation(False)
        mod_DebugLog.Log "[Sheet1] RefreshBranchOfficeValidation result needComboPrompt=" & needComboPrompt
    End If

    If Not Intersect(target, Me.Range("B6")) Is Nothing Then
        mod_VendorMaster.RefreshVendorListForBasicInfo Me
    End If

    If Not Intersect(target, Me.Range("F9")) Is Nothing Then
        mod_VendorMaster.SyncVendorBlocksFromCount Me
        Set vendorNameRange = mod_VendorMaster.GetVendorNameRange(Me)
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals
    End If

    If Not Intersect(target, Me.Range("B6,C6")) Is Nothing Then
        mod_DebugLog.Log "[Sheet1] B6/C6changed IsSuppressing=" & mod_FillManagerName.IsSuppressingC6Change()
        If Not mod_FillManagerName.IsSuppressingC6Change() Then
            mod_DebugLog.Log "[Sheet1] SilentClearBasicInfo call"
            mod_BasicInfoUpdate.SilentClearBasicInfo Me
            mod_DebugLog.Log "[Sheet1] SilentClearBasicInfo done"
        End If
    End If

    If Not Intersect(target, Me.Range("C6")) Is Nothing Then
        If Not mod_FillManagerName.IsSuppressingC6Change() Then
            mod_FillManagerName.FillManagerNameToBasicInfo
        End If
    End If

    If Not Intersect(target, Me.Range("C10")) Is Nothing Then
        mod_MaterialPriceImport.AutoFillUnitPriceFieldsFromWorkName Me
    End If

    If Not Intersect(target, Me.Range("F2")) Is Nothing Then
        mod_BasicInfoUpdate.ApplyInitialBillingPeriodFromStartDate Me
    End If

    If Not Intersect(target, Me.Range("C15:C16")) Is Nothing Then
        mod_BasicInfoUpdate.ApplyWorkDaysFromWorkDates Me
    End If

    If Not Intersect(target, Me.Range("C20")) Is Nothing Then
        mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation Me, False
    End If

    If Not Intersect(target, Me.Range("C24")) Is Nothing Then
        mod_MaterialPriceImport.HandleImportedLineNamesCellChange Me, target
        If Len(Trim$(CStr(Me.Range("C24").MergeArea.Cells(1, 1).value))) > 0 Then
            mod_MaterialPriceImport.FormatImportedLineNamesCell Me
        End If
        If Not mod_MaterialPriceImport.IsImportingUnitPriceData() Then
            mod_Construction_Order_Import.RefreshConstructionReferenceUnitPricesOnExistingSheets
        End If
    End If

    If Not Intersect(target, Me.Range("C22,C23")) Is Nothing Then
        If Trim$(CStr(Me.Range("C24").value)) <> "" Then
            mod_MaterialPriceImport.SilentClearUnitPriceForBasicInfo Me
        End If
        mod_Construction_Order_Import.RefreshConstructionReferenceUnitPricesOnExistingSheets
    End If

    If Not Intersect(target, vendorNameRange) Is Nothing Then
        Dim changedVendorCell As Range
        Set changedVendorCell = Intersect(target, vendorNameRange).Cells(1, 1)
        mod_VendorMaster.FillVendorInfoToBasicInfo Me, changedVendorCell
        mod_VendorMaster.CancelScheduledVendorListDropdown
        mod_VendorMaster.HideVendorComboBox Me
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals _
            mod_VendorMaster.GetVendorIndexFromValueColumnPublic(changedVendorCell.Column)
    End If

    Dim monitorHit As Range
    Set monitorHit = Intersect(target, vendorUnitPriceMonitorRange)
    If Not monitorHit Is Nothing Then
        Dim nameHit As Range
        Set nameHit = Intersect(monitorHit, vendorNameRange)
        If nameHit Is Nothing Then
            mod_VendorMaster.HandleVendorUnitPriceMonitorChange Me, target
        ElseIf monitorHit.Cells.CountLarge > nameHit.Cells.CountLarge Then
            mod_VendorMaster.HandleVendorUnitPriceMonitorChange Me, target
        End If
    End If

    ' g????K?C?h?igh?F+?R???g?g?j?X?V
    mod_BasicInfoGuide.OnCellChanged Me, target

    mod_DebugLog.Log "[Sheet1] needComboPrompt=" & needComboPrompt
    If needComboPrompt Then
        mod_DebugLog.Log "[Sheet1] PromptOfficeComboBox before-call EnableEvents=" & Application.EnableEvents
        mod_FillManagerName.PromptOfficeComboBox
        mod_DebugLog.Log "[Sheet1] PromptOfficeComboBox after-call EnableEvents=" & Application.EnableEvents
    End If

ExitHandler:
    mod_DebugLog.Log "[Sheet1] Worksheet_Change done EnableEvents=" & Application.EnableEvents
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    If IsUnitPriceImportDoubleClickTarget(target) Then
        Cancel = True
        mod_MaterialPriceImport.ImportConstructionUnitPriceForBasicInfo
        Exit Sub
    End If

    Dim vendorNameRange As Range
    Set vendorNameRange = mod_VendorMaster.GetVendorNameRange(Me)

    If Not Intersect(target, vendorNameRange) Is Nothing Then
        Cancel = True
        mod_VendorMaster.CancelScheduledVendorListDropdown
        mod_VendorMaster.HideVendorComboBox Me
        mod_VendorMaster.ShowAllVendorSelection Me, Intersect(target, vendorNameRange).Cells(1, 1)
        Exit Sub
    End If

    mod_BasicInfoCalendar.HandleBasicInfoProjectNameDoubleClick target, Cancel
    If Cancel Then Exit Sub

    mod_BasicInfoCalendar.HandleBasicInfoDateDoubleClick target, Cancel
End Sub

Private Function IsUnitPriceImportDoubleClickTarget(ByVal target As Range) As Boolean
    If target Is Nothing Then Exit Function

    Dim targetArea As Range
    On Error Resume Next
    Set targetArea = target.MergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    IsUnitPriceImportDoubleClickTarget = _
        (targetArea.Cells(1, 1).Address(False, False) = Me.Range("C24").Address(False, False) And _
         targetArea.Column = Me.Range("C24").Column And _
         targetArea.Columns.Count = 1)
End Function

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    mod_DebugLog.Log "[Sheet1] SelectionChange: " & target.Address & " IsPrompting=" & mod_FillManagerName.IsPromptingOfficeComboBox()
    mod_BasicInfoCalendar.HandleBasicInfoSelectionChange target

    If Not mod_FillManagerName.IsPromptingOfficeComboBox() Then
        If Intersect(target, Me.Range("C6")) Is Nothing Then
            mod_DebugLog.Log "[Sheet1] SelectionChange: C6-other -> HideOfficeComboBox"
            mod_FillManagerName.HideOfficeComboBox Me
        End If
    Else
        mod_DebugLog.Log "[Sheet1] SelectionChange: IsPrompting=True -> Hide-skip"
    End If

    If Not Intersect(target, Me.Range("C21")) Is Nothing Then
        mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation Me, True
    End If

    Dim vendorNameRange As Range
    Set vendorNameRange = mod_VendorMaster.GetVendorNameRange(Me)

    If Not Intersect(target, vendorNameRange) Is Nothing Then
        mod_VendorMaster.RefreshVendorListForBasicInfo Me
        mod_VendorMaster.ScheduleVendorListDropdown Me
    Else
        mod_VendorMaster.CancelScheduledVendorListDropdown
        mod_VendorMaster.HideVendorComboBox Me
    End If
End Sub

Private Sub ComboBox1_Click()
    mod_FillManagerName.CommitOfficeComboBoxSelection
End Sub

Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mod_FillManagerName.CommitOfficeComboBoxSelection
    End If
End Sub

Private Sub ComboBox1_LostFocus()
    mod_DebugLog.Log "[Sheet1] ComboBox1_LostFocus IsPrompting=" & mod_FillManagerName.IsPromptingOfficeComboBox()
    If mod_FillManagerName.IsPromptingOfficeComboBox() Then
        mod_DebugLog.Log "[Sheet1] ComboBox1_LostFocus: IsPrompting=True -> skip"
        Exit Sub
    End If
    mod_FillManagerName.CommitOfficeComboBoxSelection False
End Sub

Private Sub ComboBoxVendor_Click()
    mod_VendorMaster.CommitVendorComboBoxSelection Me
End Sub

Private Sub ComboBoxVendor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mod_VendorMaster.CommitVendorComboBoxSelection Me
    End If
End Sub

Private Sub btnUpdateBilling_Click()
    mod_BasicInfoUpdate.UpdateBasicInfoPeriod
End Sub

Private Sub btnClearBasicInfo_Click()
    mod_BasicInfoUpdate.ClearBasicInfo
End Sub

Private Function GetTargetValueText(ByVal target As Range) As String
    On Error GoTo ErrorHandler

    If target Is Nothing Then
        GetTargetValueText = ""
        Exit Function
    End If

    If target.Cells.CountLarge > 1 Then
        GetTargetValueText = "(multi-cell)"
        Exit Function
    End If

    Dim cellValue As Variant
    cellValue = target.value

    If IsError(cellValue) Then
        GetTargetValueText = "(error)"
    ElseIf IsArray(cellValue) Then
        GetTargetValueText = "(array)"
    Else
        GetTargetValueText = CStr(cellValue)
    End If
    Exit Function

ErrorHandler:
    GetTargetValueText = "(unreadable)"
End Function
