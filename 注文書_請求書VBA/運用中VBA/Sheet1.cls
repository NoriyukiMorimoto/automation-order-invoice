Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo ExitHandler
    Application.EnableEvents = False

    mod_FillManagerName.HideOfficeComboBox Me
    mod_FillManagerName.RefreshBranchOfficeValidation True
    mod_VendorMaster.RefreshVendorListForBasicInfo Me
    mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation Me

ExitHandler:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    mod_DebugLog.Log "[Sheet1] Worksheet_Change 発火: " & Target.Address & " Value=[" & CStr(Target.Value) & "] EnableEvents=" & Application.EnableEvents
    If Intersect(Target, Me.Range("B4,B6,C6,C10,C15:C16,F11,C20,C22,C23,C24,F2")) Is Nothing Then
        mod_DebugLog.Log "[Sheet1] 監視対象外 -> Exit"
        Exit Sub
    End If

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    Dim needComboPrompt As Boolean
    needComboPrompt = False

    If Not Intersect(Target, Me.Range("B4,B6")) Is Nothing Then
        mod_DebugLog.Log "[Sheet1] B4/B6変更 -> RefreshBranchOfficeValidation(False) 呼び出し"
        needComboPrompt = mod_FillManagerName.RefreshBranchOfficeValidation(False)
        mod_DebugLog.Log "[Sheet1] RefreshBranchOfficeValidation 戻り値 needComboPrompt=" & needComboPrompt
    End If

    If Not Intersect(Target, Me.Range("B6")) Is Nothing Then
        mod_VendorMaster.RefreshVendorListForBasicInfo Me
    End If

    ' B6 または C6 が変更された場合：確認なしで基本情報・単価シートをクリア
    If Not Intersect(Target, Me.Range("B6,C6")) Is Nothing Then
        mod_DebugLog.Log "[Sheet1] B6/C6変更 IsSuppressing=" & mod_FillManagerName.IsSuppressingC6Change()
        If Not mod_FillManagerName.IsSuppressingC6Change() Then
            mod_DebugLog.Log "[Sheet1] SilentClearBasicInfo 呼び出し"
            mod_BasicInfoUpdate.SilentClearBasicInfo Me
            mod_DebugLog.Log "[Sheet1] SilentClearBasicInfo 完了"
        End If
    End If

    If Not Intersect(Target, Me.Range("C6")) Is Nothing Then
        If Not mod_FillManagerName.IsSuppressingC6Change() Then
            mod_FillManagerName.FillManagerNameToBasicInfo
        End If
    End If

    If Not Intersect(Target, Me.Range("C10")) Is Nothing Then
        mod_MaterialPriceImport.AutoFillLineTypeFromWorkName Me
        mod_MaterialPriceImport.AutoFillProjectNameFromWorkName Me
    End If

    If Not Intersect(Target, Me.Range("F2")) Is Nothing Then
        mod_BasicInfoUpdate.ApplyInitialBillingPeriodFromStartDate Me
    End If

    If Not Intersect(Target, Me.Range("C15:C16")) Is Nothing Then
        mod_BasicInfoUpdate.ApplyWorkDaysFromWorkDates Me
    End If

    If Not Intersect(Target, Me.Range("C20")) Is Nothing Then
        mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation Me, False
    End If

    If Not Intersect(Target, Me.Range("C24")) Is Nothing Then
        mod_MaterialPriceImport.FormatImportedLineNamesCell Me
    End If

    If Not Intersect(Target, Me.Range("C22,C23")) Is Nothing Then
        If Trim$(CStr(Me.Range("C24").Value)) <> "" Then
            mod_MaterialPriceImport.SilentClearUnitPriceForBasicInfo Me
        End If
    End If

    If Not Intersect(Target, Me.Range("F11")) Is Nothing Then
        mod_VendorMaster.FillVendorInfoToBasicInfo Me
        mod_VendorMaster.CancelScheduledVendorListDropdown
        mod_VendorMaster.HideVendorComboBox Me
    End If

    mod_DebugLog.Log "[Sheet1] needComboPrompt=" & needComboPrompt
    If needComboPrompt Then
        mod_DebugLog.Log "[Sheet1] PromptOfficeComboBox 呼び出し前 EnableEvents=" & Application.EnableEvents
        mod_FillManagerName.PromptOfficeComboBox
        mod_DebugLog.Log "[Sheet1] PromptOfficeComboBox 呼び出し後 EnableEvents=" & Application.EnableEvents
    End If

ExitHandler:
    mod_DebugLog.Log "[Sheet1] Worksheet_Change 終了 EnableEvents=" & Application.EnableEvents
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Me.Range("C24")) Is Nothing Then
        Cancel = True
        mod_MaterialPriceImport.ImportConstructionUnitPriceForBasicInfo
        Exit Sub
    End If

    If Not Intersect(Target, Me.Range("F11")) Is Nothing Then
        Cancel = True
        mod_VendorMaster.CancelScheduledVendorListDropdown
        mod_VendorMaster.HideVendorComboBox Me
        mod_VendorMaster.ShowAllVendorSelection Me
        Exit Sub
    End If

    mod_BasicInfoCalendar.HandleBasicInfoProjectNameDoubleClick Target, Cancel
    If Cancel Then Exit Sub

    mod_BasicInfoCalendar.HandleBasicInfoDateDoubleClick Target, Cancel
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    mod_DebugLog.Log "[Sheet1] SelectionChange: " & Target.Address & " IsPrompting=" & mod_FillManagerName.IsPromptingOfficeComboBox()
    mod_BasicInfoCalendar.HandleBasicInfoSelectionChange Target

    ' (#22) PromptOfficeComboBox 処理中はコンボを消さない
    If Not mod_FillManagerName.IsPromptingOfficeComboBox() Then
        If Intersect(Target, Me.Range("C6")) Is Nothing Then
            mod_DebugLog.Log "[Sheet1] SelectionChange: C6以外 -> HideOfficeComboBox"
            mod_FillManagerName.HideOfficeComboBox Me
        End If
    Else
        mod_DebugLog.Log "[Sheet1] SelectionChange: IsPrompting=True -> HideスキップII"
    End If

    If Not Intersect(Target, Me.Range("C21")) Is Nothing Then
        mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation Me, True
    End If

    If Not Intersect(Target, Me.Range("F11")) Is Nothing Then
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
    ' (#22) PromptOfficeComboBox 処理中の LostFocus は無視する
    '       （コンボ表示直後に誤って確定処理が走るのを防ぐ）
    If mod_FillManagerName.IsPromptingOfficeComboBox() Then
        mod_DebugLog.Log "[Sheet1] ComboBox1_LostFocus: IsPrompting=True -> スキップ"
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
