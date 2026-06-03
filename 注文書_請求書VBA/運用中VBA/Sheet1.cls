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
    If Intersect(Target, Me.Range("B4,B6,C6,C10,C15:C16,F11,C20,C22,C23,C24,F2")) Is Nothing Then Exit Sub

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    Dim needComboPrompt As Boolean
    needComboPrompt = False

    If Not Intersect(Target, Me.Range("B4,B6")) Is Nothing Then
        ' (#19) 戻り値でコンボ表示スケジュールの要否を受け取る
        needComboPrompt = mod_FillManagerName.RefreshBranchOfficeValidation(False)
    End If

    If Not Intersect(Target, Me.Range("B6")) Is Nothing Then
        mod_VendorMaster.RefreshVendorListForBasicInfo Me
    End If

    ' B6 または C6 が変更された場合：確認なしで基本情報・単価シートをクリア
    If Not Intersect(Target, Me.Range("B6,C6")) Is Nothing Then
        If Not mod_FillManagerName.IsSuppressingC6Change() Then
            mod_BasicInfoUpdate.SilentClearBasicInfo Me
        End If
    End If

    If Not Intersect(Target, Me.Range("C6")) Is Nothing Then
        ' ComboBox選択確定による書き込み中はChangeイベントをスキップ
        If Not mod_FillManagerName.IsSuppressingC6Change() Then
            mod_FillManagerName.FillManagerNameToBasicInfo
        End If
    End If

    ' C10 が変更された場合：在幹区分マスタを参照して C20・C21 を自動入力
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

    ' C22 または C23 が変更された場合：C24 に内容があれば（単価シート取込済み）、
    ' 確認なしで C24 ・単価シート・購入充当単価シート・レール溶接単価シートを削除する
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

ExitHandler:
    Application.EnableEvents = True
    ' (#19) EnableEvents=True 復帰後にコンボ表示をスケジュール
    ' （EnableEvents=False 中に OnTime 予約すると SilentClearBasicInfo との
    '   競合でコンボが表示される前に非表示にされる問題を回避）
    If needComboPrompt Then mod_FillManagerName.ScheduleOfficeComboBoxPrompt
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' C24 ダブルクリック → 積算線区単価取込
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
    mod_BasicInfoCalendar.HandleBasicInfoSelectionChange Target

    If Intersect(Target, Me.Range("C6")) Is Nothing Then
        mod_FillManagerName.HideOfficeComboBox Me
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
