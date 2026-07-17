Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "btnUpdateBilling, 1, 0, MSForms, CommandButton"
Attribute VB_Control = "btnClearBasicInfo, 2, 1, MSForms, CommandButton"
Attribute VB_Control = "ComboBox1, 3, 2, MSForms, ComboBox"
Option Explicit

' ��{���V�[�g��Activate���A�d���č\�z(���͋K���̍č\�z�E���v�Čv�Z�E���K�V�[�|��)��
' �Z�b�V��������̂ݎ��s���邽�߂̃t���O�B2��ڈȍ~�̐ؑւł͏ȗ�����(�e�����͑Ή�����
' �Z���ύX�C�x���g/�捞�������ňێ�����邽�߁A�ؑւ̓s�x��蒼���K�v���Ȃ�)�B
Private mBasicInfoActivateInitialized As Boolean

' �ύX�Ď��Q�[�g�͈͂̐ÓI��(�Œ�Z��+�{�H��у u���b�N��)�̃L���b�V���B
' ���C�A�E�g�͐ÓI�Ȃ��߈�x�\�z����΍ė��p�ł��A���ύX�ł�60����Union�č\�z�����B
' �����Ώۃ Z���W���͏]���Ɠ���(�Œ�Z��+�{�H��у u���b�N10��6�s)�B�O�̂���Activate�Ŗ��������B
Private mStaticGateRangeCache As Range

Private Sub Worksheet_Activate()
    Dim g As New clsPerfGuard
    g.Suspend

    On Error GoTo ExitHandler

    ' ����K�v�Ȍy�ʏ���(�L���b�V���������EUI��n���E�����������E���̓K�C�h)
    mod_VendorMaster.ClearVendorRowsCache
    Set mStaticGateRangeCache = Nothing   ' �ύX�Q�[�g�͈͂̐ÓI���L���b�V�����O�̂��ߍč\�z
    mod_FillManagerName.HideOfficeComboBox Me
    mod_VendorMaster.EnsureVendorCountInputValidation Me
    mod_OrderTpl_Shared.OrderTplRepairAllGeneratedPlaceholderFormulas
    mod_VendorMaster.InitVendorBlockCountFromSheet Me
    mod_BasicInfoGuide.InitBasicInfoGuide Me

    ' �d���č\�z(���K�V�[�|���E���͋K���̍č\�z�E���v�Čv�Z)�̓Z�b�V��������̂ݎ��s����B
    ' 2��ڈȍ~�̐ؑւł͏ȗ�(��������Ή�����Z���ύX�C�x���g/�捞�������ōŐV�Ɉێ������)�B
    If Not mBasicInfoActivateInitialized Then
        mod_FillManagerName.CleanupLegacyBasicInfoValidationListDebris Me
        mod_VendorMaster.CleanupLegacyVendorListDebrisInColumnAD Me
        mod_FillManagerName.RefreshBranchOfficeValidation True
        mod_VendorMaster.RefreshVendorListForBasicInfo Me
        mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation Me
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals
        mBasicInfoActivateInitialized = True
    End If

ExitHandler:
    g.Restore
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    mod_DebugLog.Log "[Sheet1] Worksheet_Change: " & target.Address & " Value=[" & GetTargetValueText(target) & "] EnableEvents=" & Application.EnableEvents

    If Not Intersect(target, Me.Range("C22,C23")) Is Nothing Then
        If mod_BasicInfoCellDropdown.IsCellDropdownSessionOpen(Me) Then
            mod_BasicInfoCellDropdown.CompleteCellDropdownFromSheetChange Me
        End If
    End If

    Dim vendorNameRange As Range
    Set vendorNameRange = mod_VendorMaster.GetVendorNameRange(Me)

    Dim vendorUnitPriceMonitorRange As Range
    Set vendorUnitPriceMonitorRange = mod_VendorMaster.GetVendorUnitPriceMonitorRange(Me)

    If IsIgnoredBasicInfoChange(Me, target, vendorNameRange, vendorUnitPriceMonitorRange) Then
        mod_DebugLog.Log "[Sheet1] �Ď��ΏۊO -> Exit"
        Exit Sub
    End If

    If mod_VendorMaster.IsSyncVendorBlocksInProgress() Then
        If Not Intersect(target, Me.Range("F9")) Is Nothing Then
            mod_VendorMaster.ResetVendorBlockSyncState
        Else
            Exit Sub
        End If
    End If

    Dim g As New clsPerfGuard
    g.Suspend

    On Error GoTo ExitHandler

    Dim needComboPrompt As Boolean
    needComboPrompt = False

    If Not Intersect(target, Me.Range("B4,B6")) Is Nothing Then
        mod_DebugLog.Log "[Sheet1] B4/B6�ύX -> RefreshBranchOfficeValidation(False) �Ăяo��"
        needComboPrompt = mod_FillManagerName.RefreshBranchOfficeValidation(False)
        mod_DebugLog.Log "[Sheet1] RefreshBranchOfficeValidation �߂�l needComboPrompt=" & needComboPrompt
    End If

    If Not Intersect(target, Me.Range("B6")) Is Nothing Then
        mod_VendorMaster.RefreshVendorListForBasicInfo Me
    End If

    If Not Intersect(target, Me.Range("F9")) Is Nothing Then
        mod_VendorMaster.SyncVendorBlocksFromCount Me
        Set vendorNameRange = mod_VendorMaster.GetVendorNameRange(Me)
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals
        ' ��А������őΏۊO�ɂȂ����������e���v���[�g�V�[�g���폜
        mod_OrderTpl_Generate.RemoveOrphanGeneratedSheets Me
    End If

    If Not Intersect(target, Me.Range("B6,C6")) Is Nothing Then
        mod_DebugLog.Log "[Sheet1] B6/C6�ύX IsSuppressing=" & mod_FillManagerName.IsSuppressingC6Change()
        If Not mod_FillManagerName.IsSuppressingC6Change() Then
            mod_DebugLog.Log "[Sheet1] SilentClearBasicInfo �Ăяo��"
            mod_BasicInfoUpdate.SilentClearBasicInfo Me
            mod_DebugLog.Log "[Sheet1] SilentClearBasicInfo ����"
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
        If Len(Trim$(CStr(Me.Range("C24").mergeArea.Cells(1, 1).value))) > 0 Then
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

    Dim changedVendorCell As Range
    Set changedVendorCell = mod_VendorMaster.ResolveVendorNameChangeCell(Me, target)
    If changedVendorCell Is Nothing And Not vendorNameRange Is Nothing Then
        If Not Intersect(target, vendorNameRange) Is Nothing Then
            Set changedVendorCell = Intersect(target, vendorNameRange).Cells(1, 1)
        End If
    End If
    If Not changedVendorCell Is Nothing Then
        mod_VendorMaster.FillVendorInfoToBasicInfo Me, changedVendorCell
        mod_VendorMaster.CancelScheduledVendorListDropdown
        mod_VendorMaster.HideVendorComboBox Me
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals _
            mod_VendorMaster.GetVendorIndexFromValueColumnPublic(changedVendorCell.Column)
        ' �{�H��Њm��ɔ����������e���v���[�g�V�[�g�̐���/�폜
        mod_OrderTpl_Generate.HandleVendorNameCellChange Me, changedVendorCell
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

    mod_BasicInfoGuide.OnCellChanged Me, target

    ' ��{���̃w�b�_�[�]�L���Z���ύX�𒍕����e���v���[�g�e�V�[�g�֔��f
    ' (10�s�ڂ̍H���敪�F�̓K�p��Ƀ^�u�F��ǂݎ�邽�߁A�K�C�h�X�V�̌�Ɏ��s����)
    mod_OrderTpl_Header.HandleBasicInfoHeaderSourceChange Me, target

    mod_DebugLog.Log "[Sheet1] needComboPrompt=" & needComboPrompt
    If needComboPrompt Then
        mod_DebugLog.Log "[Sheet1] PromptOfficeComboBox �Ăяo���O EnableEvents=" & Application.EnableEvents
        mod_FillManagerName.PromptOfficeComboBox
        mod_DebugLog.Log "[Sheet1] PromptOfficeComboBox �Ăяo���� EnableEvents=" & Application.EnableEvents
    End If

ExitHandler:
    mod_DebugLog.Log "[Sheet1] Worksheet_Change done EnableEvents=" & Application.EnableEvents
    g.Restore
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    If IsUnitPriceImportDoubleClickTarget(target) Then
        Cancel = True
        mod_MaterialPriceImport.ImportConstructionUnitPriceForBasicInfo
        Exit Sub
    End If

    If mod_BasicInfoCellDropdown.IsCellDropdownTarget(Me, target) Then
        Cancel = True
        mod_BasicInfoCellDropdown.ShowCellValidationDropdown Me, target
        Exit Sub
    End If

    If Not Intersect(target, Me.Range("C13")) Is Nothing Then
        Cancel = True
        mod_PrefectureSelector.RequestPrefectureSelection
        Exit Sub
    End If

    If mod_BasicInfoOrderNumberKeypad.IsOrderNumberKeypadTarget(Me, target) Then
        Cancel = True
        mod_BasicInfoOrderNumberKeypad.RequestOrderNumberKeypad Me, target
        Exit Sub
    End If

    If mod_BasicInfoExclusiveChoice.IsPartialPaymentTarget(Me, target) Then
        Cancel = True
        mod_BasicInfoExclusiveChoice.RequestPartialPaymentSelection Me, target
        Exit Sub
    End If

    If mod_BasicInfoExclusiveChoice.IsLaborInsuranceTarget(Me, target) Then
        Cancel = True
        mod_BasicInfoExclusiveChoice.RequestLaborInsuranceSelection Me, target
        Exit Sub
    End If

    If mod_BasicInfoSupplyLoan.IsSupplyMaterialTarget(Me, target) Then
        Cancel = True
        mod_BasicInfoSupplyLoan.RequestSupplyMaterialSelection Me, target
        Exit Sub
    End If

    If mod_BasicInfoSupplyLoan.IsLoanItemTarget(Me, target) Then
        Cancel = True
        mod_BasicInfoSupplyLoan.RequestLoanItemSelection Me, target
        Exit Sub
    End If

    If mod_BasicInfoExclusiveChoice.IsRecycleLawTarget(Me, target) Then
        Cancel = True
        mod_BasicInfoExclusiveChoice.RequestRecycleLawSelection Me, target
        Exit Sub
    End If

    Dim vendorNameCell As Range
    Set vendorNameCell = mod_VendorMaster.ResolveVendorNameChangeCell(Me, target)
    If vendorNameCell Is Nothing Then
        Dim vendorNameRange As Range
        Set vendorNameRange = mod_VendorMaster.GetVendorNameRange(Me)
        If Not vendorNameRange Is Nothing Then
            If Not Intersect(target, vendorNameRange) Is Nothing Then
                Set vendorNameCell = Intersect(target, vendorNameRange).Cells(1, 1)
            End If
        End If
    End If

    If Not vendorNameCell Is Nothing Then
        Cancel = True
        mod_VendorMaster.CancelScheduledVendorListDropdown
        mod_VendorMaster.HideVendorComboBox Me
        mod_VendorMaster.ShowAllVendorSelection Me, vendorNameCell
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
    Set targetArea = target.mergeArea
    On Error GoTo 0
    If targetArea Is Nothing Then Set targetArea = target

    IsUnitPriceImportDoubleClickTarget = _
        (targetArea.Cells(1, 1).Address(False, False) = Me.Range("C24").Address(False, False) And _
         targetArea.Column = Me.Range("C24").Column And _
         targetArea.Columns.Count = 1)
End Function

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    mod_DebugLog.Log "[Sheet1] SelectionChange: " & target.Address & " IsPromptingOffice=" & mod_FillManagerName.IsPromptingOfficeComboBox() & " CellDropdownOpen=" & mod_BasicInfoCellDropdown.IsCellDropdownSessionOpen(Me)
    mod_BasicInfoCalendar.HandleBasicInfoSelectionChange target

    If Not mod_FillManagerName.IsPromptingOfficeComboBox() Then
        If Intersect(target, Me.Range("C6")) Is Nothing Then
            mod_DebugLog.Log "[Sheet1] SelectionChange: C6�ȊO -> HideOfficeComboBox"
            mod_FillManagerName.HideOfficeComboBox Me
        End If
    Else
        mod_DebugLog.Log "[Sheet1] SelectionChange: IsPrompting=True -> Hide�}��"
    End If

    If mod_BasicInfoCellDropdown.IsCellDropdownSessionOpen(Me) Then
        mod_BasicInfoCellDropdown.PollCellDropdownSelection
    End If

    If Not mod_BasicInfoCellDropdown.IsCellDropdownTarget(Me, target) Then
        mod_BasicInfoCellDropdown.HideCellDropdown Me
    End If

    If Not Intersect(target, Me.Range("C21")) Is Nothing Then
        mod_MaterialPriceImport.RefreshUnitPriceProjectNameValidation Me, True
    End If

    Dim vendorNameCell As Range
    Set vendorNameCell = mod_VendorMaster.ResolveVendorNameChangeCell(Me, target)

    If Not vendorNameCell Is Nothing Then
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

Private Sub ComboBox1_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If keyCode = vbKeyReturn Then
        mod_FillManagerName.CommitOfficeComboBoxSelection
    End If
End Sub

Private Sub ComboBox1_LostFocus()
    mod_DebugLog.Log "[Sheet1] ComboBox1_LostFocus IsPrompting=" & mod_FillManagerName.IsPromptingOfficeComboBox()
    If mod_FillManagerName.IsPromptingOfficeComboBox() Then
        mod_DebugLog.Log "[Sheet1] ComboBox1_LostFocus: IsPrompting=True -> �X�L�b�v"
        Exit Sub
    End If
    mod_FillManagerName.CommitOfficeComboBoxSelection False
End Sub

Private Sub ComboBoxVendor_Click()
    mod_VendorMaster.CommitVendorComboBoxSelection Me
End Sub

Private Sub ComboBoxVendor_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If keyCode = vbKeyReturn Then
        mod_VendorMaster.CommitVendorComboBoxSelection Me
    End If
End Sub

Private Sub ComboBoxCellDropdown_Change()
    mod_BasicInfoCellDropdown.ScheduleDeferredCommitCellDropdownSelection
End Sub

Private Sub ComboBoxCellDropdown_Click()
    mod_BasicInfoCellDropdown.ScheduleDeferredCommitCellDropdownSelection
End Sub

Private Sub ComboBoxCellDropdown_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If keyCode = vbKeyReturn Then
        mod_BasicInfoCellDropdown.ScheduleDeferredCommitCellDropdownSelection
    End If
End Sub

Private Sub ComboBoxCellDropdown_LostFocus()
    mod_BasicInfoCellDropdown.HandleCellDropdownLostFocus Me
End Sub

Private Sub btnUpdateBilling_Click()
    mod_BasicInfoUpdate.UpdateBasicInfoPeriod
End Sub

Private Sub btnClearBasicInfo_Click()
    mod_BasicInfoUpdate.ClearBasicInfo
End Sub

Private Function IsIgnoredBasicInfoChange(ByVal ws As Worksheet, _
                                          ByVal target As Range, _
                                          ByVal vendorNameRange As Range, _
                                          ByVal vendorUnitPriceMonitorRange As Range) As Boolean
    If target Is Nothing Then
        IsIgnoredBasicInfoChange = True
        Exit Function
    End If

    If IntersectsBasicInfoMonitorTarget(target, ws.Range("B4,B6,C6,C10,C13,C15:C16,F9,C20,C22,C23,C24,F2")) Then Exit Function
    If IntersectsBasicInfoMonitorTarget(target, vendorNameRange) Then Exit Function
    If IntersectsBasicInfoMonitorTarget(target, vendorUnitPriceMonitorRange) Then Exit Function
    If IntersectsBasicInfoMonitorTarget(target, GetBasicInfoChangeGateRange(ws)) Then Exit Function

    IsIgnoredBasicInfoChange = True
End Function

' �ύX�Ď��Q�[�g: �w�b�_�[�]�L�� + �K�C�h�Ď��͈� + �Œ�Z�� + �Ǝ҃u���b�N(�ő�10��)�̊Ď��s
Private Function GetBasicInfoChangeGateRange(ByVal ws As Worksheet) As Range
    Dim result As Range

    On Error Resume Next
    Set result = mod_OrderTpl_Header.GetBasicInfoHeaderSourceMonitorRange(ws)

    Dim guideRange As Range
    Set guideRange = mod_BasicInfoGuide.GetVendorGuideMonitorRangePublic(ws)
    If Not guideRange Is Nothing Then
        If result Is Nothing Then
            Set result = guideRange
        Else
            Set result = Union(result, guideRange)
        End If
    End If

    ' �Œ�Z��+�{�H��у u���b�N(10��6�s)�͐ÓI�Ȃ��߃L���b�V�����ė��p����
    ' (���ύX���̑傫��Union�č\�z�����B�Ď��Ώۃ Z���W���͏]������v)�B
    Dim staticPart As Range
    Set staticPart = GetStaticGateRange(ws)
    If Not staticPart Is Nothing Then
        If result Is Nothing Then
            Set result = staticPart
        Else
            Set result = Union(result, staticPart)
        End If
    End If
    On Error GoTo 0

    Set GetBasicInfoChangeGateRange = result
End Function

' �Œ�Z��+�{�H��у u���b�N��̐ÓI�ȃQ�[�g�͈͂��\�z���ĕԂ�(�L���b�V��)�B
' BasicInfoVendorColumn(i) �� i �̏������Ȃ̂ł����̏W���͊��S�ɐÓI�B
Private Function GetStaticGateRange(ByVal ws As Worksheet) As Range
    If Not mStaticGateRangeCache Is Nothing Then
        Set GetStaticGateRange = mStaticGateRangeCache
        Exit Function
    End If

    Dim result As Range
    On Error Resume Next
    Set result = ws.Range("C2,C9,C22,C23,F6,F9")

    ' �Ǝ҃u���b�N(F��n�܂�3�񂨂�)�� 10/11/16/27/29/31 �s��
    Dim monitorRows As Variant
    monitorRows = Array(10, 11, 16, 27, 29, 31)

    Dim i As Long
    For i = 1 To BASIC_INFO_VENDOR_MAX_BLOCKS
        Dim valueColumn As Long
        valueColumn = mod_Construction_BasicTotals.BasicInfoVendorColumn(i)

        Dim r As Long
        For r = LBound(monitorRows) To UBound(monitorRows)
            Set result = Union(result, ws.Cells(CLng(monitorRows(r)), valueColumn))
        Next r
    Next i
    On Error GoTo 0

    Set mStaticGateRangeCache = result
    Set GetStaticGateRange = result
End Function

Private Function IntersectsBasicInfoMonitorTarget(ByVal target As Range, _
                                                  ByVal monitorRange As Range) As Boolean
    If target Is Nothing Then Exit Function
    If monitorRange Is Nothing Then Exit Function

    On Error Resume Next
    IntersectsBasicInfoMonitorTarget = Not Application.Intersect(target, monitorRange) Is Nothing
    On Error GoTo 0
End Function

Private Function GetTargetValueText(ByVal target As Range) As String
    On Error GoTo Unreadable

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

Unreadable:
    GetTargetValueText = "(unreadable)"
End Function
