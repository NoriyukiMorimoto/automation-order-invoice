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

' 本シートのActivateごとに重い再構築(ガイド再構築・検証再計算・クリーンアップ)を毎回行わないためのフラグ。
' セッション中は最初の1回のみ実行すればよいための仕組み(2回目以降の切替では省略して負荷を抑える)。
' シート変更イベント/選択変更などでは維持し続けるため、切替の都度リセットする必要はない。
Private mBasicInfoActivateInitialized As Boolean

' 変更監視ゲート範囲のうち、静的な部分(基本情報固定セル+施工会社ブロックの監視行)のキャッシュ。
' レイアウトは静的なため一度構築すれば再利用でき、変更のたびに60個以上のUnionを再構築せずに済む。
' 対象セル集合は従来と同じ(基本情報固定セル+施工会社ブロック10社×6行)。Activateのたびに再構築する。
Private mStaticGateRangeCache As Range

Private Sub Worksheet_Activate()
    Dim g As New clsPerfGuard
    g.Suspend

    On Error GoTo ExitHandler

    ' 毎回必要な軽量処理(キャッシュクリア・コンボボックス非表示・検証初期化など)
    mod_VendorMaster.ClearVendorRowsCache
    Set mStaticGateRangeCache = Nothing   ' 変更ゲート範囲の静的キャッシュをクリア(次回アクセス時に再構築)
    mod_FillManagerName.HideOfficeComboBox Me
    mod_VendorMaster.EnsureVendorCountInputValidation Me
    mod_OrderTpl_Shared.OrderTplRepairAllGeneratedPlaceholderFormulas
    mod_VendorMaster.InitVendorBlockCountFromSheet Me
    mod_BasicInfoGuide.InitBasicInfoGuide Me

    ' 重い再構築(ガイド再構築・レガシー検証の整理・検証再計算)はセッション中で最初の1回だけ行う。
    ' 2回目以降の切替では省略する(内容は変わらないため、変更イベント/選択変更などで最新に維持される)。
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
        mod_DebugLog.Log "[Sheet1] 変更対象外 -> Exit"
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
        mod_DebugLog.Log "[Sheet1] B4/B6変更 -> RefreshBranchOfficeValidation(False) 呼び出し"
        needComboPrompt = mod_FillManagerName.RefreshBranchOfficeValidation(False)
        mod_DebugLog.Log "[Sheet1] RefreshBranchOfficeValidation 戻り値 needComboPrompt=" & needComboPrompt
    End If

    If Not Intersect(target, Me.Range("B6")) Is Nothing Then
        mod_VendorMaster.RefreshVendorListForBasicInfo Me
    End If

    If Not Intersect(target, Me.Range("F9")) Is Nothing Then
        mod_VendorMaster.SyncVendorBlocksFromCount Me
        Set vendorNameRange = mod_VendorMaster.GetVendorNameRange(Me)
        mod_Construction_Order_Import.RefreshBasicInfoConstructionTotals
        ' 行削除等で対象外になった受注者用テンプレートシートを削除
        mod_OrderTpl_Generate.RemoveOrphanGeneratedSheets Me
    End If

    If Not Intersect(target, Me.Range("B6,C6")) Is Nothing Then
        mod_DebugLog.Log "[Sheet1] B6/C6変更 IsSuppressing=" & mod_FillManagerName.IsSuppressingC6Change()
        If Not mod_FillManagerName.IsSuppressingC6Change() Then
            mod_DebugLog.Log "[Sheet1] SilentClearBasicInfo 呼び出し"
            mod_BasicInfoUpdate.SilentClearBasicInfo Me
            mod_DebugLog.Log "[Sheet1] SilentClearBasicInfo 完了"
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
        ' 施工会社確定に伴い、対応する受注者用テンプレートシートへ反映/転記
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

    ' 基本情報のヘッダー転記元セル変更を監視し、各受注者用テンプレート各シートへ反映
    ' (10行目の工事区分など列単位の適用にタブ色反映が絡むため、ガイド更新の後に実行する)
    mod_OrderTpl_Header.HandleBasicInfoHeaderSourceChange Me, target

    mod_DebugLog.Log "[Sheet1] needComboPrompt=" & needComboPrompt
    If needComboPrompt Then
        mod_DebugLog.Log "[Sheet1] PromptOfficeComboBox 呼び出し前 EnableEvents=" & Application.EnableEvents
        mod_FillManagerName.PromptOfficeComboBox
        mod_DebugLog.Log "[Sheet1] PromptOfficeComboBox 呼び出し後 EnableEvents=" & Application.EnableEvents
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
            mod_DebugLog.Log "[Sheet1] SelectionChange: C6以外 -> HideOfficeComboBox"
            mod_FillManagerName.HideOfficeComboBox Me
        End If
    Else
        mod_DebugLog.Log "[Sheet1] SelectionChange: IsPrompting=True -> Hideしない"
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
        mod_DebugLog.Log "[Sheet1] ComboBox1_LostFocus: IsPrompting=True -> スキップ"
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

' 変更監視ゲート: ヘッダー転記元 + ガイド監視範囲 + 静的セル + 施工会社ブロック(最大10社)の監視行
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

    ' 静的セル+施工会社ブロック(10社×6行)は静的なのでキャッシュして流用する
    ' (基本情報変更のたびに大きいUnionを再構築しない。監視対象セル集合は従来通り)
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

' 静的セル+施工会社ブロックの静的なゲート範囲を構築して返す(キャッシュ)
' BasicInfoVendorColumn(i) は i の最大数が固定なので、この集合は完全に静的。
Private Function GetStaticGateRange(ByVal ws As Worksheet) As Range
    If Not mStaticGateRangeCache Is Nothing Then
        Set GetStaticGateRange = mStaticGateRangeCache
        Exit Function
    End If

    Dim result As Range
    On Error Resume Next
    Set result = ws.Range("C2,C9,C22,C23,F6,F9")

    ' 施工会社ブロック(F列始まり3列おき)の 10/11/16/27/29/31 行を監視
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