Option Explicit

Private Const SUBCON_DATA_START_ROW As Long = 2

Private mSubconSavedSel As Range

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    mod_WorkbookOptimize.OptimizeWorkbookSilent ThisWorkbook
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    mod_FillManagerName.CancelScheduledOfficeComboBoxPrompt
    mod_VendorMaster.CancelScheduledVendorListDropdown
    mod_BasicInfoCalendar.ClearProjectSelectionState
    mod_OrderTpl_Generate.CancelScheduledOrderDetailRefresh
    mod_OrderTpl_Generate.CancelScheduledVendorSheetGeneration
    On Error GoTo 0
End Sub

Private Function IsSubconSheet(ByVal Sh As Object) As Boolean
    IsSubconSheet = False
    If TypeName(Sh) <> "Worksheet" Then Exit Function
    On Error Resume Next
    If mod_Construction_Order_Import.IsWeldingOutputSheet(Sh) Then
        IsSubconSheet = True
        Exit Function
    End If
    If Trim$(CStr(Sh.Cells(1, 1).Value)) = SubconHeaderVendorText() And _
       Trim$(CStr(Sh.Cells(1, 2).Value)) = SubconHeaderSeiriText() Then
        IsSubconSheet = True
    End If
    On Error GoTo 0
End Function

Private Function IsVendorSelectionTargetColumn(ByVal Sh As Object, ByVal columnIndex As Long) As Boolean
    If mod_Construction_Order_Import.IsWeldingOutputSheet(Sh) Then
        IsVendorSelectionTargetColumn = (columnIndex = 1 Or columnIndex = 2)
    Else
        IsVendorSelectionTargetColumn = (columnIndex = 1)
    End If
End Function

Private Function GetVendorSelectionColumns(ByVal Sh As Object) As Range
    If mod_Construction_Order_Import.IsWeldingOutputSheet(Sh) Then
        Set GetVendorSelectionColumns = Sh.Columns("A:B")
    Else
        Set GetVendorSelectionColumns = Sh.Columns(1)
    End If
End Function

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsSubconSheet(Sh) Then
        Set mSubconSavedSel = Nothing
        Exit Sub
    End If

    If Target.CountLarge > 1 Then
        Set mSubconSavedSel = Target
    Else
        If Not mSubconSavedSel Is Nothing Then
            If Not (mSubconSavedSel.Worksheet Is Sh) Then
                Set mSubconSavedSel = Nothing
            ElseIf Application.Intersect(mSubconSavedSel, Target) Is Nothing Then
                Set mSubconSavedSel = Nothing
            End If
        End If
    End If
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not IsSubconSheet(Sh) Then Exit Sub
    If Not IsVendorSelectionTargetColumn(Sh, Target.Column) Then Exit Sub
    If Target.Row < SUBCON_DATA_START_ROW Then Exit Sub

    Cancel = True

    mod_subcontractorselector.SetSubcontractorTrigger "dblclick"
    mod_subcontractorselector.PrepareSubcontractorSelection Sh, Target, mSubconSavedSel
    mod_subcontractorselector.RequestSubcontractorSelection
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not IsSubconSheet(Sh) Then Exit Sub
    If Not IsVendorSelectionTargetColumn(Sh, Target.Column) Then Exit Sub
    If Target.Row < SUBCON_DATA_START_ROW Then Exit Sub

    mod_subcontractorselector.SetSubcontractorTrigger "rightclick-menu"
    mod_subcontractorselector.PrepareSubcontractorSelection Sh, Target

    Dim cbMenu As Object
    Set cbMenu = Application.CommandBars("Cell")

    Dim ctrl As Object
    For Each ctrl In cbMenu.Controls
        If ctrl.Caption = SubconMenuCaptionText() Then
            ctrl.Delete
            Exit For
        End If
    Next ctrl

    Dim btn As Object
    Set btn = cbMenu.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
    btn.Caption = SubconMenuCaptionText()
    btn.OnAction = "mod_subcontractorselector.RequestSubcontractorSelection"
    On Error Resume Next
    btn.FaceId = 626
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If mod_VendorMaster.IsSyncVendorBlocksInProgress() Then Exit Sub

    ' “à–ó–¾×ƒV[ƒg‚Ì‘O‰ñ–˜(I—ñ)/¡‰ñ(L—ñ)‚Ì”—Ê“ü—Í‚Å’P‰¿—ñ(J/M)‚ðŽ©“®“ü—Í
    mod_OrderTpl_Detail.HandleBreakdownQuantityCellChange Sh, Target

    If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(Sh) Then
        If Application.Intersect(Target, Sh.Range("B:B,E:F")) Is Nothing Then Exit Sub

        Dim unitPricePrevEE As Boolean
        unitPricePrevEE = Application.EnableEvents
        On Error GoTo UnitPriceCleanup
        Application.EnableEvents = False

        mod_VendorMaster.HandleConstructionUnitPriceSheetChange Sh, Target

UnitPriceCleanup:
        Application.EnableEvents = unitPricePrevEE
        Exit Sub
    End If

    If Not IsSubconSheet(Sh) Then Exit Sub

    ' ŠÄŽ‹‘ÎÛ: Ž{H‰ïŽÐ—ñ + “à–ó–¾×‚Ö‚Ì“]‹LŒ³ƒf[ƒ^—ñ(®—”Ô†`ŠÇ—Žº)
    Dim monitorCols As Range
    Set monitorCols = GetOrderDetailMonitorColumns(Sh)
    If Application.Intersect(Target, monitorCols) Is Nothing Then Exit Sub

    Dim prevEE As Boolean
    prevEE = Application.EnableEvents
    On Error GoTo Cleanup
    Application.EnableEvents = False

    ' Ž{H‰ïŽÐ—ñE®—”Ô†—ñ‚Ì•ÏXŽž‚ÍŽ{H‰ïŽÐ•Ê’P‰¿—ñ‚ðÄŒvŽZ‚·‚é
    Dim priceHit As Range
    Set priceHit = Application.Intersect(Target, GetPriceLookupColumns(Sh))
    If Not priceHit Is Nothing Then
        If Sh.AutoFilterMode Then
            mod_Construction_Order_Import.RefreshSubcontractorPriceColumns Sh
        Else
            mod_Construction_Order_Import.RefreshSubcontractorPriceColumns Sh, CollectMonitorChangeRows(Sh, priceHit)
        End If
    End If

    ' Ž{H‰ïŽÐ—ñ‚Ì•ÏXŽž‚Ì‚Ý‹ÆŽÒ–¼—ñ‚Ì•‚ðŽ©“®’²®‚·‚é
    If Not Application.Intersect(Target, GetVendorSelectionColumns(Sh)) Is Nothing Then
        GetVendorSelectionColumns(Sh).AutoFit
    End If

    ' Ž{H‰ïŽÐ—ñE–¾×ƒf[ƒ^—ñ(®—”Ô†`ŠÇ—Žº)‚Ì•ÏX‚ðŠeŽÐ‚Ì“à–ó–¾×‚ÖŽ©“®”½‰f(’x‰„—\–ñ)
    mod_OrderTpl_Generate.ScheduleOrderDetailRefresh

Cleanup:
    Application.EnableEvents = prevEE
End Sub
Public Sub RunScheduledSubcontractorSelection()
    mod_subcontractorselector.RunScheduledSubcontractorSelection
End Sub

' “à–ó–¾×‚Ö‚Ì“]‹L‚É‰e‹¿‚·‚éŠÄŽ‹‘ÎÛ—ñ
' (Ž{H‰ïŽÐ—ñ + –¾×ƒf[ƒ^—ñ ®—”Ô†`ŠÇ—Žº)B—nÚƒV[ƒg‚Í—ñƒIƒtƒZƒbƒg+1B
Private Function GetOrderDetailMonitorColumns(ByVal Sh As Object) As Range
    Dim lastCol As Long
    If mod_Construction_Order_Import.IsWeldingOutputSheet(Sh) Then
        lastCol = COL_MGR + WELDING_OUTPUT_COL_OFFSET
    Else
        lastCol = COL_MGR
    End If
    Set GetOrderDetailMonitorColumns = Sh.Range(Sh.Columns(1), Sh.Columns(lastCol))
End Function

' Ž{H‰ïŽÐ•Ê’P‰¿‚ÌÄŒvŽZ‚ª•K—v‚È—ñ(Ž{H‰ïŽÐ—ñ + ®—”Ô†—ñ)B
Private Function GetPriceLookupColumns(ByVal Sh As Object) As Range
    Dim seiriCol As Long
    If mod_Construction_Order_Import.IsWeldingOutputSheet(Sh) Then
        seiriCol = COL_SEIRI + WELDING_OUTPUT_COL_OFFSET
    Else
        seiriCol = COL_SEIRI
    End If
    Set GetPriceLookupColumns = Application.Union(GetVendorSelectionColumns(Sh), Sh.Columns(seiriCol))
End Function

' •ÏXƒZƒ‹‚©‚çd•¡‚ðœ‚¢‚½s”Ô†ƒRƒŒƒNƒVƒ‡ƒ“‚ð•Ô‚·
' (hitRange ‚Í Intersect(Target, ŠÄŽ‹—ñ) Ï‚Ý‚Ì”ÍˆÍ)
Private Function CollectMonitorChangeRows(ByVal ws As Worksheet, ByVal hitRange As Range) As Collection
    Dim result As New Collection
    If hitRange Is Nothing Then
        Set CollectMonitorChangeRows = result
        Exit Function
    End If

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim c As Range
    For Each c In hitRange.Cells
        If c.Row >= SUBCON_DATA_START_ROW Then
            Dim rowKey As String
            rowKey = CStr(c.Row)
            If Not seen.Exists(rowKey) Then
                seen.Add rowKey, True
                result.Add c.Row
            End If
        End If
    Next c

    Set CollectMonitorChangeRows = result
End Function

Private Function SubconHeaderVendorText() As String
    SubconHeaderVendorText = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H696D) & ChrW$(&H8005)
End Function

Private Function SubconHeaderSeiriText() As String
    SubconHeaderSeiriText = ChrW$(&H6574) & ChrW$(&H7406) & ChrW$(&H756A) & ChrW$(&H53F7)
End Function

Private Function SubconMenuCaptionText() As String
    SubconMenuCaptionText = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H4F1A) & ChrW$(&H793E) & _
                            ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3059) & ChrW$(&H308B)
End Function

