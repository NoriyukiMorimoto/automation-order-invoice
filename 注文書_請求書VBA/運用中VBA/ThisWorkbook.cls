Option Explicit

Private Const SUBCON_HEADER_VENDOR  As String = "é{çHã∆é“"
Private Const SUBCON_HEADER_SEIRI   As String = "êÆóùî‘çÜ"
Private Const SUBCON_DATA_START_ROW As Long = 2
Private Const SUBCON_MENU_CAPTION   As String = "é{çHâÔé–ÇëIëÇ∑ÇÈ"

Private mSubconSavedSel As Range

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    mod_FillManagerName.CancelScheduledOfficeComboBoxPrompt
    mod_VendorMaster.CancelScheduledVendorListDropdown
    mod_BasicInfoCalendar.ClearProjectSelectionState
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
    If Trim$(CStr(Sh.Cells(1, 1).Value)) = SUBCON_HEADER_VENDOR And _
       Trim$(CStr(Sh.Cells(1, 2).Value)) = SUBCON_HEADER_SEIRI Then
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

    If Target.Cells.Count > 1 Then
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
    Application.EnableEvents = True

    If Not mSubconSavedSel Is Nothing Then
        If mSubconSavedSel.Worksheet Is Sh Then
            If Not Application.Intersect(mSubconSavedSel, Target) Is Nothing Then
                If mSubconSavedSel.Cells.Count > 1 Then mSubconSavedSel.Select
            End If
        End If
    End If

    mod_subcontractorselector.SelectSubcontractorForSelection
End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not IsSubconSheet(Sh) Then Exit Sub
    If Not IsVendorSelectionTargetColumn(Sh, Target.Column) Then Exit Sub
    If Target.Row < SUBCON_DATA_START_ROW Then Exit Sub

    Dim cbMenu As Object
    Set cbMenu = Application.CommandBars("Cell")

    Dim ctrl As Object
    For Each ctrl In cbMenu.Controls
        If ctrl.Caption = SUBCON_MENU_CAPTION Then
            ctrl.Delete
            Exit For
        End If
    Next ctrl

    Dim btn As Object
    Set btn = cbMenu.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
    btn.Caption = SUBCON_MENU_CAPTION
    btn.OnAction = "SelectSubcontractorForSelection"
    On Error Resume Next
    btn.FaceId = 626
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If mod_VendorMaster.IsSyncVendorBlocksInProgress() Then Exit Sub

    If mod_MaterialPriceImport.IsConstructionUnitPriceSheet(Sh) Then
    If Application.Intersect(Target, Sh.Range("B:B,E:F")) Is Nothing Then Exit Sub

    Dim unitPricePrevEE As Boolean
    Dim unitPricePrevScreen As Boolean
    unitPricePrevEE = Application.EnableEvents
    unitPricePrevScreen = Application.ScreenUpdating
    On Error GoTo UnitPriceCleanup
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    mod_VendorMaster.HandleConstructionUnitPriceSheetChange Sh, Target

UnitPriceCleanup:
    Application.ScreenUpdating = unitPricePrevScreen
    Application.EnableEvents = unitPricePrevEE
        Exit Sub
    End If

    If Not IsSubconSheet(Sh) Then Exit Sub
    If Application.Intersect(Target, GetVendorSelectionColumns(Sh)) Is Nothing Then Exit Sub

    Dim prevEE As Boolean
    Dim prevScreen As Boolean
    prevEE = Application.EnableEvents
    prevScreen = Application.ScreenUpdating
    On Error GoTo Cleanup
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    mod_Construction_Order_Import.RefreshSubcontractorPriceColumns Sh
    GetVendorSelectionColumns(Sh).AutoFit

Cleanup:
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEE
End Sub
