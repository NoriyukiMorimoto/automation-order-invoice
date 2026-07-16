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
            Dim savedWs As Worksheet
            On Error Resume Next
            Set savedWs = mSubconSavedSel.Worksheet
            On Error GoTo 0
            If savedWs Is Nothing Then
                ' 保存していた選択範囲のシートが削除・再作成されて参照切れになった場合の保護
                Set mSubconSavedSel = Nothing
            ElseIf Not (savedWs Is Sh) Then
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

    ' 内訳明旛シ拏トの前回迄(I列)/擅回(L列)の摧量入力で単価列(J/M)を自動入力
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

    ' 監視対象: 施工会社列 + 内訳明細への転記元データ列(整理番号～管理室)
    Dim monitorCols As Range
    Set monitorCols = GetOrderDetailMonitorColumns(Sh)
    If Application.Intersect(Target, monitorCols) Is Nothing Then Exit Sub

    Dim prevEE As Boolean
    prevEE = Application.EnableEvents
    On Error GoTo Cleanup
    Application.EnableEvents = False

    ' 施工会社列・整理番号列の変更時は施工会社別単価列を再計算する
    Dim priceHit As Range
    Set priceHit = Application.Intersect(Target, GetPriceLookupColumns(Sh))
    If Not priceHit Is Nothing Then
        If Sh.AutoFilterMode Then
            mod_Construction_Order_Import.RefreshSubcontractorPriceColumns Sh
        Else
            mod_Construction_Order_Import.RefreshSubcontractorPriceColumns Sh, CollectMonitorChangeRows(Sh, priceHit)
        End If
    End If

    ' 施工会社列の変更時のみ業者名列の幅を自動調整する
    If Not Application.Intersect(Target, GetVendorSelectionColumns(Sh)) Is Nothing Then
        GetVendorSelectionColumns(Sh).AutoFit
    End If

    ' 施工会社列・明細データ列(整理番号～管理室)の変更を各社の内訳明細へ自動反映(遅延予約)
    mod_OrderTpl_Generate.ScheduleOrderDetailRefresh

Cleanup:
    Application.EnableEvents = prevEE
End Sub
Public Sub RunScheduledSubcontractorSelection()
    mod_subcontractorselector.RunScheduledSubcontractorSelection
End Sub

' 内訳明細への転記に影響する監視対象列
' (施工会社列 + 明細データ列 整理番号～管理室)。溶接シートは列オフセット+1。
Private Function GetOrderDetailMonitorColumns(ByVal Sh As Object) As Range
    Dim lastCol As Long
    If mod_Construction_Order_Import.IsWeldingOutputSheet(Sh) Then
        lastCol = COL_MGR + WELDING_OUTPUT_COL_OFFSET
    Else
        lastCol = COL_MGR
    End If
    Set GetOrderDetailMonitorColumns = Sh.Range(Sh.Columns(1), Sh.Columns(lastCol))
End Function

' 施工会社別単価の再計算が必要な列(施工会社列 + 整理番号列)。
Private Function GetPriceLookupColumns(ByVal Sh As Object) As Range
    Dim seiriCol As Long
    If mod_Construction_Order_Import.IsWeldingOutputSheet(Sh) Then
        seiriCol = COL_SEIRI + WELDING_OUTPUT_COL_OFFSET
    Else
        seiriCol = COL_SEIRI
    End If
    Set GetPriceLookupColumns = Application.Union(GetVendorSelectionColumns(Sh), Sh.Columns(seiriCol))
End Function

' 変更セルから重複を除いた行番号コレクションを返す
' (hitRange は Intersect(Target, 監視列) 済みの範囲)
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

