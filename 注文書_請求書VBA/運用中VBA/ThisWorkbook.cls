Option Explicit

'==========================================================================
'  施工会社選択(施工指示書・施工通知書シートのA列)
'    工事番号選択(U/V列)と同じ操作感を施工会社選択へ移植したもの。
'      ・A列ダブルクリック        -> frmSubconSelector を表示(複数選択行へ一括適用)
'      ・A列を複数選択して右クリック -> 「施工会社を選択する」メニューから一括適用
'    実体は mod_subcontractorselector.SelectSubcontractorForSelection。
'    対象シートはヘッダー署名(A1="施工業者" / B1="整理番号")で判定する。
'    (取込シート名は取込元A3により動的に決まるため、名前では判定しない。
'      購入充当通知シートはA列を削除しA1=整理番号となるため自動的に対象外)
'==========================================================================
Private Const SUBCON_HEADER_VENDOR  As String = "施工業者"
Private Const SUBCON_HEADER_SEIRI   As String = "整理番号"
Private Const SUBCON_DATA_START_ROW As Long = 2
Private Const SUBCON_MENU_CAPTION   As String = "施工会社を選択する"

' 直前の複数選択範囲(ダブルクリックで単一セルに崩れても保持するため)
Private mSubconSavedSel As Range

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    mod_FillManagerName.CancelScheduledOfficeComboBoxPrompt
    mod_VendorMaster.CancelScheduledVendorListDropdown
    mod_BasicInfoCalendar.ClearProjectSelectionState
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
'  対象シート判定: A1=施工業者 かつ B1=整理番号 のワークシートのみ True
'--------------------------------------------------------------------------
Private Function IsSubconSheet(ByVal Sh As Object) As Boolean
    IsSubconSheet = False
    If TypeName(Sh) <> "Worksheet" Then Exit Function
    On Error Resume Next
    If Trim$(CStr(Sh.Cells(1, 1).Value)) = SUBCON_HEADER_VENDOR And _
       Trim$(CStr(Sh.Cells(1, 2).Value)) = SUBCON_HEADER_SEIRI Then
        IsSubconSheet = True
    End If
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------
'  選択範囲の保持
'    ・複数セル選択時はその範囲を保存
'    ・単一セルに崩れたときは、その単一セルが保存済み範囲内にある場合のみ保持
'      (範囲外クリック・対象外シートでは破棄)
'    これによりダブルクリック直前の複数行選択を失わずに済む。
'--------------------------------------------------------------------------
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

'--------------------------------------------------------------------------
'  A列ダブルクリック -> 施工会社選択フォーム
'    直前に複数行を選択していて、その範囲が今のセルを含むなら範囲を復元してから
'    一括適用する(Excel仕様上、クリック時点で選択は単一セルへ崩れるため)。
'--------------------------------------------------------------------------
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not IsSubconSheet(Sh) Then Exit Sub
    If Target.Column <> 1 Then Exit Sub
    If Target.Row < SUBCON_DATA_START_ROW Then Exit Sub

    Cancel = True                       ' セル編集モードへの移行を抑止
    Application.EnableEvents = True      ' フォーム表示前にイベントを有効化

    ' 直前の複数選択がこのセルを含むなら、その範囲を選択し直す
    If Not mSubconSavedSel Is Nothing Then
        If mSubconSavedSel.Worksheet Is Sh Then
            If Not Application.Intersect(mSubconSavedSel, Target) Is Nothing Then
                If mSubconSavedSel.Cells.Count > 1 Then mSubconSavedSel.Select
            End If
        End If
    End If

    mod_subcontractorselector.SelectSubcontractorForSelection
End Sub

'--------------------------------------------------------------------------
'  A列右クリック -> 「施工会社を選択する」を右クリックメニュー先頭へ追加
'    複数選択 -> 右クリック -> 実行 で一括適用が可能(右クリックは選択を保持)。
'    Temporary:=True のためブックを閉じると自動的に消える。
'--------------------------------------------------------------------------
Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not IsSubconSheet(Sh) Then Exit Sub
    If Target.Column <> 1 Then Exit Sub
    If Target.Row < SUBCON_DATA_START_ROW Then Exit Sub

    Dim cbMenu As Object
    Set cbMenu = Application.CommandBars("Cell")

    ' 既存の同名コマンドを削除(重複防止)
    Dim ctrl As Object
    For Each ctrl In cbMenu.Controls
        If ctrl.Caption = SUBCON_MENU_CAPTION Then
            ctrl.Delete
            Exit For
        End If
    Next ctrl

    ' 先頭に追加(OnActionは標準モジュールのPublic Sub)
    Dim btn As Object
    Set btn = cbMenu.Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
    btn.Caption = SUBCON_MENU_CAPTION
    btn.OnAction = "SelectSubcontractorForSelection"
    On Error Resume Next
    btn.FaceId = 626
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
'  A列で施工会社が入力/選択されたら、その名前にA列幅をフィットさせる
'    (ドロップダウン選択・手入力経路。フォーム経由はフォーム側でフィット済み)
'--------------------------------------------------------------------------
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsSubconSheet(Sh) Then Exit Sub
    If Application.Intersect(Target, Sh.Columns(1)) Is Nothing Then Exit Sub

    Dim prevEE As Boolean
    prevEE = Application.EnableEvents
    Application.EnableEvents = False
    Sh.Columns(1).AutoFit
    Application.EnableEvents = prevEE
End Sub
