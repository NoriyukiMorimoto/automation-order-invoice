Option Explicit

' ========================================================
' mod_BasicInfoCellDropdown
' 基本情報シートの C22 / C23 をダブルクリックすると、
' そのセルの入力規則(リスト)からドロップダウンを表示する。
'
' 方式: SendKeys は使わず(NumLock保護)、セル上に一時 Forms.ComboBox.1
'       を重ねて .DropDown で開く(office C6 / 業者名 と同方式)。
' 選択値はセルへ反映し、既存の Worksheet_Change を発火させる
' (EnableEvents=True のまま書込むため、単価再計算・ガイド解除が走る)。
'
' リスト内容はセルの入力規則(Validation.Formula1)を実行時に読むため、
' 範囲参照(=$AG$2:$AG$3 等)・インライン("あり,なし" 等)の双方に追従する。
'
' 改修履歴: CHANGELOG.md 参照
' ========================================================

Private Const CELL_DROPDOWN_COMBO_NAME As String = "ComboBoxCellDropdown"
Private mCellDropdownTargetAddress As String
Private mInCellDropdownPrompt As Boolean

Public Function IsPromptingCellDropdown() As Boolean
    IsPromptingCellDropdown = mInCellDropdownPrompt
End Function

' ドロップダウン対象セル(値セル。結合は左上)。必要に応じて追加可能。
Private Function CellDropdownTargetAddresses() As Variant
    CellDropdownTargetAddresses = Array("C22", "C23")
End Function

' 指定セルがドロップダウン対象(C22/C23、結合範囲含む)か
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

' 対象セルのアンカー(結合左上)を返す
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

' ダブルクリックから呼ぶ: セル上にコンボを出してドロップダウンを開く
Public Sub ShowCellValidationDropdown(ByVal wsInfo As Worksheet, ByVal target As Range)
    If wsInfo Is Nothing Or target Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = ResolveCellDropdownAnchor(wsInfo, target)
    If anchor Is Nothing Then Exit Sub

    On Error GoTo CleanFail
    DeleteCellDropdownComboBox wsInfo

    Dim ole As OLEObject
    Set ole = GetCellDropdownComboBox(wsInfo, anchor)
    If ole Is Nothing Then Exit Sub

    LoadComboItemsFromValidation anchor, ole
    If ole.Object.ListCount = 0 Then
        ' 入力規則リストが無い/空 -> コンボは出さない
        DeleteCellDropdownComboBox wsInfo
        Exit Sub
    End If

    mCellDropdownTargetAddress = anchor.Address(False, False)
    mInCellDropdownPrompt = True
    FitComboToCell anchor, ole

    wsInfo.Activate
    anchor.Select
    ole.Visible = True
    ole.Activate
    On Error Resume Next
    ole.Object.DropDown
    On Error GoTo CleanFail
    Exit Sub

CleanFail:
    ResetCellDropdownSession
    DeleteCellDropdownComboBox wsInfo
End Sub

' コンボの Click / Change / Enter / LostFocus から呼ぶ: 選択値をセルへ反映
Public Sub CommitCellDropdownSelection(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub

    If TryCommitCellDropdownSelection(wsInfo) Then
        CloseCellDropdownSession wsInfo
    End If
End Sub

' コンボからフォーカスが外れた時: 選択済みなら反映、未選択なら閉じる
Public Sub HandleCellDropdownLostFocus(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    If Not mInCellDropdownPrompt Then Exit Sub

    If TryCommitCellDropdownSelection(wsInfo) Then
        CloseCellDropdownSession wsInfo
    Else
        HideCellDropdown wsInfo
    End If
End Sub

' 選択が対象セルから外れた時などに呼ぶ: コンボを消す
Public Sub HideCellDropdown(ByVal wsInfo As Worksheet)
    If wsInfo Is Nothing Then Exit Sub
    ResetCellDropdownSession
    DeleteCellDropdownComboBox wsInfo
End Sub

Private Function TryCommitCellDropdownSelection(ByVal wsInfo As Worksheet) As Boolean
    On Error GoTo CleanFail

    If Len(mCellDropdownTargetAddress) = 0 Then Exit Function

    Dim combo As Object
    Set combo = wsInfo.OLEObjects(CELL_DROPDOWN_COMBO_NAME).Object

    Dim selectedValue As String
    selectedValue = ReadComboSelectedText(combo)
    If Len(selectedValue) = 0 Then Exit Function

    Dim anchor As Range
    Set anchor = wsInfo.Range(mCellDropdownTargetAddress)

    ' 実際に値が変わる場合のみ書込み(不要な Worksheet_Change の連鎖を避ける)
    If StrComp(selectedValue, CStr(anchor.value), vbBinaryCompare) <> 0 Then
        anchor.value = selectedValue
    End If

    TryCommitCellDropdownSelection = True
    Exit Function

CleanFail:
    TryCommitCellDropdownSelection = False
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
    targetAddress = mCellDropdownTargetAddress

    ResetCellDropdownSession
    DeleteCellDropdownComboBox wsInfo

    On Error Resume Next
    If Len(targetAddress) > 0 Then wsInfo.Range(targetAddress).Select
    On Error GoTo 0
End Sub

Private Sub ResetCellDropdownSession()
    mInCellDropdownPrompt = False
    mCellDropdownTargetAddress = ""
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
        .Placement = xlMoveAndSize
    End With
    On Error GoTo 0
End Sub

' セルの入力規則(リスト)からコンボの項目を読み込む。
' 範囲参照(=...)・定義名・インライン("a,b,c")の三形態に対応。
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
                Set listRange = Nothing
                Set listRange = anchor.Worksheet.Range(Mid$(f1, 2))
                If listRange Is Nothing Then
                    Set listRange = anchor.Worksheet.Parent.Names(Mid$(f1, 2)).RefersToRange
                End If
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
        .value = CStr(anchor.value)
    End With

    On Error GoTo 0
End Sub

Private Sub DeleteCellDropdownComboBox(ByVal wsInfo As Worksheet)
    On Error Resume Next
    With wsInfo.OLEObjects(CELL_DROPDOWN_COMBO_NAME)
        .Visible = False
        .Delete
    End With
    On Error GoTo 0
End Sub
