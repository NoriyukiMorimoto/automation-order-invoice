Option Explicit

Private mSelectedLineNames As Collection
Private mSelectionConfirmed As Boolean
Private mLineCount As Long

Public Sub InitLineNames(ByVal lineNames As Collection)
    Dim frame As Object
    Set frame = LineFrame()
    If frame Is Nothing Then Exit Sub

    mSelectionConfirmed = False
    Set mSelectedLineNames = New Collection
    mLineCount = 0

    ConfigureStaticCaptions frame
    HideExistingLineCheckBoxes frame

    If lineNames Is Nothing Then Exit Sub
    mLineCount = lineNames.Count
    If mLineCount = 0 Then Exit Sub

    Dim itemHeight As Single
    itemHeight = 26

    Dim i As Long
    Dim chk As MSForms.CheckBox
    For i = 1 To mLineCount
        Set chk = GetOrCreateLineCheckBox(frame, i)
        With chk
            .Caption = CStr(lineNames(i))
            .Font.Size = 12
            .Left = 10
            .Top = 14 + (i - 1) * itemHeight
            .Width = Application.Max(120, frame.Width - 28)
            .Height = 20
            .Tag = CStr(i)
            .value = False
            .Visible = True
        End With
    Next i

    On Error Resume Next
    frame.ScrollBars = fmScrollBarsVertical
    frame.ScrollHeight = Application.Max(frame.Height, 20 + mLineCount * itemHeight)
    On Error GoTo 0
End Sub

Public Sub InitRooms(ByRef rooms() As String, ByVal cnt As Long)
    Dim lineNames As Collection
    Set lineNames = New Collection

    Dim i As Long
    For i = 0 To cnt - 1
        lineNames.Add rooms(i)
    Next i

    InitLineNames lineNames
End Sub

Public Property Get SelectionConfirmed() As Boolean
    SelectionConfirmed = mSelectionConfirmed
End Property

Public Function GetSelectedLineNames() As Collection
    Dim result As Collection
    Set result = New Collection

    If Not mSelectedLineNames Is Nothing Then
        Dim item As Variant
        For Each item In mSelectedLineNames
            result.Add CStr(item)
        Next item
    End If

    Set GetSelectedLineNames = result
End Function

Private Sub cmdOK_Click()
    Set mSelectedLineNames = New Collection

    Dim frame As Object
    Set frame = LineFrame()
    If frame Is Nothing Then Exit Sub

    Dim ctrl As Control
    For Each ctrl In frame.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Visible And ctrl.value = True Then
                mSelectedLineNames.Add CStr(ctrl.Caption)
            End If
        End If
    Next ctrl

    If mSelectedLineNames.Count = 0 Then
        MsgBox "積算線区名を1つ以上選択してください。", vbExclamation
        Exit Sub
    End If

    mSelectionConfirmed = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    mSelectionConfirmed = False
    Me.Hide
End Sub

Private Sub cmdSelectAll_Click()
    SetLineCheckBoxesValue True
End Sub

Private Sub cmdClearAll_Click()
    SetLineCheckBoxesValue False
End Sub

Private Sub ConfigureStaticCaptions(ByVal frame As Object)
    Me.Caption = "対応積算線区選択"
    frame.Caption = "積算線区名を選択してください。"

    SetButtonCaption "cmdSelectAll", "全選択"
    SetButtonCaption "cmdClearAll", "全解除"
    SetButtonCaption "cmdOK", "OK"
    SetButtonCaption "cmdCancel", "キャンセル"
End Sub

Private Sub SetButtonCaption(ByVal buttonName As String, ByVal captionText As String)
    On Error Resume Next
    Me.Controls(buttonName).Caption = captionText
    On Error GoTo 0
End Sub

Private Sub HideExistingLineCheckBoxes(ByVal frame As Object)
    Dim ctrl As Control
    For Each ctrl In frame.Controls
        If TypeName(ctrl) = "CheckBox" Then ctrl.Visible = False
    Next ctrl
End Sub

Private Function GetOrCreateLineCheckBox(ByVal frame As Object, ByVal index As Long) As MSForms.CheckBox
    Dim controlName As String
    controlName = "chkLine" & CStr(index)

    On Error Resume Next
    Set GetOrCreateLineCheckBox = frame.Controls(controlName)
    On Error GoTo 0

    If GetOrCreateLineCheckBox Is Nothing Then
        Set GetOrCreateLineCheckBox = frame.Controls.Add("Forms.CheckBox.1", controlName, True)
    End If
End Function

Private Sub SetLineCheckBoxesValue(ByVal checkedValue As Boolean)
    Dim frame As Object
    Set frame = LineFrame()
    If frame Is Nothing Then Exit Sub

    Dim ctrl As Control
    For Each ctrl In frame.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Visible Then ctrl.value = checkedValue
        End If
    Next ctrl
End Sub

Private Function LineFrame() As Object
    On Error Resume Next
    Set LineFrame = Me.Controls("Frame1")
    On Error GoTo 0
End Function
