Option Explicit

Private mSelectedLineNames As Collection
Private mSelectionConfirmed As Boolean
Private mLineCount As Long

Private Const FORM_FONT_SIZE As Single = 12
Private Const LINE_ITEM_HEIGHT As Single = 28
Private Const FORM_MIN_WIDTH As Single = 340
Private Const FORM_MIN_HEIGHT As Single = 350
Private Const FORM_MARGIN As Single = 8
Private Const BUTTON_WIDTH As Single = 70
Private Const BUTTON_HEIGHT As Single = 30
Private Const BUTTON_GAP As Single = 12

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
    itemHeight = LINE_ITEM_HEIGHT

    Dim i As Long
    Dim chk As MSForms.CheckBox
    For i = 1 To mLineCount
        Set chk = GetOrCreateLineCheckBox(frame, i)
        With chk
            .Caption = CStr(lineNames(i))
            .Left = 10
            .Top = 24 + (i - 1) * itemHeight
            .Width = Application.Max(120, frame.Width - 28)
            .Height = 24
            .Tag = CStr(i)
            .value = False
            .Visible = True
            ApplyControlFont chk, True
        End With
    Next i

    On Error Resume Next
    frame.ScrollBars = fmScrollBarsVertical
    frame.ScrollHeight = Application.Max(frame.Height, 32 + mLineCount * itemHeight)
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
        MsgBox SelectAtLeastOneLineNameMessageText(), vbExclamation
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
    Me.Caption = SelectLineNameFormCaptionText()
    ApplyControlFont Me, False

    HideDuplicateTitleLabels
    LayoutStaticControls frame

    frame.Caption = SelectLineNamePromptText()
    ApplyControlFont frame, True

    SetButtonCaption "cmdSelectAll", SelectAllButtonText()
    SetButtonCaption "cmdClearAll", ClearAllButtonText()
    SetButtonCaption "cmdOK", "OK"
    SetButtonCaption "cmdCancel", CancelButtonText()
End Sub

Private Sub LayoutStaticControls(ByVal frame As Object)
    On Error Resume Next
    If Me.Width < FORM_MIN_WIDTH Then Me.Width = FORM_MIN_WIDTH
    If Me.Height < FORM_MIN_HEIGHT Then Me.Height = FORM_MIN_HEIGHT

    Dim insideWidth As Single
    Dim insideHeight As Single
    insideWidth = Me.insideWidth
    insideHeight = Me.insideHeight
    If insideWidth <= 0 Then insideWidth = Me.Width
    If insideHeight <= 0 Then insideHeight = Me.Height

    Dim buttonTop As Single
    buttonTop = insideHeight - FORM_MARGIN - BUTTON_HEIGHT
    If buttonTop < 250 Then buttonTop = 250

    frame.Left = FORM_MARGIN
    frame.Top = FORM_MARGIN
    frame.Width = Application.Max(120, insideWidth - FORM_MARGIN * 2)
    frame.Height = Application.Max(180, buttonTop - FORM_MARGIN - BUTTON_GAP)

    Dim startLeft As Single
    startLeft = (insideWidth - (BUTTON_WIDTH * 4 + BUTTON_GAP * 3)) / 2
    If startLeft < FORM_MARGIN Then startLeft = FORM_MARGIN

    ConfigureButton "cmdSelectAll", startLeft, buttonTop
    ConfigureButton "cmdClearAll", startLeft + (BUTTON_WIDTH + BUTTON_GAP), buttonTop
    ConfigureButton "cmdOK", startLeft + (BUTTON_WIDTH + BUTTON_GAP) * 2, buttonTop
    ConfigureButton "cmdCancel", startLeft + (BUTTON_WIDTH + BUTTON_GAP) * 3, buttonTop
    On Error GoTo 0
End Sub

Private Sub ConfigureButton(ByVal buttonName As String, ByVal buttonLeft As Single, ByVal buttonTop As Single)
    On Error Resume Next
    With Me.Controls(buttonName)
        .Left = buttonLeft
        .Top = buttonTop
        .Width = BUTTON_WIDTH
        .Height = BUTTON_HEIGHT
        .Visible = True
    End With
    On Error GoTo 0
End Sub

Private Sub HideDuplicateTitleLabels()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "Label" Then
            ctrl.Caption = ""
            ctrl.Visible = False
        End If
    Next ctrl
End Sub

Private Sub SetButtonCaption(ByVal buttonName As String, ByVal captionText As String)
    On Error Resume Next
    Me.Controls(buttonName).Caption = captionText
    ApplyControlFont Me.Controls(buttonName), False
    On Error GoTo 0
End Sub

Private Sub ApplyControlFont(ByVal targetControl As Object, ByVal makeBold As Boolean)
    On Error Resume Next
    targetControl.Font.Name = FormFontNameText()
    targetControl.Font.Size = FORM_FONT_SIZE
    targetControl.Font.Bold = makeBold
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

Private Function FormFontNameText() As String
    Static cached As String
    If cached = "" Then
        cached = "BIZ UD" & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
    End If
    FormFontNameText = cached
End Function

Private Function SelectLineNameFormCaptionText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H5BFE) & ChrW$(&H5FDC) & ChrW$(&H7A4D) & ChrW$(&H7B97) & _
                 ChrW$(&H7DDA) & ChrW$(&H533A) & ChrW$(&H9078) & ChrW$(&H629E)
    End If
    SelectLineNameFormCaptionText = cached
End Function

Private Function SelectLineNamePromptText() As String
    Static cached As String
    If cached = "" Then
        cached = LineNameText() & ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & _
                 ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & _
                 ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
    End If
    SelectLineNamePromptText = cached
End Function

Private Function SelectAtLeastOneLineNameMessageText() As String
    Static cached As String
    If cached = "" Then
        cached = LineNameText() & ChrW$(&H3092) & "1" & ChrW$(&H3064) & _
                 ChrW$(&H4EE5) & ChrW$(&H4E0A) & ChrW$(&H9078) & ChrW$(&H629E) & _
                 ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & _
                 ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
    End If
    SelectAtLeastOneLineNameMessageText = cached
End Function

Private Function LineNameText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H7A4D) & ChrW$(&H7B97) & ChrW$(&H7DDA) & ChrW$(&H533A) & ChrW$(&H540D)
    End If
    LineNameText = cached
End Function

Private Function SelectAllButtonText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H5168) & ChrW$(&H9078) & ChrW$(&H629E)
    SelectAllButtonText = cached
End Function

Private Function ClearAllButtonText() As String
    Static cached As String
    If cached = "" Then cached = ChrW$(&H5168) & ChrW$(&H89E3) & ChrW$(&H9664)
    ClearAllButtonText = cached
End Function

Private Function CancelButtonText() As String
    Static cached As String
    If cached = "" Then
        cached = ChrW$(&H30AD) & ChrW$(&H30E3) & ChrW$(&H30F3) & ChrW$(&H30BB) & ChrW$(&H30EB)
    End If
    CancelButtonText = cached
End Function
