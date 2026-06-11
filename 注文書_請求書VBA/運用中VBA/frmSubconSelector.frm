Option Explicit

Public confirmed As Boolean
Public SelectedCompany As String

Private Const FORM_FONT_SIZE As Single = 14
Private Const COMPANY_ITEM_HEIGHT As Single = 28
Private Const COMPANY_OPTION_LEFT As Single = 6

Private mAll As Variant
Private mCompanyFrame As Object
Private mVisibleCompanyCount As Long

Private Sub UserForm_Initialize()
    confirmed = False
    SelectedCompany = ""
    Me.Caption = "施工会社選択"

    ConfigureStaticControls
    HideCompanyOptions
End Sub

Public Sub SetCompanies(ByVal names As Variant)
    mAll = names
    RefreshList ""
End Sub

Private Sub RefreshList(ByVal keyword As String)
    Dim frame As Object
    Set frame = CompanyFrame()
    If frame Is Nothing Then Exit Sub

    HideCompanyOptions

    If Not IsArray(mAll) Then Exit Sub

    Dim kw As String
    kw = Trim$(keyword)

    Dim i As Long
    Dim nm As String
    Dim itemIndex As Long
    itemIndex = 0

    For i = LBound(mAll) To UBound(mAll)
        nm = CStr(mAll(i))
        If kw = "" Or InStr(1, nm, kw, vbTextCompare) > 0 Then
            itemIndex = itemIndex + 1
            ShowCompanyOption frame, itemIndex, nm
        End If
    Next i

    mVisibleCompanyCount = itemIndex

    On Error Resume Next
    frame.ScrollBars = fmScrollBarsVertical
    frame.ScrollHeight = Application.Max(frame.Height, 8 + itemIndex * COMPANY_ITEM_HEIGHT)
    On Error GoTo 0
End Sub

Private Sub txtSearch_Change()
    Dim t As Object
    Set t = FindCtl("txtSearch")
    If t Is Nothing Then Exit Sub
    RefreshList CStr(t.text)
End Sub

Private Sub cmdOK_Click()
    ConfirmSelection
End Sub

Private Sub cmdCancel_Click()
    confirmed = False
    Me.Hide
End Sub

Private Sub ConfirmSelection()
    Dim chosen As String
    chosen = GetSelectedCompanyName()
    If chosen = "" Then
        MsgBox "施工会社を選択してください。", vbExclamation
        Exit Sub
    End If

    SelectedCompany = chosen
    confirmed = True
    Me.Hide
End Sub

Private Function GetSelectedCompanyName() As String
    Dim frame As Object
    Set frame = CompanyFrame()
    If frame Is Nothing Then Exit Function

    Dim ctrl As Control
    For Each ctrl In frame.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.Visible And ctrl.value = True Then
                GetSelectedCompanyName = CStr(ctrl.Caption)
                Exit Function
            End If
        End If
    Next ctrl
End Function

Private Sub ShowCompanyOption(ByVal frame As Object, ByVal itemIndex As Long, ByVal companyName As String)
    Dim opt As MSForms.OptionButton
    Set opt = GetOrCreateCompanyOption(frame, itemIndex)

    With opt
        .Caption = companyName
        .Left = COMPANY_OPTION_LEFT
        .Top = 4 + (itemIndex - 1) * COMPANY_ITEM_HEIGHT
        .Width = Application.Max(120, frame.Width - (COMPANY_OPTION_LEFT * 2))
        .Height = COMPANY_ITEM_HEIGHT - 4
        .value = False
        .Visible = True
        ApplyControlFont opt, True
    End With
End Sub

Private Sub HideCompanyOptions()
    Dim frame As Object
    Set frame = CompanyFrame()
    If frame Is Nothing Then Exit Sub

    Dim ctrl As Control
    For Each ctrl In frame.Controls
        If TypeName(ctrl) = "OptionButton" Then
            ctrl.Visible = False
            ctrl.value = False
        End If
    Next ctrl

    mVisibleCompanyCount = 0
End Sub

Private Function CompanyFrame() As Object
    Dim lst As Object
    Set lst = FindCtl("lstCompanies")
    If lst Is Nothing Then Exit Function

    If mCompanyFrame Is Nothing Then
        On Error Resume Next
        Set mCompanyFrame = Me.Controls("fraCompanies")
        On Error GoTo 0

        If mCompanyFrame Is Nothing Then
            Set mCompanyFrame = Me.Controls.Add("Forms.Frame.1", "fraCompanies", True)
            With mCompanyFrame
                .Left = lst.Left
                .Top = lst.Top
                .Width = lst.Width
                .Height = lst.Height
                .Caption = ""
                .BorderStyle = fmBorderStyleSingle
            End With
        End If

        lst.Visible = False
    End If

    Set CompanyFrame = mCompanyFrame
End Function

Private Function GetOrCreateCompanyOption(ByVal frame As Object, ByVal index As Long) As MSForms.OptionButton
    Dim controlName As String
    controlName = "optCompany" & CStr(index)

    On Error Resume Next
    Set GetOrCreateCompanyOption = frame.Controls(controlName)
    On Error GoTo 0

    If GetOrCreateCompanyOption Is Nothing Then
        Set GetOrCreateCompanyOption = frame.Controls.Add("Forms.OptionButton.1", controlName, True)
    End If
End Function

Private Sub ConfigureStaticControls()
    On Error Resume Next
    Me.Controls("cmdOK").Caption = "適用"
    Me.Controls("cmdCancel").Caption = "キャンセル"
    On Error GoTo 0

    ApplyControlFont Me, False

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox", "CommandButton"
                ApplyControlFont ctrl, False
        End Select
    Next ctrl
End Sub

Private Sub ApplyControlFont(ByVal targetControl As Object, ByVal makeBold As Boolean)
    On Error Resume Next
    targetControl.Font.Name = FormFontNameText()
    targetControl.Font.Size = FORM_FONT_SIZE
    targetControl.Font.Bold = makeBold
    On Error GoTo 0
End Sub

Private Function FormFontNameText() As String
    Static cached As String
    If cached = "" Then
        cached = "BIZ UD" & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
    End If
    FormFontNameText = cached
End Function

Private Function FindCtl(ByVal nm As String) As Object
    On Error Resume Next
    Set FindCtl = Me.Controls(nm)
    On Error GoTo 0
End Function
