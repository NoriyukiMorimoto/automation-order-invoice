Option Explicit

' frmSubconSelector: 施工会社選択フォーム
' 改修: 動的 OptionButton 生成を廃止し、既存の lstCompanies(ListBox)へ
'       AddItem する方式へ統一。動的コントロール生成によるハングを回避。

Public confirmed As Boolean
Public SelectedCompany As String

Private Const FORM_FONT_SIZE As Single = 14

Private mAll As Variant

Private Sub UserForm_Initialize()
    confirmed = False
    SelectedCompany = ""
    Me.Caption = CaptionText()

    ConfigureStaticControls
    InitCompanyListBox
End Sub

Public Sub SetCompanies(ByVal names As Variant)
    mAll = names
    RefreshList ""
End Sub

Private Sub InitCompanyListBox()
    Dim lst As Object
    Set lst = CompanyListBox()
    If lst Is Nothing Then Exit Sub

    On Error Resume Next
    lst.Visible = True
    lst.ColumnCount = 1
    lst.MultiSelect = fmMultiSelectSingle
    lst.Clear
    ApplyControlFont lst, False
    On Error GoTo 0
End Sub

Private Sub RefreshList(ByVal keyword As String)
    Dim lst As Object
    Set lst = CompanyListBox()
    If lst Is Nothing Then Exit Sub

    lst.Clear
    If Not IsArray(mAll) Then Exit Sub

    Dim kw As String
    kw = Trim$(keyword)

    Dim i As Long
    Dim nm As String
    For i = LBound(mAll) To UBound(mAll)
        nm = CStr(mAll(i))
        If kw = "" Or InStr(1, nm, kw, vbTextCompare) > 0 Then
            lst.AddItem nm
        End If
    Next i

    If lst.ListCount > 0 Then lst.ListIndex = 0
End Sub

Private Sub txtSearch_Change()
    Dim t As Object
    Set t = FindCtl("txtSearch")
    If t Is Nothing Then Exit Sub
    RefreshList CStr(t.Text)
End Sub

Private Sub lstCompanies_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ConfirmSelection
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
        MsgBox SelectPromptText(), vbExclamation
        Exit Sub
    End If

    SelectedCompany = chosen
    confirmed = True
    Me.Hide
End Sub

Private Function GetSelectedCompanyName() As String
    Dim lst As Object
    Set lst = CompanyListBox()
    If lst Is Nothing Then Exit Function
    If lst.ListIndex < 0 Then Exit Function
    GetSelectedCompanyName = CStr(lst.List(lst.ListIndex))
End Function

Private Function CompanyListBox() As Object
    Dim ctrl As Object
    On Error Resume Next
    Set ctrl = Me.Controls("lstCompanies")
    On Error GoTo 0
    If ctrl Is Nothing Then Exit Function
    If TypeName(ctrl) = "ListBox" Then Set CompanyListBox = ctrl
End Function

Private Sub ConfigureStaticControls()
    On Error Resume Next
    Me.Controls("cmdOK").Caption = OkCaptionText()
    Me.Controls("cmdCancel").Caption = CancelCaptionText()
    On Error GoTo 0

    ApplyControlFont Me, False

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox", "CommandButton", "ListBox", "Label"
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

Private Function FindCtl(ByVal nm As String) As Object
    On Error Resume Next
    Set FindCtl = Me.Controls(nm)
    On Error GoTo 0
End Function

' "BIZ UDゴシック"
Private Function FormFontNameText() As String
    Static cached As String
    If cached = "" Then
        cached = "BIZ UD" & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
    End If
    FormFontNameText = cached
End Function

' "施工会社選択"
Private Function CaptionText() As String
    CaptionText = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H4F1A) & ChrW$(&H793E) & _
                  ChrW$(&H9078) & ChrW$(&H629E)
End Function

' "適用"
Private Function OkCaptionText() As String
    OkCaptionText = ChrW$(&H9069) & ChrW$(&H7528)
End Function

' "キャンセル"
Private Function CancelCaptionText() As String
    CancelCaptionText = ChrW$(&H30AD) & ChrW$(&H30E3) & ChrW$(&H30F3) & ChrW$(&H30BB) & ChrW$(&H30EB)
End Function

' "施工会社を選択してください。"
Private Function SelectPromptText() As String
    SelectPromptText = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H4F1A) & ChrW$(&H793E) & _
                       ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3057) & _
                       ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & _
                       ChrW$(&H3044) & ChrW$(&H3002)
End Function

