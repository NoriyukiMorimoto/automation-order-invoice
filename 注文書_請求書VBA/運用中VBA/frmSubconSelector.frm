Option Explicit

Public confirmed As Boolean
Public SelectedCompany As String

Private mAll As Variant

Private Sub UserForm_Initialize()
    confirmed = False
    SelectedCompany = ""
    Me.Caption = "Ž{H‰ïŽÐ‘I‘ð"

    On Error Resume Next
    Me.Font.Size = 14
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        ctrl.Font.Size = 14
    Next ctrl
    On Error GoTo 0

    On Error Resume Next
    Me.Controls("cmdOK").Caption = "“K—p"
    Me.Controls("cmdCancel").Caption = "ƒLƒƒƒ“ƒZƒ‹"
    On Error GoTo 0
End Sub

Public Sub SetCompanies(ByVal names As Variant)
    mAll = names
    RefreshList ""
End Sub

Private Sub RefreshList(ByVal keyword As String)
    Dim lst As Object
    Set lst = FindCtl("lstCompanies")
    If lst Is Nothing Then Exit Sub

    lst.Clear
    If Not IsArray(mAll) Then Exit Sub

    Dim kw As String
    kw = Trim$(keyword)

    Dim i As Long, nm As String
    For i = LBound(mAll) To UBound(mAll)
        nm = CStr(mAll(i))
        If kw = "" Or InStr(1, nm, kw, vbTextCompare) > 0 Then
            lst.AddItem nm
        End If
    Next i
End Sub

Private Sub txtSearch_Change()
    Dim t As Object
    Set t = FindCtl("txtSearch")
    If t Is Nothing Then Exit Sub
    RefreshList CStr(t.text)
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
    Dim lst As Object
    Set lst = FindCtl("lstCompanies")
    If lst Is Nothing Then Exit Sub
    If lst.ListIndex < 0 Then
        MsgBox "Ž{H‰ïŽÐ‚ð‘I‘ð‚µ‚Ä‚­‚¾‚³‚¢B", vbExclamation
        Exit Sub
    End If
    SelectedCompany = CStr(lst.List(lst.ListIndex))
    confirmed = True
    Me.Hide
End Sub

Private Function FindCtl(ByVal nm As String) As Object
    On Error Resume Next
    Set FindCtl = Me.Controls(nm)
    On Error GoTo 0
End Function
