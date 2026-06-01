VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AllVenderSelection 
   Caption         =   "工事番号選択（最新順）"
   ClientHeight    =   13110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10635
   OleObjectBlob   =   "AllVenderSelection.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Option Explicit

Private mVendorData As Variant
Private WithEvents mSelectButton As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Me.Caption = AllVendorSelectionCaptionText()

    Set mSelectButton = SelectCommandButton()
    If Not mSelectButton Is Nothing Then mSelectButton.Caption = SelectButtonText()

    SetupListView
    mVendorData = mod_VendorMaster.GetAllVendorSelectionData()
    RefreshList ""
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    SearchTextBox().SetFocus
    On Error GoTo 0
End Sub

Private Sub SetupListView()
    Dim vendorList As Object
    Set vendorList = VendorListView()
    If vendorList Is Nothing Then Exit Sub

    With vendorList
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideColumnHeaders = False
        .MultiSelect = False
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , BranchHeaderText(), 110
        .ColumnHeaders.Add , , VendorNameHeaderText(), 260
        .ColumnHeaders.Add , , VendorCodeHeaderText(), 120
    End With
End Sub

Private Sub TextBox1_Change()
    Dim searchBox As Object
    Set searchBox = SearchTextBox()
    If searchBox Is Nothing Then Exit Sub

    RefreshList searchBox.text
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then ApplySelectedVendor
End Sub

Private Sub ListView1_DblClick()
    ApplySelectedVendor
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then ApplySelectedVendor
End Sub

Private Sub mSelectButton_Click()
    ApplySelectedVendor
End Sub

Private Sub RefreshList(ByVal keyword As String)
    Dim vendorList As Object
    Set vendorList = VendorListView()
    If vendorList Is Nothing Then Exit Sub

    vendorList.ListItems.Clear
    If IsEmpty(mVendorData) Then Exit Sub

    keyword = NormalizeFormText(keyword)

    Dim i As Long
    Dim searchText As String
    Dim item As Object
    For i = LBound(mVendorData, 1) To UBound(mVendorData, 1)
        searchText = NormalizeFormText(CStr(mVendorData(i, 1)) & " " & CStr(mVendorData(i, 2)) & " " & CStr(mVendorData(i, 3)))
        If keyword = "" Or InStr(1, searchText, keyword, vbTextCompare) > 0 Then
            Set item = vendorList.ListItems.Add(, , CStr(mVendorData(i, 1)))
            item.SubItems(1) = CStr(mVendorData(i, 2))
            item.SubItems(2) = CStr(mVendorData(i, 3))
            item.Tag = CStr(mVendorData(i, 1)) & vbTab & CStr(mVendorData(i, 2))
        End If
    Next i
End Sub

Private Sub ApplySelectedVendor()
    Dim vendorList As Object
    Set vendorList = VendorListView()
    If vendorList Is Nothing Then Exit Sub
    If vendorList.SelectedItem Is Nothing Then Exit Sub

    Dim values As Variant
    values = Split(CStr(vendorList.SelectedItem.Tag), vbTab)
    If UBound(values) < 1 Then Exit Sub

    mod_VendorMaster.ApplyVendorSelection CStr(values(0)), CStr(values(1))
    Unload Me
End Sub

Private Function SearchTextBox() As Object
    Set SearchTextBox = FindControlByNameOrType("TextBox1", "TextBox")
End Function

Private Function SelectCommandButton() As Object
    Set SelectCommandButton = FindControlByNameOrType("CommandButton1", "CommandButton")
End Function

Private Function VendorListView() As Object
    Set VendorListView = FindControlByNameOrType("ListView1", "ListView")
End Function

Private Function FindControlByNameOrType(ByVal controlName As String, ByVal controlTypeName As String) As Object
    On Error Resume Next
    Set FindControlByNameOrType = Me.Controls(controlName)
    On Error GoTo 0
    If Not FindControlByNameOrType Is Nothing Then Exit Function

    Dim ctrl As Object
    For Each ctrl In Me.Controls
        If StrComp(TypeName(ctrl), controlTypeName, vbTextCompare) = 0 Then
            Set FindControlByNameOrType = ctrl
            Exit Function
        End If
    Next ctrl
End Function

Private Function NormalizeFormText(ByVal value As String) As String
    NormalizeFormText = Trim$(Replace$(Replace$(Replace$(value, vbCr, ""), vbLf, ""), vbTab, " "))
    NormalizeFormText = Replace$(NormalizeFormText, ChrW$(&H3000), " ")
    Do While InStr(NormalizeFormText, "  ") > 0
        NormalizeFormText = Replace$(NormalizeFormText, "  ", " ")
    Loop
    NormalizeFormText = Trim$(NormalizeFormText)
End Function

Private Function AllVendorSelectionCaptionText() As String
    AllVendorSelectionCaptionText = ChrW$(&H5168) & ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H9078) & ChrW$(&H629E)
End Function

Private Function SelectButtonText() As String
    SelectButtonText = ChrW$(&H9078) & ChrW$(&H629E)
End Function

Private Function BranchHeaderText() As String
    BranchHeaderText = ChrW$(&H652F) & ChrW$(&H5E97)
End Function

Private Function VendorNameHeaderText() As String
    VendorNameHeaderText = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H540D)
End Function

Private Function VendorCodeHeaderText() As String
    VendorCodeHeaderText = ChrW$(&H696D) & ChrW$(&H8005) & ChrW$(&H30B3) & ChrW$(&H30FC) & ChrW$(&H30C9)
End Function
