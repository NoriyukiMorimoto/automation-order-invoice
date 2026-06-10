VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSubconSelector 
   Caption         =   "施工協力会社選択"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmSubconSelector.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSubconSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==========================================================================
'  施工会社選択フォーム  frmSubconSelector (コードビハインド)
'    必要コントロール:
'      txtSearch     … TextBox  (インクリメンタル検索)
'      lstCompanies  … ListBox  (MultiSelect=0 単一選択)
'      cmdOK         … CommandButton (適用)
'      cmdCancel     … CommandButton (キャンセル)
'    フォント14ptは Initialize で全コントロールに設定
'==========================================================================

Public confirmed As Boolean
Public SelectedCompany As String

Private mAll As Variant   ' 全社名(1次元)

Private Sub UserForm_Initialize()
    confirmed = False
    SelectedCompany = ""
    Me.Caption = "施工会社選択"

    ' フォント14pt(フォーム本体・全コントロール)
    On Error Resume Next
    Me.Font.Size = 14
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        ctrl.Font.Size = 14
    Next ctrl
    On Error GoTo 0

    On Error Resume Next
    Me.Controls("cmdOK").Caption = "適用"
    Me.Controls("cmdCancel").Caption = "キャンセル"
    On Error GoTo 0
End Sub

'  呼び出し側から候補社名(1次元配列)を渡す
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
        MsgBox "施工会社を選択してください。", vbExclamation
        Exit Sub
    End If
    ' MultiSelect/BoundColumn に依存しないよう List(ListIndex) で取得(.Value は複数選択時 Null)
    SelectedCompany = CStr(lst.List(lst.ListIndex))
    confirmed = True
    Me.Hide
End Sub

Private Function FindCtl(ByVal nm As String) As Object
    On Error Resume Next
    Set FindCtl = Me.Controls(nm)
    On Error GoTo 0
End Function

