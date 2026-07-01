VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrefectureSelector 
   Caption         =   "施工協力会社選択"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmPrefectureSelector.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPrefectureSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' frmPrefectureSelector: 都道府県選択フォーム（複数選択）
' frmSubconSelector を複製し、既存の lstCompanies リストボックスを
' MultiSelect 化して流用する（動的コントロール生成をしない最軽量構成）。
' lstCompanies / txtSearch / cmdOK / cmdCancel は複製元フォームのコントロールをそのまま使用。

Public confirmed As Boolean

Private Const FORM_FONT_SIZE As Single = 13

Private mAll As Variant
Private mSelectedSet As Object   ' Scripting.Dictionary: 選択済み名（検索フィルタを跨いで保持）

Private Sub UserForm_Initialize()
    confirmed = False
    EnsureSelectedSet
    Me.Caption = CaptionText()
    ConfigureControls
End Sub

Public Sub SetPrefectures(ByVal names As Variant)
    mAll = names
    RefreshList ""
End Sub

Public Sub SetSelectedPrefectures(ByVal names As Variant)
    EnsureSelectedSet
    If Not IsArray(names) Then Exit Sub

    Dim i As Long
    Dim nm As String
    For i = LBound(names) To UBound(names)
        nm = Trim$(CStr(names(i)))
        If nm <> "" Then
            If Not mSelectedSet.Exists(nm) Then mSelectedSet.Add nm, True
        End If
    Next i
End Sub

Public Function GetSelectedPrefectures() As Collection
    CaptureVisibleSelection

    Dim result As Collection
    Set result = New Collection

    If IsArray(mAll) Then
        Dim i As Long
        Dim nm As String
        For i = LBound(mAll) To UBound(mAll)
            nm = CStr(mAll(i))
            If mSelectedSet.Exists(nm) Then result.Add nm
        Next i
    End If

    Set GetSelectedPrefectures = result
End Function

Private Sub ConfigureControls()
    Dim lst As Object
    Set lst = FindCtl("lstCompanies")
    If Not lst Is Nothing Then
        On Error Resume Next
        lst.Visible = True
        lst.MultiSelect = fmMultiSelectMulti
        lst.ListStyle = fmListStyleOption
        On Error GoTo 0
    End If

    On Error Resume Next
    Me.Controls("cmdOK").Caption = OkCaptionText()
    Me.Controls("cmdCancel").Caption = CancelCaptionText()
    On Error GoTo 0

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "ListBox", "TextBox", "CommandButton", "Label"
                ApplyControlFont ctrl
        End Select
    Next ctrl
End Sub

Private Sub RefreshList(ByVal keyword As String)
    Dim lst As Object
    Set lst = FindCtl("lstCompanies")
    If lst Is Nothing Then Exit Sub
    If Not IsArray(mAll) Then Exit Sub

    Dim kw As String
    kw = Trim$(keyword)

    On Error Resume Next
    lst.Clear
    On Error GoTo 0

    Dim i As Long
    Dim nm As String
    For i = LBound(mAll) To UBound(mAll)
        nm = CStr(mAll(i))
        If kw = "" Or InStr(1, nm, kw, vbTextCompare) > 0 Then
            lst.AddItem nm
            If mSelectedSet.Exists(nm) Then lst.selected(lst.ListCount - 1) = True
        End If
    Next i
End Sub

' 表示中リストの選択状態を辞書に取り込む（フィルタ切替・確定時）
Private Sub CaptureVisibleSelection()
    Dim lst As Object
    Set lst = FindCtl("lstCompanies")
    If lst Is Nothing Then Exit Sub
    EnsureSelectedSet

    Dim i As Long
    Dim nm As String
    For i = 0 To lst.ListCount - 1
        nm = CStr(lst.List(i))
        If lst.selected(i) Then
            If Not mSelectedSet.Exists(nm) Then mSelectedSet.Add nm, True
        Else
            If mSelectedSet.Exists(nm) Then mSelectedSet.Remove nm
        End If
    Next i
End Sub

Private Sub txtSearch_Change()
    CaptureVisibleSelection

    Dim t As Object
    Set t = FindCtl("txtSearch")
    If t Is Nothing Then Exit Sub
    RefreshList CStr(t.text)
End Sub

Private Sub cmdOK_Click()
    CaptureVisibleSelection
    If mSelectedSet.Count = 0 Then
        MsgBox SelectPromptText(), vbExclamation
        Exit Sub
    End If
    confirmed = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    confirmed = False
    Me.Hide
End Sub

Private Sub EnsureSelectedSet()
    If mSelectedSet Is Nothing Then
        Set mSelectedSet = CreateObject("Scripting.Dictionary")
        mSelectedSet.CompareMode = vbTextCompare
    End If
End Sub

Private Sub ApplyControlFont(ByVal targetControl As Object)
    On Error Resume Next
    targetControl.Font.Name = FormFontNameText()
    targetControl.Font.Size = FORM_FONT_SIZE
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

' "都道府県選択"
Private Function CaptionText() As String
    CaptionText = ChrW$(&H90FD) & ChrW$(&H9053) & ChrW$(&H5E9C) & ChrW$(&H770C) & _
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

' "都道府県を選択してください。"
Private Function SelectPromptText() As String
    SelectPromptText = ChrW$(&H90FD) & ChrW$(&H9053) & ChrW$(&H5E9C) & ChrW$(&H770C) & _
                       ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3057) & _
                       ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & _
                       ChrW$(&H3044) & ChrW$(&H3002)
End Function


