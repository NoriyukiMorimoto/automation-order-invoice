VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCheckboxRelay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsCheckboxRelay: フォームが実行時に Controls.Add したチェックボックス1個ごとの
' Click ハンドラ。動的生成コントロールは既定でイベントを拾えないため、
' このクラスに WithEvents で受け止め、Owner(フォーム)の OnCheckboxClick へ橋渡しする。
' frmNumericKeypad(排他2択) と frmSubconSelector(2セクション) で共用する。
' フォーム側はインスタンスを Collection に保持し続けること
' (保持しないと WithEvents が解放されクリックが効かなくなる)。
' 改修履歴: CHANGELOG.md 参照

Public WithEvents Chk As MSForms.CheckBox
Public Key As String
Public Owner As Object

Private Sub Chk_Click()
    If Not Owner Is Nothing Then Owner.OnCheckboxClick Key, (Chk.Value = True)
End Sub
