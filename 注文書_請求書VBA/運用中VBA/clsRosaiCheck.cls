VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRosaiCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsRosaiCheck: frmNumericKeypad が労災モードで Controls.Add したチェックボックス1個ごとの
' Click ハンドラ。動的生成コントロールは既定でイベントを拾えないため、
' このクラスに WithEvents で受け止め、Owner(フォーム)の HandleRosaiCheck へ橋渡しする。
' フォーム側はインスタンスを Collection に保持し続けること
' (保持しないと WithEvents が解放されクリックが効かなくなる)。
' 改修履歴: CHANGELOG.md 参照

Public WithEvents Chk As MSForms.CheckBox
Public Key As String
Public Owner As Object

Private Sub Chk_Click()
    If Not Owner Is Nothing Then Owner.HandleRosaiCheck Key, (Chk.Value = True)
End Sub
