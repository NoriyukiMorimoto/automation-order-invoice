VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKeypadBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsKeypadBtn: frmNumericKeypad が実行時に Controls.Add で生成するボタン1個ごとの
' クリックハンドラ。動的生成コントロールは既定でイベントを拾えないため、
' このクラスに WithEvents で受け止め、Owner(フォーム)の HandleKeypadKey へ橋渡しする。
' フォーム側はインスタンスを Collection に保持し続けること
' (保持しないと WithEvents が解放されクリックが効かなくなる)。
' 改修履歴: CHANGELOG.md 参照

Public WithEvents Btn As MSForms.CommandButton
Public Key As String
Public Owner As Object

Private Sub Btn_Click()
    If Not Owner Is Nothing Then Owner.HandleKeypadKey Key
End Sub
