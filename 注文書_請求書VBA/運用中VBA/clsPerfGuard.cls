VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerfGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsPerfGuard: 重い処理の前後で Application の画面更新/計算方式/イベントを
' 退避・抑制し、スコープ終了時(Class_Terminate)に確実に復元するためのガードクラス。
' 各所に散在していた save/set/restore の定型コードを集約し、エラー経路での
' 復元漏れ(手動計算・イベント無効のまま残る=体感フリーズ)を構造的に防ぐ。
'
' 使い方:
'     Dim g As New clsPerfGuard
'     g.Suspend                 ' 画面更新Off/手動計算/イベントOff
'     ' ...重い処理...
'     g.Restore                 ' 明示復元(省略時もスコープ終了で自動復元)
'
' イベントを止めたくない場合: g.Suspend DisableEvents:=False

Private mScreen As Boolean
Private mCalc As XlCalculation
Private mEvents As Boolean
Private mActive As Boolean

' 現在の状態を退避し、画面更新Off・手動計算・(既定で)イベントOff にする。
Public Sub Suspend(Optional ByVal DisableEvents As Boolean = True)
    If mActive Then Exit Sub
    mScreen = Application.ScreenUpdating
    mCalc = Application.Calculation
    mEvents = Application.EnableEvents
    mActive = True

    On Error Resume Next
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    If DisableEvents Then Application.EnableEvents = False
    On Error GoTo 0
End Sub

' 退避した状態へ復元する(多重呼び出し・未Suspend時は無害)。
Public Sub Restore()
    If Not mActive Then Exit Sub
    mActive = False

    On Error Resume Next
    Application.Calculation = mCalc
    Application.ScreenUpdating = mScreen
    Application.EnableEvents = mEvents
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Restore
End Sub
