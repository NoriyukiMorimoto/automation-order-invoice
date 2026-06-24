VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "単価一覧表 作成中"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5865
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' === 汎用プログレスフォーム frmProgress ===
' 使い方:
'   frmProgress.ShowProgress 総件数
'   ループ内で frmProgress.UpdateProgress 完了件数, "現在の名称"
'   終了時 frmProgress.CloseProgress

' --- フォームを常に最前面に表示するためのWindows API ---
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#Else
    Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#End If

Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOACTIVATE As Long = &H10

Private mStartTime As Double
Private mTotal As Long

' 表示と初期化を同時に行う
Public Sub ShowProgress(ByVal totalCount As Long)
    mStartTime = Timer
    mTotal = totalCount
    lblProject.Caption = ""
    lblInfo.Caption = "準備中..."
    barFill.width = 0
    Me.Show vbModeless
    SetTopMost
    Me.Repaint
End Sub

' フォームを最前面（TOPMOST）に固定する
Private Sub SetTopMost()
#If VBA7 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
    On Error Resume Next
    hWnd = FindWindowW(StrPtr("ThunderDFrame"), StrPtr(Me.Caption))
    If hWnd <> 0 Then
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    End If
    On Error GoTo 0
End Sub

' done = ここまでに完了した件数
Public Sub UpdateProgress(ByVal done As Long, ByVal currentName As String)
    Dim ratio As Double
    Dim elapsed As Double
    Dim remain As Double

    If mTotal <= 0 Then Exit Sub
    If done < 0 Then done = 0
    If done > mTotal Then done = mTotal

    ratio = done / mTotal
    barFill.width = barBack.width * ratio

    elapsed = Timer - mStartTime
    If elapsed < 0 Then elapsed = elapsed + 86400  ' 日付またぎ補正

    lblProject.Caption = "処理中: " & currentName

    If done = 0 Then
        lblInfo.Caption = CStr(done) & " / " & CStr(mTotal) & "  (0%)  残り時間 計算中..."
    Else
        remain = elapsed / done * (mTotal - done)
        lblInfo.Caption = CStr(done) & " / " & CStr(mTotal) & _
            "  (" & Format(ratio, "0%") & ")  残り約 " & FormatRemain(remain)
    End If

    Me.Repaint
    DoEvents
End Sub

' 進捗の分母(総件数)を途中で拡張する(例: 保存フェーズを加算)
Public Sub SetTotal(ByVal newTotal As Long)
    If newTotal > 0 Then mTotal = newTotal
End Sub

Public Sub CloseProgress()
    On Error Resume Next
    Unload Me
    On Error GoTo 0
End Sub

' 残り秒数を 分/秒 などに整形
Private Function FormatRemain(ByVal seconds As Double) As String
    Dim s As Long
    s = CLng(seconds)
    If s < 0 Then s = 0
    If s < 60 Then
        FormatRemain = CStr(s) & "秒"
    ElseIf s < 3600 Then
        FormatRemain = CStr(s \ 60) & "分" & CStr(s Mod 60) & "秒"
    Else
        FormatRemain = CStr(s \ 3600) & "時間" & CStr((s Mod 3600) \ 60) & "分"
    End If
End Function


