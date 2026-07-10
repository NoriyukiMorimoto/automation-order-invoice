VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNumericKeypad
   Caption         =   "注文番号入力"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2700
   OleObjectBlob   =   "frmNumericKeypad.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmNumericKeypad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' frmNumericKeypad: 数値入力補助（テンキー）フォーム。
' 基本情報シートの施工会社ブロック27行目（注文番号、mod_OrderTpl_Shared.ORDER_TPL_BLOCK_ORDER_NO_ROW）
' セルをダブルクリックした際に mod_BasicInfoOrderNumberKeypad から起動する想定。
' frmPrefectureSelector を複製し、design-time コントロール（lstCompanies 等）は使用せず
' 非表示化のみ行う。表示欄・数字キー・OK/キャンセルボタンは UserForm_Initialize で
' すべて動的生成する（Controls.Add）。動的ボタンのクリックは clsKeypadBtn（WithEvents）で受け止める。
' 改修履歴: CHANGELOG.md 参照

Public confirmed As Boolean

Private Const FORM_FONT_SIZE As Single = 14
Private Const DISPLAY_FONT_SIZE As Single = 16
Private Const BTN_W As Single = 44
Private Const BTN_H As Single = 30
Private Const BTN_GAP As Single = 6
Private Const GRID_LEFT As Single = 12
Private Const GRID_TOP As Single = 48
Private Const MAX_INPUT_LEN As Long = 12

Private WithEvents mDisplay As MSForms.TextBox
Private mHandlers As Collection
Private mRosaiHandlers As Collection
Private mBuildingRosai As Boolean
Public rosaiKo As Boolean
Public rosaiOtsu As Boolean

Private Sub UserForm_Initialize()
    confirmed = False
    Me.Caption = CaptionText()

    HideDesignTimeControls
    Set mHandlers = New Collection

    BuildDisplay
    BuildKeypadButtons
    BuildActionButtons
    ResizeForm

    mDisplay.Value = ""
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    mDisplay.SetFocus
    mDisplay.SelStart = Len(CStr(mDisplay.Value))
    On Error GoTo 0
End Sub

Public Sub SetValue(ByVal v As String)
    If mDisplay Is Nothing Then Exit Sub
    mDisplay.Value = v
    On Error Resume Next
    mDisplay.SelStart = Len(CStr(mDisplay.Value))
    On Error GoTo 0
End Sub

Public Function GetValue() As String
    If mDisplay Is Nothing Then Exit Function
    GetValue = CStr(mDisplay.Value)
End Function

Private Sub HideDesignTimeControls()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        On Error Resume Next
        ctrl.Visible = False
        ctrl.Enabled = False
        On Error GoTo 0
    Next ctrl
End Sub

Private Sub BuildDisplay()
    Set mDisplay = Me.Controls.Add("Forms.TextBox.1", "txtKeypadDisplay", True)
    With mDisplay
        .Left = GRID_LEFT
        .Top = 10
        .Width = BTN_W * 3 + BTN_GAP * 2
        .Height = 26
        .TextAlign = fmTextAlignRight
        .Font.Name = FormFontNameText()
        .Font.Size = DISPLAY_FONT_SIZE
        .MaxLength = MAX_INPUT_LEN
        .Value = ""
    End With
End Sub

' 7 8 9 / 4 5 6 / 1 2 3 / C 0 BS(バックスペース) の順で配置する
Private Sub BuildKeypadButtons()
    Dim layout As Variant
    layout = Array( _
        Array("7", 0, 0), Array("8", 1, 0), Array("9", 2, 0), _
        Array("4", 0, 1), Array("5", 1, 1), Array("6", 2, 1), _
        Array("1", 0, 2), Array("2", 1, 2), Array("3", 2, 2), _
        Array(ClearKeyText(), 0, 3), Array("0", 1, 3), Array(BackspaceKeyText(), 2, 3))

    Dim i As Long
    For i = LBound(layout) To UBound(layout)
        Dim item As Variant
        item = layout(i)

        Dim label As String
        Dim col As Long
        Dim row As Long
        label = CStr(item(0))
        col = CLng(item(1))
        row = CLng(item(2))

        Dim keyCode As String
        Select Case label
            Case ClearKeyText(): keyCode = "CLR"
            Case BackspaceKeyText(): keyCode = "BS"
            Case Else: keyCode = label
        End Select

        AddKeypadButton "cmdKey" & CStr(i), label, keyCode, _
                        GRID_LEFT + col * (BTN_W + BTN_GAP), _
                        GRID_TOP + row * (BTN_H + BTN_GAP)
    Next i
End Sub

Private Sub BuildActionButtons()
    Dim actionsTop As Single
    Dim gridWidth As Single
    Dim okWidth As Single
    Dim cancelWidth As Single

    actionsTop = GRID_TOP + 4 * (BTN_H + BTN_GAP) + 6
    gridWidth = BTN_W * 3 + BTN_GAP * 2
    okWidth = BTN_W
    cancelWidth = gridWidth - okWidth - BTN_GAP

    AddKeypadButton "cmdKeyOK", OkCaptionText(), "OK", _
                    GRID_LEFT, actionsTop, okWidth
    AddKeypadButton "cmdKeyCancel", CancelCaptionText(), "CANCEL", _
                    GRID_LEFT + okWidth + BTN_GAP, actionsTop, cancelWidth
End Sub

Private Sub AddKeypadButton(ByVal ctrlName As String, ByVal caption As String, ByVal keyCode As String, _
                            ByVal leftPos As Single, ByVal topPos As Single, _
                            Optional ByVal widthOverride As Single = 0)
    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", ctrlName, True)
    With btn
        .Left = leftPos
        .Top = topPos
        .Width = IIf(widthOverride > 0, widthOverride, BTN_W)
        .Height = BTN_H
        .Caption = caption
        .Font.Name = FormFontNameText()
        .Font.Size = FORM_FONT_SIZE
    End With

    Dim handler As clsKeypadBtn
    Set handler = New clsKeypadBtn
    Set handler.Btn = btn
    handler.Key = keyCode
    Set handler.Owner = Me
    mHandlers.Add handler
End Sub

Private Sub ResizeForm()
    Dim contentWidth As Single
    Dim contentHeight As Single
    contentWidth = GRID_LEFT * 2 + BTN_W * 3 + BTN_GAP * 2
    contentHeight = GRID_TOP + 4 * (BTN_H + BTN_GAP) + BTN_H + 24

    Me.Width = contentWidth + 20
    Me.Height = contentHeight + 40
End Sub

' ボタンクリック受付(clsKeypadBtn から呼ばれる)
Public Sub HandleKeypadKey(ByVal keyCode As String)
    Select Case keyCode
        Case "CLR"
            mDisplay.Value = ""
        Case "BS"
            Dim s As String
            s = CStr(mDisplay.Value)
            If Len(s) > 0 Then mDisplay.Value = Left$(s, Len(s) - 1)
        Case "OK"
            ConfirmValue
        Case "CANCEL"
            confirmed = False
            Me.Hide
        Case Else
            If Len(CStr(mDisplay.Value)) < MAX_INPUT_LEN Then
                mDisplay.Value = CStr(mDisplay.Value) & keyCode
            End If
    End Select

    On Error Resume Next
    mDisplay.SetFocus
    mDisplay.SelStart = Len(CStr(mDisplay.Value))
    On Error GoTo 0
End Sub

Private Sub ConfirmValue()
    Dim s As String
    s = Trim$(CStr(mDisplay.Value))

    If s <> "" Then
        If Not IsNumeric(s) Then
            MsgBox InvalidNumberText(), vbExclamation
            Exit Sub
        End If
    End If

    confirmed = True
    Me.Hide
End Sub

Private Sub mDisplay_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 13 ' Enter
            KeyCode = 0
            ConfirmValue
        Case 27 ' Esc
            KeyCode = 0
            confirmed = False
            Me.Hide
    End Select
End Sub

Private Sub mDisplay_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 8 ' 0-9, BackSpace
        Case Else
            KeyAscii = 0
    End Select
End Sub

' "BIZ UDゴシック"
Private Function FormFontNameText() As String
    Static cached As String
    If cached = "" Then
        cached = "BIZ UD" & ChrW$(&H30B4) & ChrW$(&H30B7) & ChrW$(&H30C3) & ChrW$(&H30AF)
    End If
    FormFontNameText = cached
End Function

' "注文番号入力"
Private Function CaptionText() As String
    CaptionText = ChrW$(&H6CE8) & ChrW$(&H6587) & ChrW$(&H756A) & ChrW$(&H53F7) & _
                  ChrW$(&H5165) & ChrW$(&H529B)
End Function

' "適用"
Private Function OkCaptionText() As String
    OkCaptionText = ChrW$(&H9069) & ChrW$(&H7528)
End Function

' "キャンセル"
Private Function CancelCaptionText() As String
    CancelCaptionText = ChrW$(&H30AD) & ChrW$(&H30E3) & ChrW$(&H30F3) & ChrW$(&H30BB) & ChrW$(&H30EB)
End Function

' "C"（クリア）
Private Function ClearKeyText() As String
    ClearKeyText = "C"
End Function

' "BS"（バックスペース。記号グリフの未対応環境を避けるため半角英字表記とする）
Private Function BackspaceKeyText() As String
    BackspaceKeyText = "BS"
End Function

' "数値を入力してください。"
Private Function InvalidNumberText() As String
    InvalidNumberText = ChrW$(&H6570) & ChrW$(&H5024) & ChrW$(&H3092) & ChrW$(&H5165) & _
                        ChrW$(&H529B) & ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H304F) & _
                        ChrW$(&H3060) & ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function

' ===== 労災保険 加入負担 選択モード (frmNumericKeypad 再利用) =====
' New 直後に本メソッドを呼ぶと、UserForm_Initialize が生成したテンキーを撤去し、
' 甲/乙 チェックボックスを上下に並べた選択フォームへ組み替える。
' どちらかにチェックが入った段階で確定してフォームを閉じる(甲/乙は排他)。
Public Sub ConfigureLaborInsuranceMode(ByVal koChecked As Boolean, ByVal otsuChecked As Boolean)
    TearDownKeypadControls
    Me.Caption = LaborInsuranceCaptionText()
    Set mRosaiHandlers = New Collection
    BuildLaborInsuranceCheckboxes koChecked, otsuChecked
    Me.Width = 210
    Me.Height = 116
End Sub

' Initialize が Controls.Add したテンキー系コントロールを撤去する
Private Sub TearDownKeypadControls()
    On Error Resume Next
    Me.Controls.Remove "txtKeypadDisplay"
    Dim i As Long
    For i = 0 To 11
        Me.Controls.Remove "cmdKey" & CStr(i)
    Next i
    Me.Controls.Remove "cmdKeyOK"
    Me.Controls.Remove "cmdKeyCancel"
    Set mHandlers = Nothing
    Set mDisplay = Nothing
    On Error GoTo 0
End Sub

Private Sub BuildLaborInsuranceCheckboxes(ByVal koChecked As Boolean, ByVal otsuChecked As Boolean)
    mBuildingRosai = True

    Dim chkKo As MSForms.CheckBox
    Set chkKo = Me.Controls.Add("Forms.CheckBox.1", "chkRosaiKo", True)
    With chkKo
        .Left = 18
        .Top = 12
        .Width = 150
        .Height = 26
        .Caption = KoText()
        .Font.Name = FormFontNameText()
        .Font.Size = FORM_FONT_SIZE
        .Value = koChecked
    End With

    Dim chkOtsu As MSForms.CheckBox
    Set chkOtsu = Me.Controls.Add("Forms.CheckBox.1", "chkRosaiOtsu", True)
    With chkOtsu
        .Left = 18
        .Top = 44
        .Width = 150
        .Height = 26
        .Caption = OtsuText()
        .Font.Name = FormFontNameText()
        .Font.Size = FORM_FONT_SIZE
        .Value = otsuChecked
    End With

    AddRosaiHandler chkKo, "KO"
    AddRosaiHandler chkOtsu, "OTSU"

    rosaiKo = koChecked
    rosaiOtsu = otsuChecked
    mBuildingRosai = False
End Sub

Private Sub AddRosaiHandler(ByVal chk As MSForms.CheckBox, ByVal keyCode As String)
    Dim handler As clsRosaiCheck
    Set handler = New clsRosaiCheck
    Set handler.Chk = chk
    handler.Key = keyCode
    Set handler.Owner = Me
    mRosaiHandlers.Add handler
End Sub

' clsRosaiCheck から呼ばれる。チェックが入った段階で確定してフォームを閉じる。
' 未チェック(オフ)へ戻す操作では閉じず、状態のみ更新する。
Public Sub HandleRosaiCheck(ByVal keyCode As String, ByVal isChecked As Boolean)
    If mBuildingRosai Then Exit Sub

    If Not isChecked Then
        Select Case keyCode
            Case "KO": rosaiKo = False
            Case "OTSU": rosaiOtsu = False
        End Select
        Exit Sub
    End If

    Select Case keyCode
        Case "KO"
            rosaiKo = True
            rosaiOtsu = False
        Case "OTSU"
            rosaiKo = False
            rosaiOtsu = True
    End Select

    confirmed = True
    Me.Hide
End Sub

' "労災保険 加入負担選択"
Private Function LaborInsuranceCaptionText() As String
    LaborInsuranceCaptionText = ChrW$(&H52B4) & ChrW$(&H707D) & ChrW$(&H4FDD) & ChrW$(&H967A) & _
                                " " & ChrW$(&H52A0) & ChrW$(&H5165) & ChrW$(&H8CA0) & ChrW$(&H62C5) & _
                                ChrW$(&H9078) & ChrW$(&H629E)
End Function

' "甲"
Private Function KoText() As String
    KoText = ChrW$(&H7532)
End Function

' "乙"
Private Function OtsuText() As String
    OtsuText = ChrW$(&H4E59)
End Function
