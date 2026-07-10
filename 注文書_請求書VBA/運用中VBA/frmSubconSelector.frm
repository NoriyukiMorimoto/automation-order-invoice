Option Explicit

' OptionButton を Frame 内に配置し、COMPANY_ITEM_HEIGHT で行間を調整する。

Public confirmed As Boolean
Public SelectedCompany As String

Private Const FORM_FONT_SIZE As Single = 14
Private Const COMPANY_ITEM_HEIGHT As Single = 28
Private Const COMPANY_OPTION_LEFT As Single = 8
Private Const COMPANY_OPTION_TOP_MARGIN As Single = 8

Private mAll As Variant
Private mCompanyFrame As Object
Private mVisibleCompanyCount As Long
Private mSCChkHandlers As Collection
Private mSCBtnHandlers As Collection
Private mBuildingSC As Boolean
Private mSCAllowMulti As Boolean
Private mSCNoneKey As String
Private mSCTopKeys As Collection
Private mSCBottomKeys As Collection
Private mSCCtlByKey As Object
Public resultTop As String
Public resultBottom As String

Private Sub UserForm_Initialize()
    confirmed = False
    SelectedCompany = ""
    Me.Caption = CaptionText()

    ConfigureStaticControls
    HideCompanyOptions
End Sub

Public Sub SetCompanies(ByVal names As Variant)
    mAll = names
    RefreshList ""
End Sub

Private Sub RefreshList(ByVal keyword As String)
    Dim frame As Object
    Set frame = CompanyFrame()
    If frame Is Nothing Then Exit Sub

    HideCompanyOptions

    If Not IsArray(mAll) Then Exit Sub

    Dim kw As String
    kw = Trim$(keyword)

    Dim i As Long
    Dim nm As String
    Dim itemIndex As Long
    itemIndex = 0

    For i = LBound(mAll) To UBound(mAll)
        nm = CStr(mAll(i))
        If kw = "" Or InStr(1, nm, kw, vbTextCompare) > 0 Then
            itemIndex = itemIndex + 1
            ShowCompanyOption frame, itemIndex, nm
        End If
    Next i

    mVisibleCompanyCount = itemIndex

    If itemIndex > 0 Then SelectFirstVisibleCompanyOption frame

    On Error Resume Next
    frame.ScrollBars = fmScrollBarsVertical
    frame.ScrollHeight = Application.Max(frame.Height, _
        COMPANY_OPTION_TOP_MARGIN + itemIndex * COMPANY_ITEM_HEIGHT)
    On Error GoTo 0
End Sub

Private Sub txtSearch_Change()
    Dim t As Object
    Set t = FindCtl("txtSearch")
    If t Is Nothing Then Exit Sub
    RefreshList CStr(t.Text)
End Sub

Private Sub cmdOK_Click()
    ConfirmSelection
End Sub

Private Sub cmdCancel_Click()
    confirmed = False
    Me.Hide
End Sub

Private Sub ConfirmSelection()
    Dim chosen As String
    chosen = GetSelectedCompanyName()
    If chosen = "" Then
        MsgBox SelectPromptText(), vbExclamation
        Exit Sub
    End If

    SelectedCompany = chosen
    confirmed = True
    Me.Hide
End Sub

Private Function GetSelectedCompanyName() As String
    Dim frame As Object
    Set frame = CompanyFrame()
    If frame Is Nothing Then Exit Function

    Dim ctrl As Control
    For Each ctrl In frame.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.Visible And ctrl.Value = True Then
                GetSelectedCompanyName = CStr(ctrl.Caption)
                Exit Function
            End If
        End If
    Next ctrl
End Function

Private Sub ShowCompanyOption(ByVal frame As Object, ByVal itemIndex As Long, ByVal companyName As String)
    Dim opt As MSForms.OptionButton
    Set opt = GetOrCreateCompanyOption(frame, itemIndex)

    With opt
        .Caption = companyName
        .Left = COMPANY_OPTION_LEFT
        .Top = COMPANY_OPTION_TOP_MARGIN + (itemIndex - 1) * COMPANY_ITEM_HEIGHT
        .Width = Application.Max(120, frame.Width - (COMPANY_OPTION_LEFT * 2))
        .Height = COMPANY_ITEM_HEIGHT - 8
        .Value = False
        .Visible = True
        ApplyControlFont opt, True
    End With
End Sub

Private Sub SelectFirstVisibleCompanyOption(ByVal frame As Object)
    Dim ctrl As Control
    For Each ctrl In frame.Controls
        If TypeName(ctrl) = "OptionButton" Then
            If ctrl.Visible Then
                ctrl.Value = True
                Exit Sub
            End If
        End If
    Next ctrl
End Sub

Private Sub HideCompanyOptions()
    Dim frame As Object
    Set frame = CompanyFrame()
    If frame Is Nothing Then Exit Sub

    Dim ctrl As Control
    For Each ctrl In frame.Controls
        If TypeName(ctrl) = "OptionButton" Then
            ctrl.Visible = False
            ctrl.Value = False
        End If
    Next ctrl

    mVisibleCompanyCount = 0
End Sub

Private Function CompanyFrame() As Object
    Dim lst As Object
    Set lst = FindCtl("lstCompanies")
    If lst Is Nothing Then Exit Function

    If mCompanyFrame Is Nothing Then
        On Error Resume Next
        Set mCompanyFrame = Me.Controls("fraCompanies")
        On Error GoTo 0

        If mCompanyFrame Is Nothing Then
            Set mCompanyFrame = Me.Controls.Add("Forms.Frame.1", "fraCompanies", True)
            With mCompanyFrame
                .Left = lst.Left
                .Top = lst.Top
                .Width = lst.Width
                .Height = lst.Height
                .Caption = ""
                .BorderStyle = fmBorderStyleSingle
            End With
        End If

        lst.Visible = False
    End If

    Set CompanyFrame = mCompanyFrame
End Function

Private Function GetOrCreateCompanyOption(ByVal frame As Object, ByVal index As Long) As MSForms.OptionButton
    Dim controlName As String
    controlName = "optCompany" & CStr(index)

    On Error Resume Next
    Set GetOrCreateCompanyOption = frame.Controls(controlName)
    On Error GoTo 0

    If GetOrCreateCompanyOption Is Nothing Then
        Set GetOrCreateCompanyOption = frame.Controls.Add("Forms.OptionButton.1", controlName, True)
    End If
End Function

Private Sub ConfigureStaticControls()
    On Error Resume Next
    Me.Controls("cmdOK").Caption = OkCaptionText()
    Me.Controls("cmdCancel").Caption = CancelCaptionText()
    On Error GoTo 0

    ApplyControlFont Me, False

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox", "CommandButton"
                ApplyControlFont ctrl, False
        End Select
    Next ctrl
End Sub

Private Sub ApplyControlFont(ByVal targetControl As Object, ByVal makeBold As Boolean)
    On Error Resume Next
    targetControl.Font.Name = FormFontNameText()
    targetControl.Font.Size = FORM_FONT_SIZE
    targetControl.Font.Bold = makeBold
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

' "施工会社選択"
Private Function CaptionText() As String
    CaptionText = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H4F1A) & ChrW$(&H793E) & _
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

' "施工会社を選択してください。"
Private Function SelectPromptText() As String
    SelectPromptText = ChrW$(&H65BD) & ChrW$(&H5DE5) & ChrW$(&H4F1A) & ChrW$(&H793E) & _
                       ChrW$(&H3092) & ChrW$(&H9078) & ChrW$(&H629E) & ChrW$(&H3057) & _
                       ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & ChrW$(&H3055) & _
                       ChrW$(&H3044) & ChrW$(&H3002)
End Function

' ===== 2セクション選択モード (frmSubconSelector 再利用) =====
' New 直後に本メソッドを呼ぶと、施工会社選択UIを撤去し、
' 上段(有償/無償の排他)と下段(マスタ項目)の2セクション選択フォームへ組み替える。
'   allowMultiBottom=False: 下段は単一選択(行40 支給材料)
'   allowMultiBottom=True : 下段は複数選択可。ただし noneLabel(なし) は排他(行41 貸与品)
' 上下ともに選択後、OK で確定する。結果は resultTop / resultBottom(複数は読点連結)。
Public Sub ConfigureTwoSectionMode(ByVal formCaption As String, _
                                   ByVal topLabel1 As String, ByVal topLabel2 As String, _
                                   ByVal bottomItems As Variant, ByVal allowMultiBottom As Boolean, _
                                   ByVal noneLabel As String, ByVal topInit As String, ByVal bottomInit As Variant)
    mBuildingSC = True
    resultTop = ""
    resultBottom = ""
    mSCAllowMulti = allowMultiBottom
    mSCNoneKey = ""
    Set mSCChkHandlers = New Collection
    Set mSCBtnHandlers = New Collection
    Set mSCTopKeys = New Collection
    Set mSCBottomKeys = New Collection
    Set mSCCtlByKey = CreateObject("Scripting.Dictionary")

    SCHideAllControls
    Me.Caption = formCaption
    Me.Width = 340

    Dim cw As Single
    cw = Me.Width - 44

    ' --- 上段: 有償/無償 (排他) ---
    SCAddCheckbox Me, "T1", topLabel1, 12, 8, cw, (topInit = topLabel1)
    SCAddCheckbox Me, "T2", topLabel2, 12, 34, cw, (topInit = topLabel2)
    mSCTopKeys.Add "T1"
    mSCTopKeys.Add "T2"

    ' --- 下段: マスタ項目 (スクロール枠) ---
    Dim fr As MSForms.Frame
    Set fr = Me.Controls.Add("Forms.Frame.1", "fraBottom", True)
    fr.Left = 10
    fr.Top = 64
    fr.Width = Me.Width - 26
    fr.Height = 150
    fr.Caption = ""
    fr.Font.Name = FormFontNameText()
    fr.Font.Size = FORM_FONT_SIZE

    Dim i As Long
    Dim itemCount As Long
    itemCount = 0
    If IsArray(bottomItems) Then
        For i = LBound(bottomItems) To UBound(bottomItems)
            Dim nm As String
            nm = CStr(bottomItems(i))
            If nm <> "" Then
                Dim key As String
                key = "B" & CStr(itemCount)
                SCAddCheckbox fr, key, nm, 8, 6 + itemCount * 26, fr.Width - 24, SCInitContains(bottomInit, nm)
                mSCBottomKeys.Add key
                If noneLabel <> "" Then
                    If StrComp(nm, noneLabel, vbTextCompare) = 0 Then mSCNoneKey = key
                End If
                itemCount = itemCount + 1
            End If
        Next i
    End If

    On Error Resume Next
    fr.ScrollBars = fmScrollBarsVertical
    fr.ScrollHeight = Application.Max(fr.Height, 12 + itemCount * 26)
    On Error GoTo 0

    ' --- OK / キャンセル ---
    Dim btnTop As Single
    btnTop = 64 + 150 + 8
    SCAddButton "OK", OkCaptionText(), 12, btnTop, 120
    SCAddButton "CANCEL", CancelCaptionText(), 140, btnTop, 120

    Me.Height = btnTop + 34 + 28
    mBuildingSC = False
End Sub

' 施工会社選択UIを含む既存コントロールをすべて隠す
Private Sub SCHideAllControls()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        On Error Resume Next
        ctrl.Visible = False
        On Error GoTo 0
    Next ctrl
End Sub

Private Function SCAddCheckbox(ByVal container As Object, ByVal key As String, ByVal caption As String, _
                               ByVal leftPos As Single, ByVal topPos As Single, ByVal widthPos As Single, _
                               ByVal isChecked As Boolean) As MSForms.CheckBox
    Dim chk As MSForms.CheckBox
    Set chk = container.Controls.Add("Forms.CheckBox.1", "chkSC_" & key, True)
    With chk
        .Left = leftPos
        .Top = topPos
        .Width = widthPos
        .Height = 22
        .Caption = caption
        .Font.Name = FormFontNameText()
        .Font.Size = FORM_FONT_SIZE
        .Value = isChecked
    End With

    Dim handler As clsCheckboxRelay
    Set handler = New clsCheckboxRelay
    Set handler.Chk = chk
    handler.Key = key
    Set handler.Owner = Me
    mSCChkHandlers.Add handler

    Set mSCCtlByKey.Item(key) = chk
    Set SCAddCheckbox = chk
End Function

Private Sub SCAddButton(ByVal key As String, ByVal caption As String, _
                        ByVal leftPos As Single, ByVal topPos As Single, ByVal widthPos As Single)
    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", "cmdSC_" & key, True)
    With btn
        .Left = leftPos
        .Top = topPos
        .Width = widthPos
        .Height = 28
        .Caption = caption
        .Font.Name = FormFontNameText()
        .Font.Size = FORM_FONT_SIZE
    End With

    Dim handler As clsKeypadBtn
    Set handler = New clsKeypadBtn
    Set handler.Btn = btn
    handler.Key = key
    Set handler.Owner = Me
    mSCBtnHandlers.Add handler
End Sub

' bottomInit(配列) に nm が含まれるか
Private Function SCInitContains(ByVal bottomInit As Variant, ByVal nm As String) As Boolean
    If Not IsArray(bottomInit) Then Exit Function
    Dim i As Long
    For i = LBound(bottomInit) To UBound(bottomInit)
        If StrComp(Trim$(CStr(bottomInit(i))), nm, vbTextCompare) = 0 Then
            SCInitContains = True
            Exit Function
        End If
    Next i
End Function

' clsCheckboxRelay から呼ばれる。上段は排他、下段は単一/複数(なし排他)を制御する。
Public Sub OnCheckboxClick(ByVal key As String, ByVal isChecked As Boolean)
    If mBuildingSC Then Exit Sub
    If Not isChecked Then Exit Sub

    If Left$(key, 1) = "T" Then
        SCUncheckOthers mSCTopKeys, key
        Exit Sub
    End If

    ' 下段
    If Not mSCAllowMulti Then
        SCUncheckOthers mSCBottomKeys, key
    Else
        If mSCNoneKey <> "" And key = mSCNoneKey Then
            SCUncheckOthers mSCBottomKeys, key
        ElseIf mSCNoneKey <> "" Then
            SCSetChecked mSCNoneKey, False
        End If
    End If
End Sub

Private Sub SCUncheckOthers(ByVal keys As Collection, ByVal exceptKey As String)
    mBuildingSC = True
    Dim k As Variant
    For Each k In keys
        If CStr(k) <> exceptKey Then SCSetCheckedRaw CStr(k), False
    Next k
    mBuildingSC = False
End Sub

Private Sub SCSetChecked(ByVal key As String, ByVal v As Boolean)
    mBuildingSC = True
    SCSetCheckedRaw key, v
    mBuildingSC = False
End Sub

Private Sub SCSetCheckedRaw(ByVal key As String, ByVal v As Boolean)
    On Error Resume Next
    If mSCCtlByKey.Exists(key) Then mSCCtlByKey.Item(key).Value = v
    On Error GoTo 0
End Sub

' clsKeypadBtn から呼ばれる(OK/キャンセル)
Public Sub HandleKeypadKey(ByVal keyCode As String)
    Select Case keyCode
        Case "OK": SCConfirm
        Case "CANCEL"
            confirmed = False
            Me.Hide
    End Select
End Sub

Private Sub SCConfirm()
    Dim topSel As String
    topSel = SCGetTopSelection()
    Dim bottomSel As String
    bottomSel = SCGetBottomSelection()

    If topSel = "" Or bottomSel = "" Then
        MsgBox SCSelectPromptText(), vbExclamation
        Exit Sub
    End If

    resultTop = topSel
    resultBottom = bottomSel
    confirmed = True
    Me.Hide
End Sub

Private Function SCGetTopSelection() As String
    Dim k As Variant
    For Each k In mSCTopKeys
        If SCIsChecked(CStr(k)) Then
            SCGetTopSelection = SCCaption(CStr(k))
            Exit Function
        End If
    Next k
End Function

' 下段選択を読点(、)連結で返す
Private Function SCGetBottomSelection() As String
    Dim result As String
    Dim k As Variant
    For Each k In mSCBottomKeys
        If SCIsChecked(CStr(k)) Then
            If result <> "" Then result = result & ChrW$(&H3001)
            result = result & SCCaption(CStr(k))
        End If
    Next k
    SCGetBottomSelection = result
End Function

Private Function SCIsChecked(ByVal key As String) As Boolean
    On Error Resume Next
    If mSCCtlByKey.Exists(key) Then SCIsChecked = (mSCCtlByKey.Item(key).Value = True)
    On Error GoTo 0
End Function

Private Function SCCaption(ByVal key As String) As String
    On Error Resume Next
    If mSCCtlByKey.Exists(key) Then SCCaption = CStr(mSCCtlByKey.Item(key).Caption)
    On Error GoTo 0
End Function

' "上段と下段をそれぞれ選択してください。"
Private Function SCSelectPromptText() As String
    SCSelectPromptText = ChrW$(&H4E0A) & ChrW$(&H6BB5) & ChrW$(&H3068) & ChrW$(&H4E0B) & _
                         ChrW$(&H6BB5) & ChrW$(&H3092) & ChrW$(&H305D) & ChrW$(&H308C) & _
                         ChrW$(&H305E) & ChrW$(&H308C) & ChrW$(&H9078) & ChrW$(&H629E) & _
                         ChrW$(&H3057) & ChrW$(&H3066) & ChrW$(&H304F) & ChrW$(&H3060) & _
                         ChrW$(&H3055) & ChrW$(&H3044) & ChrW$(&H3002)
End Function
