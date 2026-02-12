Attribute VB_Name = "modEvalTextBuilder"
Public Sub Preview_NameToHeader()
    Dim f As Object
    Set f = frmEval



    Dim hdr As MSForms.Frame
    Set hdr = f.Controls("frHeader")

    Dim btn As MSForms.CommandButton
    Set btn = f.Controls("cmdClearHeader")

    Dim gap As Single: gap = 10
    Dim pad As Single: pad = 8

    Dim lbl As MSForms.label
    Dim txt As MSForms.TextBox

    '--- create or get header label ---
    On Error Resume Next
    Set lbl = hdr.Controls("lblHdrName")
    On Error GoTo 0
    If lbl Is Nothing Then
    Set lbl = hdr.Controls.Add("Forms.Label.1", "lblHdrName", True)
    lbl.caption = "éÅñº"
    lbl.AutoSize = True
    lbl.Width = lbl.Width + 8   ' Å© Ç±Ç±
End If


    '--- create or get header textbox ---
    On Error Resume Next
    Set txt = hdr.Controls("txtHdrName")
    On Error GoTo 0
    If txt Is Nothing Then
        Set txt = hdr.Controls.Add("Forms.TextBox.1", "txtHdrName", True)
        txt.SpecialEffect = f.Controls("txtName").SpecialEffect
        txt.Font.name = f.Controls("txtName").Font.name
        txt.Font.Size = f.Controls("txtName").Font.Size
        txt.Height = f.Controls("txtName").Height
        txt.Width = f.Controls("txtName").Width
    End If


        txt.IMEMode = fmIMEModeHiragana



    '--- value sync (one-way preview) ---
    txt.Text = f.Controls("txtName").Text

    '--- position: [éÅñº][txt] [cmdClearHeader][cmdSaveHeader][cmdCloseHeader] ---
    txt.Top = btn.Top + (btn.Height - txt.Height) / 2
    lbl.Top = btn.Top + (btn.Height - lbl.Height) / 2

    txt.Left = btn.Left - pad - txt.Width
    lbl.Left = txt.Left - gap - lbl.Width



        '--- create or get header PID label/textbox ---
    Dim lblID As MSForms.label
    Dim txtID As MSForms.TextBox

    On Error Resume Next
    Set lblID = hdr.Controls("lblHdrPID")
    On Error GoTo 0
    If lblID Is Nothing Then
        Set lblID = hdr.Controls.Add("Forms.Label.1", "lblHdrPID", True)
        lblID.caption = "ID"
        lblID.AutoSize = True
        lblID.Width = lblID.Width + 8
    End If

    On Error Resume Next
    Set txtID = hdr.Controls("txtHdrPID")
    On Error GoTo 0
    If txtID Is Nothing Then
        Set txtID = hdr.Controls.Add("Forms.TextBox.1", "txtHdrPID", True)
        txtID.SpecialEffect = f.Controls("txtPID").SpecialEffect
        txtID.Font.name = f.Controls("txtPID").Font.name
        txtID.Font.Size = f.Controls("txtPID").Font.Size
        txtID.Height = f.Controls("txtPID").Height
        txtID.Width = f.Controls("txtPID").Width
    End If

    '--- value sync (one-way preview) ---
    txtID.Text = f.Controls("txtPID").Text

    '--- position: [ID][txt] [éÅñº][txt] [buttons...] ---
    txtID.Top = btn.Top + (btn.Height - txtID.Height) / 2
    lblID.Top = btn.Top + (btn.Height - lblID.Height) / 2

    txtID.Left = lbl.Left - pad - txtID.Width
    lblID.Left = txtID.Left - gap - lblID.Width




End Sub








