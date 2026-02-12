Attribute VB_Name = "modUiQuickInspect"
Sub ListFrmEvalControls()
    Dim c As Control
    For Each c In frmEval.Controls
        Debug.Print TypeName(c), c.name, "Top=" & c.Top, "Left=" & c.Left
    Next
End Sub

