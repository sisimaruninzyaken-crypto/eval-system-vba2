Attribute VB_Name = "modLayoutHeader"
Option Explicit

Public Sub Align_LoadPrevButton_NextToHdrKana(ByVal f As Object)
    Dim hdr As Object
    Dim kana As Object
    Dim btn As Object

    On Error Resume Next
    Set hdr = f.Controls("frHeader")
    If hdr Is Nothing Then Exit Sub

    Set kana = hdr.Controls("txtHdrKana")
    If kana Is Nothing Then Exit Sub

    Set btn = f.Controls("btnLoadPrevCtl")
    If btn Is Nothing Then Exit Sub
    On Error GoTo 0

    btn.Left = hdr.Left + kana.Left + kana.Width + 12
    btn.Top = hdr.Top + kana.Top
End Sub
