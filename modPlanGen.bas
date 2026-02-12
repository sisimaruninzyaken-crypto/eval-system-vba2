Attribute VB_Name = "modPlanGen"


Option Explicit

' [変更] BASIC_KEYS_V1 のキー順を固定で定義
Public Function BasicKeysV1() As Variant
    BasicKeysV1 = Split("歩行自立度|異常歩行|認知機能|精神面|care_level_band|primary_condition_cat|comorbidity_cat_list|history_flags|living_type|support_availability|bi_total|bi_low_items|iadl_limits|bed_mobility_band|rom_limit_tags|strength_band|pain_band|pain_site_tags", "|")
End Function

' [変更] 第1段階実装: frmEval から取得できる項目のみ値を埋める
Public Function BuildBasicInputV1() As String
    Dim keys As Variant
    Dim values() As String
    Dim lines() As String
    Dim i As Long

    Dim careLevelBand As String
    Dim livingType As String

    keys = BasicKeysV1()
    ReDim values(LBound(keys) To UBound(keys))
    ReDim lines(LBound(keys) To UBound(keys))

    ' [変更] frmEval ヘッダ項目からの取得（取得不可/空欄時は空文字）
    careLevelBand = GetFrmEvalControlText("cboCare")
    livingType = GetFrmEvalControlText("txtLiving")

    For i = LBound(keys) To UBound(keys)
    Select Case CStr(keys(i))
        Case "care_level_band"
            values(i) = careLevelBand
        Case "living_type"
            values(i) = livingType
        Case "bi_total"
            values(i) = GetFrmEvalControlText("txtBITotal")
        
        Case Else
            values(i) = vbNullString
    End Select

    lines(i) = CStr(keys(i)) & ": " & values(i)
Next i


    ' [修正] 同一キーが重複しても出力は1回に抑制（順序は先勝ちで維持）
    BuildBasicInputV1 = Join(UniqueLinesInOrder(lines), vbCrLf)
End Function

' [変更] frmEval 内の任意コントロール値を安全に取得
Private Function GetFrmEvalControlText(ByVal controlName As String) As String
    Dim target As Object

    Set target = FindControlRecursive(frmEval, controlName)
    If target Is Nothing Then Exit Function

    Select Case TypeName(target)
        Case "TextBox", "ComboBox"
            If LenB(CStr(target.value)) > 0 Then
                GetFrmEvalControlText = CStr(target.value)
            End If
        Case Else
            ' 第1段階では TextBox/ComboBox のみ対応
            GetFrmEvalControlText = vbNullString
    End Select
End Function

' [変更] On Error Resume Next を使わず、再帰でコントロールを探索
Public Function FindControlRecursive(ByVal container As Object, ByVal targetName As String) As Object
    Dim c As Object
    Dim child As Object

    If container Is Nothing Then Exit Function
    
    On Error Resume Next
    Set FindControlRecursive = container.Controls(targetName)
    On Error GoTo 0
    If Not FindControlRecursive Is Nothing Then Exit Function

    

    If Not HasChildControls(container) Then Exit Function

    For Each c In container.Controls
        If StrComp(c.name, targetName, vbTextCompare) = 0 Then
            Set FindControlRecursive = c
            Exit Function
        End If

        If HasChildControls(c) Then
            Set child = FindControlRecursive(c, targetName)
            If Not child Is Nothing Then
                Set FindControlRecursive = child
                Exit Function
            End If
        End If
    Next c
End Function

' [修正] 同じ文字列行の重複出力を防ぐ（順序維持）
Private Function UniqueLinesInOrder(ByRef source() As String) As Variant
    Dim dict As Object
    Dim result() As String
    Dim i As Long
    Dim outIndex As Long
    Dim key As String

    Set dict = CreateObject("Scripting.Dictionary")
    ReDim result(LBound(source) To UBound(source))

    outIndex = LBound(source)
    For i = LBound(source) To UBound(source)
        key = source(i)
        If Not dict.exists(key) Then
            dict.Add key, True
            result(outIndex) = key
            outIndex = outIndex + 1
        End If
    Next i

    ReDim Preserve result(LBound(source) To outIndex - 1)
    UniqueLinesInOrder = result
End Function

' [変更] Controls を持つコンテナ種別のみ再帰対象にする
Private Function HasChildControls(ByVal ctl As Object) As Boolean
    Select Case TypeName(ctl)
        Case "UserForm", "Frame", "MultiPage", "Page"
            HasChildControls = True
        Case Else
            HasChildControls = False
    End Select
End Function




