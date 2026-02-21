Attribute VB_Name = "modKinrenPlanBasicCore"
Public Function BuildBasicPlanStructure(ByVal mainCause As String, _
                                        ByVal needSelf As String, _
                                        ByVal needFamily As String, _
                                        ByVal needByDifficulty As String, _
                                        ByVal mmtMap As Object) As Object

                                        
    Dim result As Object
    Dim reason As String
    Dim shortCore As String
    Dim mmtTargetMuscle As String
    Dim fxCore As String

    Set result = CreateObject("Scripting.Dictionary")

    Select Case result("Activity_Long")
          Case "屋内歩行"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "股外転,背屈,膝伸展")
          Case "トイレ"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "股外転,膝伸展,体幹伸展")
          Case "屋外歩行"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "背屈,股外転,膝伸展")
          Case Else
              mmtTargetMuscle = PickMMTTarget(mmtMap)
    End Select

    result("MMT_TargetMuscle") = mmtTargetMuscle
    result("MMT_MinScore") = PickMMTMinScore(mmtMap)
    result("MainCause") = mainCause

    
    result("Activity_Long") = PickActivityLong(needSelf, needFamily, needByDifficulty)

        If Len(Trim$(needSelf)) > 0 Then
          reason = "本人希望"
        ElseIf Len(Trim$(needFamily)) > 0 Then
          reason = "家族希望"
        Else
          reason = "困難度上位"
        End If

    result("Activity_Reason") = reason

    Select Case mainCause
        Case "麻痺"
        
        If result("MMT_MinScore") <= 2 Then
            fxCore = mmtTargetMuscle & "の随意収縮獲得により"
        Else
            fxCore = mmtTargetMuscle & "の筋力改善により"
        End If

    Select Case result("Activity_Long")
    
        Case "屋内歩行"
            result("Function_Long") = fxCore & "立脚期安定性向上を図る。"
        
        Case "トイレ"
            result("Function_Long") = fxCore & "移乗時支持性向上を図る。"
        
        Case "屋外歩行"
            result("Function_Long") = fxCore & "段差昇降時の安定性向上を図る。"
        
        Case Else
            result("Function_Long") = mmtTargetMuscle & "の筋力改善を図る。"
            
    End Select


       Case "疼痛"
        result("Function_Long") = "疼痛の軽減を図る。"
       Case "困難度"
        result("Function_Long") = "下肢機能の全体的向上を図る。"
       Case Else
        result("Function_Long") = ""
    End Select

    Select Case mainCause
      Case "麻痺"
       If result("MMT_MinScore") <= 2 Then
    result("Function_Short") = mmtTargetMuscle & "の随意収縮獲得を図る。"
Else
    result("Function_Short") = mmtTargetMuscle & "の筋力改善を図る。"
End If

      Case "疼痛"
        result("Function_Short") = "疼痛誘発動作の軽減および負荷調整を図る。"
      Case "困難度"
        result("Function_Short") = "主要ボトルネック筋の機能改善を図る。"
      Case Else
        result("Function_Short") = ""
    End Select

    result("Activity_Short") = BuildActivityShort_ByActivity(mainCause, result("Activity_Long"), mmtTargetMuscle, result("MMT_MinScore"))
    result("Participation_Long") = "移動能力の向上により" & result("Activity_Long") & "の機会を持てる状態を目指す。"
      
    shortCore = Replace(result("Activity_Short"), "を図る。", "")
    result("Participation_Short") = shortCore & "を図り、" & result("Activity_Long") & "の機会拡大に向けた準備を行う。"
    


    
    Set BuildBasicPlanStructure = result
    
End Function


Public Function PickActivityLong(ByVal needSelf As String, ByVal needFamily As String, ByVal needByDifficulty As String) As String
    If Len(Trim$(needSelf)) > 0 Then
        PickActivityLong = needSelf
        Exit Function
    End If
    
    If Len(Trim$(needFamily)) > 0 Then
        PickActivityLong = needFamily
        Exit Function
    End If
    
    PickActivityLong = needByDifficulty
End Function


Public Function BuildActivityShort(ByVal mainCause As String, ByVal activityLong As String) As String
    
    Select Case mainCause
        Case "麻痺"
            BuildActivityShort = activityLong & "時の麻痺側支持性向上を図る。"
            
        Case "疼痛"
            BuildActivityShort = activityLong & "時の疼痛軽減を図る。"
            
        Case Else
            BuildActivityShort = activityLong & "動作の安定化を図る。"
    End Select
    
    
    
    
End Function


Public Function BuildActivityShort_ByActivity(ByVal mainCause As String, _
                                              ByVal activityLong As String, _
                                              ByVal mmtTargetMuscle As String, _
                                              ByVal mmtMinScore As Double) As String
                                              
                                              
    Select Case activityLong
    
        Case "トイレ"
            Select Case mainCause
                Case "麻痺": BuildActivityShort_ByActivity = "便座移乗時の麻痺側支持性向上を図る。"
                Case "疼痛": BuildActivityShort_ByActivity = "立ち上がり時の疼痛軽減を図る。"
                Case Else:  BuildActivityShort_ByActivity = "方向転換動作の安定化を図る。"
            End Select
            
        Case "屋内歩行"
            Select Case mainCause
  Case "麻痺"
    If mmtMinScore <= 2 Then
        BuildActivityShort_ByActivity = mmtTargetMuscle & "の随意収縮獲得を通じて、左右荷重差の軽減を図る。"
    Else
        BuildActivityShort_ByActivity = mmtTargetMuscle & "の筋力改善を通じて、左右荷重差の軽減を図る。"
    End If
                Case "疼痛": BuildActivityShort_ByActivity = "歩行時の疼痛軽減を図る。"
                Case "困難度": BuildActivityShort_ByActivity = "方向転換時の安定性向上を図る。"
                Case Else: BuildActivityShort_ByActivity = BuildActivityShort(mainCause, activityLong)
 
            End Select
            
        Case "屋外歩行"
            Select Case mainCause
                Case "麻痺": BuildActivityShort_ByActivity = "麻痺側下肢の支持性向上を図る。"
                Case "疼痛": BuildActivityShort_ByActivity = "屋外歩行時の疼痛軽減を図る。"
                Case Else:  BuildActivityShort_ByActivity = "段差昇降時の安定性向上を図る。"
            End Select
    
            
            
        Case Else
            BuildActivityShort_ByActivity = BuildActivityShort(mainCause, activityLong)
            
    End Select
    
End Function



Public Function DumpBasicPlan(ByVal plan As Object) As String
    Dim keys As Variant, i As Long, s As String
    
    keys = Array( _
        "MainCause", _
        "Activity_Long", _
        "Activity_Reason", _
        "Function_Long", _
        "Function_Short", _
        "Activity_Short", _
        "Participation_Long", _
        "Participation_Short" _
    )
    
    For i = LBound(keys) To UBound(keys)
        If plan.exists(keys(i)) Then
            s = s & plan(keys(i)) & vbCrLf
        Else
            s = s & "" & vbCrLf
        End If
    Next i
    
    DumpBasicPlan = s
End Function


Public Function DumpBasicGoalsOnly(ByVal plan As Object) As String
    Dim keys As Variant, i As Long, s As String
    
    keys = Array( _
    "Function_Short", _
    "Function_Long", _
    "Activity_Short", _
    "Activity_Long", _
    "Participation_Short", _
    "Participation_Long" _
)
    
    For i = LBound(keys) To UBound(keys)
        If plan.exists(keys(i)) Then
            s = s & plan(keys(i)) & vbCrLf
        Else
            s = s & "" & vbCrLf
        End If
    Next i
    
    DumpBasicGoalsOnly = s
End Function


Public Function PickMMTTarget(ByVal mmtMap As Object) As String

    Dim k As Variant
    Dim bestMuscle As String
    Dim bestScore As Double
    
    bestMuscle = ""
    bestScore = 9999
    
    For Each k In mmtMap.keys
        If IsNumeric(mmtMap(k)) Then
            If CDbl(mmtMap(k)) < bestScore Then
                bestScore = CDbl(mmtMap(k))
                bestMuscle = CStr(k)
            End If
        End If
    Next k
    
    PickMMTTarget = bestMuscle
End Function




Public Function PickMMTTarget_FromPairs(ParamArray pairs() As Variant) As String
    Dim d As Object
    Dim i As Long
    
    Set d = CreateObject("Scripting.Dictionary")
    
    i = LBound(pairs)
    Do While i <= UBound(pairs) - 1
        d(CStr(pairs(i))) = CDbl(pairs(i + 1))
        i = i + 2
    Loop
    
    PickMMTTarget_FromPairs = PickMMTTarget(d)
End Function




Public Function BuildBasicPlan_FromPairs( _
    ByVal mainCause As String, _
    ByVal needSelf As String, _
    ByVal needFamily As String, _
    ByVal needByDifficulty As String, _
    ParamArray mmtPairs() As Variant) As Object
    
    Dim d As Object
    Dim i As Long
    
    Set d = CreateObject("Scripting.Dictionary")
    
    i = LBound(mmtPairs)
    Do While i <= UBound(mmtPairs) - 1
        d(CStr(mmtPairs(i))) = CDbl(mmtPairs(i + 1))
        i = i + 2
    Loop
    
    Set BuildBasicPlan_FromPairs = _
        BuildBasicPlanStructure(mainCause, needSelf, needFamily, needByDifficulty, d)
End Function



Public Function PickMMTMinScore(ByVal mmtMap As Object) As Double
    Dim k As Variant
    Dim bestScore As Double
    
    bestScore = 9999
    
    For Each k In mmtMap.keys
        If IsNumeric(mmtMap(k)) Then
            If CDbl(mmtMap(k)) < bestScore Then
                bestScore = CDbl(mmtMap(k))
            End If
        End If
    Next k
    
    If bestScore = 9999 Then bestScore = 0
    PickMMTMinScore = bestScore
End Function




Public Function PickMMTTarget_WithPriority(ByVal mmtMap As Object, ByVal priorityCsv As String) As String
    Dim pri() As String, i As Long
    Dim best As String, bestScore As Double
    Dim k As Variant, sc As Double
    
    best = ""
    bestScore = 9999
    
    ' 最小スコアを取る
    For Each k In mmtMap.keys
        If IsNumeric(mmtMap(k)) Then
            sc = CDbl(mmtMap(k))
            If sc < bestScore Then bestScore = sc
        End If
    Next k
    
    If bestScore = 9999 Then
        PickMMTTarget_WithPriority = ""
        Exit Function
    End If
    
    ' 同率の中で優先順に選ぶ
    pri = Split(priorityCsv, ",")
    For i = LBound(pri) To UBound(pri)
        If mmtMap.exists(Trim$(pri(i))) Then
            If IsNumeric(mmtMap(Trim$(pri(i)))) Then
                If CDbl(mmtMap(Trim$(pri(i)))) = bestScore Then
                    PickMMTTarget_WithPriority = Trim$(pri(i))
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' 優先リストに無ければ最初に見つかった最小を返す
    For Each k In mmtMap.keys
        If IsNumeric(mmtMap(k)) Then
            If CDbl(mmtMap(k)) = bestScore Then
                PickMMTTarget_WithPriority = CStr(k)
                Exit Function
            End If
        End If
    Next k
End Function




Public Sub Test_PickMMTTarget_WithPriority_Tie()
    Set M = CreateObject("Scripting.Dictionary")
    M.Add "股外転", 3
    M.Add "膝伸展", 3
    
    Debug.Print PickMMTTarget_WithPriority(M, "膝伸展,股外転")
End Sub


' encoding test
