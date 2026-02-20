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
          Case "‰®“à•às"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "ŒÒŠO“],”w‹ü,•GL“W")
          Case "ƒgƒCƒŒ"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "ŒÒŠO“],•GL“W,‘ÌŠ²L“W")
          Case "‰®ŠO•às"
              mmtTargetMuscle = PickMMTTarget_WithPriority(mmtMap, "”w‹ü,ŒÒŠO“],•GL“W")
          Case Else
              mmtTargetMuscle = PickMMTTarget(mmtMap)
    End Select

    result("MMT_TargetMuscle") = mmtTargetMuscle
    result("MMT_MinScore") = PickMMTMinScore(mmtMap)
    result("MainCause") = mainCause

    
    result("Activity_Long") = PickActivityLong(needSelf, needFamily, needByDifficulty)

        If Len(Trim$(needSelf)) > 0 Then
          reason = "–{lŠó–]"
        ElseIf Len(Trim$(needFamily)) > 0 Then
          reason = "‰Æ‘°Šó–]"
        Else
          reason = "¢“ï“xãˆÊ"
        End If

    result("Activity_Reason") = reason

    Select Case mainCause
        Case "–ƒáƒ"
        
        If result("MMT_MinScore") <= 2 Then
            fxCore = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚É‚æ‚è"
        Else
            fxCore = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚É‚æ‚è"
        End If

    Select Case result("Activity_Long")
    
        Case "‰®“à•às"
            result("Function_Long") = fxCore & "—§‹rŠúˆÀ’è«Œüã‚ğ}‚éB"
        
        Case "ƒgƒCƒŒ"
            result("Function_Long") = fxCore & "ˆÚæx«Œüã‚ğ}‚éB"
        
        Case "‰®ŠO•às"
            result("Function_Long") = fxCore & "’i·¸~‚ÌˆÀ’è«Œüã‚ğ}‚éB"
        
        Case Else
            result("Function_Long") = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ}‚éB"
            
    End Select


       Case "áu’É"
        result("Function_Long") = "áu’É‚ÌŒyŒ¸‚ğ}‚éB"
       Case "¢“ï“x"
        result("Function_Long") = "‰ºˆ‹@”\‚Ì‘S‘Ì“IŒüã‚ğ}‚éB"
       Case Else
        result("Function_Long") = ""
    End Select

    Select Case mainCause
      Case "–ƒáƒ"
       If result("MMT_MinScore") <= 2 Then
    result("Function_Short") = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ}‚éB"
Else
    result("Function_Short") = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ}‚éB"
End If

      Case "áu’É"
        result("Function_Short") = "áu’É—U”­“®ì‚ÌŒyŒ¸‚¨‚æ‚Ñ•‰‰×’²®‚ğ}‚éB"
      Case "¢“ï“x"
        result("Function_Short") = "å—vƒ{ƒgƒ‹ƒlƒbƒN‹Ø‚Ì‹@”\‰ü‘P‚ğ}‚éB"
      Case Else
        result("Function_Short") = ""
    End Select

    result("Activity_Short") = BuildActivityShort_ByActivity(mainCause, result("Activity_Long"), mmtTargetMuscle, result("MMT_MinScore"))
    result("Participation_Long") = "ˆÚ“®”\—Í‚ÌŒüã‚É‚æ‚è" & result("Activity_Long") & "‚Ì‹@‰ï‚ğ‚Ä‚éó‘Ô‚ğ–Úw‚·B"
      
    shortCore = Replace(result("Activity_Short"), "‚ğ}‚éB", "")
    result("Participation_Short") = shortCore & "‚ğ}‚èA" & result("Activity_Long") & "‚Ì‹@‰ïŠg‘å‚ÉŒü‚¯‚½€”õ‚ğs‚¤B"
    


    
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
        Case "–ƒáƒ"
            BuildActivityShort = activityLong & "‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
            
        Case "áu’É"
            BuildActivityShort = activityLong & "‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
            
        Case Else
            BuildActivityShort = activityLong & "“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"
    End Select
    
    
    
    
End Function


Public Function BuildActivityShort_ByActivity(ByVal mainCause As String, _
                                              ByVal activityLong As String, _
                                              ByVal mmtTargetMuscle As String, _
                                              ByVal mmtMinScore As Double) As String
                                              
                                              
    Select Case activityLong
    
        Case "ƒgƒCƒŒ"
            Select Case mainCause
                Case "–ƒáƒ": BuildActivityShort_ByActivity = "•ÖÀˆÚæ‚Ì–ƒáƒ‘¤x«Œüã‚ğ}‚éB"
                Case "áu’É": BuildActivityShort_ByActivity = "—§‚¿ã‚ª‚è‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case Else:  BuildActivityShort_ByActivity = "•ûŒü“]Š·“®ì‚ÌˆÀ’è‰»‚ğ}‚éB"
            End Select
            
        Case "‰®“à•às"
            Select Case mainCause
  Case "–ƒáƒ"
    If mmtMinScore <= 2 Then
        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚ÌˆÓûkŠl“¾‚ğ’Ê‚¶‚ÄA¶‰E‰×d·‚ÌŒyŒ¸‚ğ}‚éB"
    Else
        BuildActivityShort_ByActivity = mmtTargetMuscle & "‚Ì‹Ø—Í‰ü‘P‚ğ’Ê‚¶‚ÄA¶‰E‰×d·‚ÌŒyŒ¸‚ğ}‚éB"
    End If
                Case "áu’É": BuildActivityShort_ByActivity = "•às‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case "¢“ï“x": BuildActivityShort_ByActivity = "•ûŒü“]Š·‚ÌˆÀ’è«Œüã‚ğ}‚éB"
                Case Else: BuildActivityShort_ByActivity = BuildActivityShort(mainCause, activityLong)
 
            End Select
            
        Case "‰®ŠO•às"
            Select Case mainCause
                Case "–ƒáƒ": BuildActivityShort_ByActivity = "–ƒáƒ‘¤‰ºˆ‚Ìx«Œüã‚ğ}‚éB"
                Case "áu’É": BuildActivityShort_ByActivity = "‰®ŠO•às‚Ìáu’ÉŒyŒ¸‚ğ}‚éB"
                Case Else:  BuildActivityShort_ByActivity = "’i·¸~‚ÌˆÀ’è«Œüã‚ğ}‚éB"
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
    
    ' Å¬ƒXƒRƒA‚ğæ‚é
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
    
    ' “¯—¦‚Ì’†‚Å—Dæ‡‚É‘I‚Ô
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
    
    ' —DæƒŠƒXƒg‚É–³‚¯‚ê‚ÎÅ‰‚ÉŒ©‚Â‚©‚Á‚½Å¬‚ğ•Ô‚·
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
    M.Add "ŒÒŠO“]", 3
    M.Add "•GL“W", 3
    
    Debug.Print PickMMTTarget_WithPriority(M, "•GL“W,ŒÒŠO“]")
End Sub
