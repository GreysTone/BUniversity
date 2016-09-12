Attribute VB_Name = "Module4"
Option Explicit

Public Sub File4Done(ByVal Id As Integer)
    
    Dim i As Integer
    Dim j As Integer
    Dim LeftVar As String
    Dim RightVar As String
    
    For i = 1 To Player(Id).CountFile(4)        'Start at Main / tank.rtm
        Select Case Player(Id).List4(i).InsID
            Case 0  'IF
                'LeftVar
                If Val(Player(Id).List4(i).V(1)) >= 0 And Val(Player(Id).List4(i).V(1)) <= 9 Then       ' Combo1/Var1 in Range 0-9
                    LeftVar = Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1)
                End If
                If Val(Player(Id).List4(i).V(1)) = 10 Then LeftVar = "F"
                If Val(Player(Id).List4(i).V(1)) = 11 Then LeftVar = "B"
                If Val(Player(Id).List4(i).V(1)) = 12 Then LeftVar = "L"
                If Val(Player(Id).List4(i).V(1)) = 13 Then LeftVar = "R"
                If Val(Player(Id).List4(i).V(1)) = 14 Then LeftVar = "S"
                'RightVar
                If Val(Player(Id).List4(i).V(3)) >= 0 And Val(Player(Id).List4(i).V(3)) <= 9 Then       ' Combo3/Var1 in Range 0-9
                    RightVar = Player(Id).SelfV(Val(Player(Id).List4(i).V(3)) + 1)
                End If
                If Val(Player(Id).List4(i).V(3)) = 10 Then RightVar = "1"
                If Val(Player(Id).List4(i).V(3)) = 11 Then RightVar = "0"
                ' Select Combo 2
                Select Case Val(Player(Id).List4(i).V(2))
                    Case 0
                        If LeftVar = RightVar Then
                            Call IFLink(Id, i, 4, 4)
                        Else
                            Call IFLink(Id, i, 5, 4)
                        End If
                    Case 1
                        If LeftVar > RightVar Then
                            'Yes
                            Call IFLink(Id, i, 4, 4)
                        Else
                            'No
                            Call IFLink(Id, i, 5, 4)
                        End If
                    Case 2
                        If LeftVar >= RightVar Then
                            Call IFLink(Id, i, 4, 4)
                        Else
                            Call IFLink(Id, i, 5, 4)
                        End If
                    Case 3
                        If LeftVar < RightVar Then
                            Call IFLink(Id, i, 4, 4)
                        Else
                            Call IFLink(Id, i, 5, 4)
                        End If
                    Case 4
                        If LeftVar <= RightVar Then
                            Call IFLink(Id, i, 4, 4)
                        Else
                            Call IFLink(Id, i, 5, 4)
                        End If
                    Case 5
                        If LeftVar <> RightVar Then
                            Call IFLink(Id, i, 4, 4)
                        Else
                            Call IFLink(Id, i, 5, 4)
                        End If
                End Select
                
            Case 1  'Move
                If Val(Player(Id).List4(i).V(1)) >= 0 And Val(Player(Id).List4(i).V(1)) <= 9 Then       ' Combo1/Var1 in Range 0-9
                    If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "F" And Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "B" And Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "L" And Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "R" And Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "S" Then
                        Call GameOver("P" & Id & " WrongFormat Error.", Id)
                    Else
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "F" Then Call Move(Id, 0)
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "B" Then Call Move(Id, 1)
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "L" Then Call Move(Id, 2)
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "R" Then Call Move(Id, 3)
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "S" Then Call Move(Id, 4)
                    End If
                Else
                    If Val(Player(Id).List4(i).V(1)) = 10 Then
                        Call Move(Id, 0)
                    Else
                        If Val(Player(Id).List4(i).V(1)) = 11 Then
                            Call Move(Id, 1)
                        Else
                            If Val(Player(Id).List4(i).V(1)) = 12 Then
                                Call Move(Id, 2)
                            Else
                                If Val(Player(Id).List4(i).V(1)) = 13 Then
                                    Call Move(Id, 3)
                                Else
                                    If Val(Player(Id).List4(i).V(1)) = 14 Then
                                        Call Move(Id, 4)
                                    Else
                                        Call GameOver("P" & Id & " RunTime Error.", Id)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Case 2  'Attack
                If Val(Player(Id).List4(i).V(1)) >= 0 And Val(Player(Id).List4(i).V(1)) <= 9 Then       ' Combo1/Var1 in Range 0-9
                    If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "F" And Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "B" And Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "L" And Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) <> "R" Then
                        Call GameOver("P" & Id & " WrongFormat Error.", Id)
                    Else
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "F" Then Call Attack(Id, 0)
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "B" Then Call Attack(Id, 1)
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "L" Then Call Attack(Id, 2)
                        If Player(Id).SelfV(Val(Player(Id).List4(i).V(1)) + 1) = "R" Then Call Attack(Id, 3)
                    End If
                Else
                    If Val(Player(Id).List4(i).V(1)) = 10 Then
                        Call Attack(Id, 0)
                    Else
                        If Val(Player(Id).List4(i).V(1)) = 11 Then
                            Call Attack(Id, 1)
                        Else
                            If Val(Player(Id).List4(i).V(1)) = 12 Then
                                Call Attack(Id, 2)
                            Else
                                If Val(Player(Id).List4(i).V(1)) = 13 Then
                                    Call Attack(Id, 3)
                                Else
                                    Call GameOver("P" & Id & " RunTime Error.", Id)
                                End If
                            End If
                        End If
                    End If
                End If
            Case 3  'GetLockOn
                Player(Id).SelfV(Val(Player(Id).List4(i).R) + 1) = GetLockOn(Id)
            Case 4  'GetFireDirection
                Player(Id).SelfV(Val(Player(Id).List4(i).R) + 1) = GetFireDirection(Id)
            Case 5  'GetFreeWay
                Player(Id).SelfV(Val(Player(Id).List4(i).R) + 1) = GetFreeWay(Id)
            Case 6  'GetFindEnemy
                Player(Id).SelfV(Val(Player(Id).List4(i).R) + 1) = GetFindEnemy(Id)
            Case Else
                Call GameOver("P" & Id & " RunTime Error.", Id)
        End Select
    Next i
    
End Sub

