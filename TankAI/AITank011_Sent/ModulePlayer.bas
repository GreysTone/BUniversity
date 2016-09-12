Attribute VB_Name = "ModulePlayer"
Option Explicit

Public Type RGBs
    R As Integer
    G As Integer
    B As Integer
End Type

Public Type Tanks
    PID As Integer
    Names As String
    Color As RGBs
    
    ListMain(110) As InList
    List1(60) As InList
    List2(60) As InList
    List3(60) As InList
    List4(60) As InList
    List5(60) As InList
    
    MapV(5, 5) As Integer
    SelfV(10) As String
    
    CountFile(6) As Integer
End Type

Public Function CheckReadAiTank(ByVal Name As String, ByVal PlayerID As Integer)    'Check Over
    
    Dim CheckHead As String
    Dim TestFile As String
    Dim i As Integer
    Dim j As Integer
    
    TestFile = Dir(App.Path & "\Tank\" & Name & "\tank.rtm")
    If TestFile = "" Then
        MsgBox "Tank [" & Name & "] 's Main file lose." & vbCrLf & " Please check.", vbOKOnly
        CheckReadAiTank = -1
        Exit Function
    End If
    
    TestFile = Dir(App.Path & "\Tank\" & Name & "\tank.rtb")
    If TestFile = "" Then
        MsgBox "Tank [" & Name & "] 's Basic file lose." & vbCrLf & " Please check.", vbOKOnly
        CheckReadAiTank = -1
        Exit Function
    End If
    
    TestFile = Dir(App.Path & "\Tank\" & Name & "\file1.rf")
    If TestFile = "" Then
        MsgBox "Tank [" & Name & "] 's Logic file_1 lose." & vbCrLf & " Please check.", vbOKOnly
        CheckReadAiTank = -1
        Exit Function
    End If
    
    TestFile = Dir(App.Path & "\Tank\" & Name & "\file2.rf")
    If TestFile = "" Then
        MsgBox "Tank [" & Name & "] 's Logic file_2 lose." & vbCrLf & " Please check.", vbOKOnly
        CheckReadAiTank = -1
        Exit Function
    End If
    
    TestFile = Dir(App.Path & "\Tank\" & Name & "\file3.rf")
    If TestFile = "" Then
        MsgBox "Tank [" & Name & "] 's Logic file_3 lose." & vbCrLf & " Please check.", vbOKOnly
        CheckReadAiTank = -1
        Exit Function
    End If
    
    TestFile = Dir(App.Path & "\Tank\" & Name & "\file4.rf")
    If TestFile = "" Then
        MsgBox "Tank [" & Name & "] 's Logic file_4 lose." & vbCrLf & " Please check.", vbOKOnly
        CheckReadAiTank = -1
        Exit Function
    End If
    
    TestFile = Dir(App.Path & "\Tank\" & Name & "\file5.rf")
    If TestFile = "" Then
        MsgBox "Tank [" & Name & "] 's Logic file_5 lose." & vbCrLf & " Please check.", vbOKOnly
        CheckReadAiTank = -1
        Exit Function
    End If
    
    '# Check Head [RS TANK DATA] / [RS TANK BASIC DATA]
    'RTM
    Open App.Path & "\Tank\" & Name & "\tank.rtm" For Input As #1
        If EOF(1) = True Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Main file." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        End If
        
        Input #1, CheckHead
        If CheckHead <> "[RS TANK DATA]" Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Main file." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        Else
            If EOF(1) = True Then
                MsgBox "Cannot read this Tank [" & Name & "] 's Main file." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                CheckReadAiTank = -1
                Exit Function
            End If
            Input #1, Player(PlayerID).CountFile(0)
                If Player(PlayerID).CountFile(0) > 100 Then
                    MsgBox "Tank [" & Name & "] 's Main file is too big." & vbCrLf & " Please check.", vbOKOnly
                    Close #1
                    CheckReadAiTank = -1
                    Exit Function
                End If
                For i = 1 To Player(PlayerID).CountFile(0)
                    If EOF(1) = True Then
                        MsgBox "Cannot read this Tank [" & Name & "] 's Main file." & vbCrLf & " Please check.", vbOKOnly
                        Close #1
                        CheckReadAiTank = -1
                        Exit Function
                    End If
                    Input #1, Player(PlayerID).ListMain(i).InsID, Player(PlayerID).ListMain(i).R, Player(PlayerID).ListMain(i).V(1), Player(PlayerID).ListMain(i).V(2), Player(PlayerID).ListMain(i).V(3), Player(PlayerID).ListMain(i).V(4), Player(PlayerID).ListMain(i).V(5)
                Next i
        End If
    Close #1
    
    'RTB
    Open App.Path & "\Tank\" & Name & "\tank.rtb" For Input As #1
        If EOF(1) = True Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Basic file." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        End If
        
        Input #1, CheckHead
        If CheckHead <> "[RS TANK BASIC DATA]" Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Basic file." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        Else
            If EOF(1) = True Then
                MsgBox "Cannot read this Tank [" & Name & "] 's Basic file." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                CheckReadAiTank = -1
                Exit Function
            End If
            Input #1, Player(PlayerID).Names
            If EOF(1) = True Then
                MsgBox "Cannot read this Tank [" & Name & "] 's Basic file." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                CheckReadAiTank = -1
                Exit Function
            End If
            Input #1, Player(PlayerID).Color.R, Player(PlayerID).Color.G, Player(PlayerID).Color.B
        End If
    Close #1
    
    'RF1
    Open App.Path & "\Tank\" & Name & "\file1.rf" For Input As #1
        If EOF(1) = True Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_1." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        End If
        
        Input #1, CheckHead
        If CheckHead <> "[RS TANK DATA]" Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_1." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        Else
            If EOF(1) = True Then
                MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_1." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                CheckReadAiTank = -1
                Exit Function
            End If
            Input #1, Player(PlayerID).CountFile(1)
                If Player(PlayerID).CountFile(1) > 50 Then
                    MsgBox "Tank [" & Name & "] 's Logic file is too big." & vbCrLf & " Please check.", vbOKOnly
                    Close #1
                    CheckReadAiTank = -1
                    Exit Function
                End If
                For i = 1 To Player(PlayerID).CountFile(1)
                    If EOF(1) = True Then
                        MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_1." & vbCrLf & " Please check.", vbOKOnly
                        Close #1
                        CheckReadAiTank = -1
                        Exit Function
                    End If
                    Input #1, Player(PlayerID).List1(i).InsID, Player(PlayerID).List1(i).R, Player(PlayerID).List1(i).V(1), Player(PlayerID).List1(i).V(2), Player(PlayerID).List1(i).V(3), Player(PlayerID).List1(i).V(4), Player(PlayerID).List1(i).V(5)
                Next i
        End If
    Close #1
    
    'RF2
    Open App.Path & "\Tank\" & Name & "\file2.rf" For Input As #1
        If EOF(1) = True Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_2." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        End If
        
        Input #1, CheckHead
        If CheckHead <> "[RS TANK DATA]" Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_2." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        Else
            If EOF(1) = True Then
                MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_2." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                CheckReadAiTank = -1
                Exit Function
            End If
            Input #1, Player(PlayerID).CountFile(2)
                If Player(PlayerID).CountFile(2) > 50 Then
                    MsgBox "Tank [" & Name & "] 's Logic file is too big." & vbCrLf & " Please check.", vbOKOnly
                    Close #1
                    CheckReadAiTank = -1
                    Exit Function
                End If
                For i = 1 To Player(PlayerID).CountFile(2)
                    If EOF(1) = True Then
                        MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_2" & vbCrLf & " Please check.", vbOKOnly
                        Close #1
                        CheckReadAiTank = -1
                        Exit Function
                    End If
                    Input #1, Player(PlayerID).List2(i).InsID, Player(PlayerID).List2(i).R, Player(PlayerID).List2(i).V(1), Player(PlayerID).List2(i).V(2), Player(PlayerID).List2(i).V(3), Player(PlayerID).List2(i).V(4), Player(PlayerID).List2(i).V(5)
                Next i
        End If
    Close #1
    
    'RF3
    Open App.Path & "\Tank\" & Name & "\file3.rf" For Input As #1
        If EOF(1) = True Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_3." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        End If
        
        Input #1, CheckHead
        If CheckHead <> "[RS TANK DATA]" Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_3." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        Else
            If EOF(1) = True Then
                MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_3." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                CheckReadAiTank = -1
                Exit Function
            End If
            Input #1, Player(PlayerID).CountFile(3)
                If Player(PlayerID).CountFile(3) > 50 Then
                    MsgBox "Tank [" & Name & "] 's Logic file is too big." & vbCrLf & " Please check.", vbOKOnly
                    Close #1
                    CheckReadAiTank = -1
                    Exit Function
                End If
                For i = 1 To Player(PlayerID).CountFile(3)
                    If EOF(1) = True Then
                        MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_3." & vbCrLf & " Please check.", vbOKOnly
                        Close #1
                        CheckReadAiTank = -1
                        Exit Function
                    End If
                    Input #1, Player(PlayerID).List3(i).InsID, Player(PlayerID).List3(i).R, Player(PlayerID).List3(i).V(1), Player(PlayerID).List3(i).V(2), Player(PlayerID).List3(i).V(3), Player(PlayerID).List3(i).V(4), Player(PlayerID).List3(i).V(5)
                Next i
        End If
    Close #1
    
    'RF4
    Open App.Path & "\Tank\" & Name & "\file4.rf" For Input As #1
        If EOF(1) = True Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_4." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        End If
        
        Input #1, CheckHead
        If CheckHead <> "[RS TANK DATA]" Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_4." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        Else
            If EOF(1) = True Then
                MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_4." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                CheckReadAiTank = -1
                Exit Function
            End If
            Input #1, Player(PlayerID).CountFile(4)
                If Player(PlayerID).CountFile(4) > 50 Then
                    MsgBox "Tank [" & Name & "] 's Logic file is too big." & vbCrLf & " Please check.", vbOKOnly
                    Close #1
                    CheckReadAiTank = -1
                    Exit Function
                End If
                For i = 1 To Player(PlayerID).CountFile(4)
                    If EOF(1) = True Then
                        MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_4." & vbCrLf & " Please check.", vbOKOnly
                        Close #1
                        CheckReadAiTank = -1
                        Exit Function
                    End If
                    Input #1, Player(PlayerID).List4(i).InsID, Player(PlayerID).List4(i).R, Player(PlayerID).List4(i).V(1), Player(PlayerID).List4(i).V(2), Player(PlayerID).List4(i).V(3), Player(PlayerID).List4(i).V(4), Player(PlayerID).List4(i).V(5)
                Next i
        End If
    Close #1
    
    'RF5
    Open App.Path & "\Tank\" & Name & "\file5.rf" For Input As #1
        If EOF(1) = True Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_5." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        End If
        
        Input #1, CheckHead
        If CheckHead <> "[RS TANK DATA]" Then
            MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_5." & vbCrLf & " Please check.", vbOKOnly
            Close #1
            CheckReadAiTank = -1
            Exit Function
        Else
            If EOF(1) = True Then
                MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_5." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                CheckReadAiTank = -1
                Exit Function
            End If
            Input #1, Player(PlayerID).CountFile(5)
                If Player(PlayerID).CountFile(5) > 50 Then
                    MsgBox "Tank [" & Name & "] 's Logic file is too big." & vbCrLf & " Please check.", vbOKOnly
                    Close #1
                    CheckReadAiTank = -1
                    Exit Function
                End If
                For i = 1 To Player(PlayerID).CountFile(5)
                    If EOF(1) = True Then
                        MsgBox "Cannot read this Tank [" & Name & "] 's Logic file_5." & vbCrLf & " Please check.", vbOKOnly
                        Close #1
                        CheckReadAiTank = -1
                        Exit Function
                    End If
                    Input #1, Player(PlayerID).List5(i).InsID, Player(PlayerID).List5(i).R, Player(PlayerID).List5(i).V(1), Player(PlayerID).List5(i).V(2), Player(PlayerID).List5(i).V(3), Player(PlayerID).List5(i).V(4), Player(PlayerID).List5(i).V(5)
                Next i
        End If
    Close #1
    
End Function

