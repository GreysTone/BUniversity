Attribute VB_Name = "ModuleGameInst"
Option Explicit

Public PositionMap As SpcMap
Public MissileMap As SpcMap

Public Sub Move(ByVal Playerk As Integer, ByVal Direction As Integer)
    
    ' 0 - Forward / 1 - Backward / 2 - Left / 3 - Right / 4 -Self
    Dim NX As Integer
    Dim NY As Integer
    Dim RX As Integer
    Dim RY As Integer
    
    Select Case Playerk
        Case 1
            If P1Move = True Then
                Call GameOver("P1 Move in Cannot Move", 1)
                Exit Sub
            End If
        Case 2
            If P2Move = True Then
                Call GameOver("P2 Move in Cannot Move", 2)
                Exit Sub
            End If
    End Select
    
    NX = PositionMap.Obc(Playerk).X
    NY = PositionMap.Obc(Playerk).Y
    
    Select Case Direction
        Case 0  'F
            If NX <= 1 Then
                RX = NX
                RY = NY
            Else
                If PositionMap.Data(NX - 1, NY) <> 0 And PositionMap.Data(NX - 1, NY) <> 2 And PositionMap.Data(NX - 1, NY) <> 3 Then
                    RX = NX - 1
                    RY = NY
                Else
                    RX = NX
                    RY = NY
                End If
            End If
            If Playerk = 1 Then
                frmMain.P1Dir.Caption = "Up"
            Else
                frmMain.P2Dir.Caption = "Up"
            End If
        Case 1  'B
            If NX >= OrMap.X Then
                RX = NX
                RY = NY
            Else
                If PositionMap.Data(NX + 1, NY) <> 0 And PositionMap.Data(NX + 1, NY) <> 2 And PositionMap.Data(NX + 1, NY) <> 3 Then
                    RX = NX + 1
                    RY = NY
                Else
                    RX = NX
                    RY = NY
                End If
            End If
            If Playerk = 1 Then
                frmMain.P1Dir.Caption = "Down"
            Else
                frmMain.P2Dir.Caption = "Down"
            End If
        Case 2  'L
            If NY <= 1 Then
                RX = NX
                RY = NY
            Else
                If PositionMap.Data(NX, NY - 1) <> 0 And PositionMap.Data(NX, NY - 1) <> 2 And PositionMap.Data(NX, NY - 1) <> 3 Then
                    RX = NX
                    RY = NY - 1
                Else
                    RX = NX
                    RY = NY
                End If
            End If
            If Playerk = 1 Then
                frmMain.P1Dir.Caption = "Left"
            Else
                frmMain.P2Dir.Caption = "Left"
            End If
        Case 3  'R
            If NY >= OrMap.Y Then
                RX = NX
                RY = NY
            Else
                If PositionMap.Data(NX, NY + 1) <> 0 And PositionMap.Data(NX, NY + 1) <> 2 And PositionMap.Data(NX, NY + 1) <> 3 Then
                    RX = NX
                    RY = NY + 1
                Else
                    RX = NX
                    RY = NY
                End If
            End If
            If Playerk = 1 Then
                frmMain.P1Dir.Caption = "Right"
            Else
                frmMain.P2Dir.Caption = "Right"
            End If
        Case 4  'S
            RX = NX
            RY = NY
            If Playerk = 1 Then
                frmMain.P1Dir.Caption = "Keep"
            Else
                frmMain.P2Dir.Caption = "Keep"
            End If
    End Select
    
    'Refresh Map Data
    PositionMap.Data(NX, NY) = 1
    PositionMap.Data(RX, RY) = Playerk + 1
    
    'Refresh Map.Obc Data
    PositionMap.Obc(Playerk).X = RX
    PositionMap.Obc(Playerk).Y = RY
    
    'Refresh Move = True
    Select Case Playerk
        Case 1
            P1Move = True
            Call DrawRadarMap(frmMain.P1Sight, 11, 1)
        Case 2
            P2Move = True
            Call DrawRadarMap(frmMain.P2Sight, 11, 2)
    End Select
    
End Sub


Public Sub Attack(ByVal Playerk As Integer, ByVal Direction As Integer)
    
    ' 0 - Forward / 1 - Backward / 2 - Left / 3 - Right
    Dim NX As Integer
    Dim NY As Integer
    
    Dim FreeId As Integer
    Dim W As Integer
    Dim i As Integer
    
    Select Case Playerk
        Case 1
            If P1Attack = True Then
                Call GameOver("P1 Attack in Cannot Attack", 1)
                Exit Sub
            End If
        Case 2
            If P2Attack = True Then
                Call GameOver("P2 Attack in Cannot Attack", 2)
                Exit Sub
            End If
    End Select
    
    NX = PositionMap.Obc(Playerk).X
    NY = PositionMap.Obc(Playerk).Y
    
    'GetFreeId
    W = 0
    For i = 1 To 10
        If MissileMap.Obc(i).Id <> 0 Then
            W = W + 1
        Else
            FreeId = i
            Exit For
        End If
    Next i
    If W >= 10 Then
        FreeId = -1
    End If
    
    If FreeId <> -1 Then                    'From -------------------------------------------------------
        MissileMap.Obc(FreeId).Id = FreeId
        MissileMap.Obc(FreeId).Belong = Playerk
        MissileMap.Obc(FreeId).FireDir = Direction
        MissileMap.Obc(FreeId).X = NX
        MissileMap.Obc(FreeId).Y = NY
        MissileMap.Data(NX, NY) = Playerk
        'MsgBox MissileMap.Obc(1).Y
    
    
    'Refresh LockOnData
    Select Case Direction
        Case 0
            For i = NX To 1 Step -1
                If OrMap.Info(i, NY) = 1 Then
                    Exit For
                Else
                    MissileLockMap.Data(i, NY).V = True                                                                 '留意新旧弹道弹道覆盖问题
                    MissileLockMap.Data(i, NY).Id = MissileMap.Obc(FreeId).Id
                End If
            Next i
            If Playerk = 1 Then
                frmMain.P1FireDir.Caption = "Up"
            Else
                frmMain.P2FireDir.Caption = "Up"
            End If
        Case 1
            For i = NX To OrMap.X
                If OrMap.Info(i, NY) = 1 Then
                    Exit For
                Else
                    MissileLockMap.Data(i, NY).V = True
                    MissileLockMap.Data(i, NY).Id = MissileMap.Obc(FreeId).Id
                End If
            Next i
            If Playerk = 1 Then
                frmMain.P1FireDir.Caption = "Down"
            Else
                frmMain.P2FireDir.Caption = "Down"
            End If
        Case 2
            For i = NY To 1 Step -1
                If OrMap.Info(NX, i) = 1 Then
                    Exit For
                Else
                    MissileLockMap.Data(NX, i).V = True
                    MissileLockMap.Data(NX, i).Id = MissileMap.Obc(FreeId).Id
                End If
            Next i
            If Playerk = 1 Then
                frmMain.P1FireDir.Caption = "Left"
            Else
                frmMain.P2FireDir.Caption = "Left"
            End If
        Case 3
            For i = NY To OrMap.Y
                If OrMap.Info(NX, i) = 1 Then
                    Exit For
                Else
                    MissileLockMap.Data(NX, i).V = True
                    MissileLockMap.Data(NX, i).Id = MissileMap.Obc(FreeId).Id
                End If
            Next i
            If Playerk = 1 Then
                frmMain.P1FireDir.Caption = "Right"
            Else
                frmMain.P2FireDir.Caption = "Right"
            End If
    End Select
    
    'Refresh Attack = True
    Select Case Playerk
        Case 1
            P1Attack = True
        Case 2
            P2Attack = True
    End Select
End If  'To------------------------------------------------------------------------------------------------------
    
End Sub
