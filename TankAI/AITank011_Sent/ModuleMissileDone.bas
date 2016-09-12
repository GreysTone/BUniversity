Attribute VB_Name = "ModuleMissileDone"
Option Explicit

Public Sub MissileDone()

    'Dim i As Integer
    Dim Mis As Integer
    Dim MisX As Integer
    Dim MisY As Integer
    Dim MisP As Integer
    Dim MisD As Integer
    
    For Mis = 1 To 10         '10 Missiles
    'MsgBox MissileMap.Obc(1).Id
        If MissileMap.Obc(Mis).Id <> 0 Then           'Able Missile
        
            'Get Basic Infomation of Missile
            MisX = MissileMap.Obc(Mis).X
            MisY = MissileMap.Obc(Mis).Y
            MisP = MissileMap.Obc(Mis).Belong
            MisD = MissileMap.Obc(Mis).FireDir
        
            'Player is Under Attack?
            Select Case MisD
                Case 0  'Up
                    'Under Attack
                    If PositionMap.Data(MisX, MisY) > 1 Then
                        If MisP + 1 <> PositionMap.Data(MisX, MisY) Then    'Under Not Self Fire
                            If PositionMap.Data(MisX, MisY) = 2 Then
                                'Destroy this Missile
                                Call GameOver("P1 was destoryed.", 1)
                            Else
                                'Destroy this Missile
                                Call GameOver("P2 was destoryed.", 2)
                            End If
                        End If
                    End If
                    'Physic Attack
                    If PositionMap.Data(MisX, MisY) = 0 Then
                        MissileMap.Obc(Mis).Id = 0
                    Else
                        'Move Missile
                        If MisX - 1 = 0 Then    ' Run out of Range
                            MissileMap.Obc(Mis).Id = 0
                        Else
                            MissileMap.Obc(Mis).X = MissileMap.Obc(Mis).X - 1
                        End If
                    End If
                    'Refresh LockOnData
                    MissileLockMap.Data(MisX, MisY).V = False
                Case 1  'Down
                    'Under Attack
                    If PositionMap.Data(MisX, MisY) > 1 Then
                        If MisP + 1 <> PositionMap.Data(MisX, MisY) Then    'Under Not Self Fire
                            If PositionMap.Data(MisX, MisY) = 2 Then
                                'Destroy this Missile
                                Call GameOver("P1 was destoryed.", 1)
                            Else
                                'Destroy this Missile
                                Call GameOver("P2 was destoryed.", 2)
                            End If
                        End If
                    End If
                    'Physic Attack
                    If PositionMap.Data(MisX, MisY) = OrMap.X Then
                        MissileMap.Obc(Mis).Id = 0
                    Else
                        'Move Missile
                        If MisX + 1 = OrMap.X Then     ' Run out of Range
                            MissileMap.Obc(Mis).Id = 0
                        Else
                            MissileMap.Obc(Mis).X = MissileMap.Obc(Mis).X + 1
                        End If
                    End If
                    'Refresh LockOnData
                    MissileLockMap.Data(MisX, MisY).V = False
                Case 2  'Left
                    'Under Attack
                    If PositionMap.Data(MisX, MisY) > 1 Then
                        If MisP + 1 <> PositionMap.Data(MisX, MisY) Then    'Under Not Self Fire
                            If PositionMap.Data(MisX, MisY) = 2 Then
                                'Destroy this Missile
                                Call GameOver("P1 was destoryed.", 1)
                            Else
                                'Destroy this Missile
                                Call GameOver("P2 was destoryed.", 2)
                            End If
                        End If
                    End If
                    'Physic Attack
                    If PositionMap.Data(MisX, MisY) = 0 Then
                        MissileMap.Obc(Mis).Id = 0
                    Else
                        'Move Missile
                        If MisY - 1 = 0 Then    ' Run out of Range
                            MissileMap.Obc(Mis).Id = 0
                        Else
                            MissileMap.Obc(Mis).Y = MissileMap.Obc(Mis).Y - 1
                        End If
                    End If
                    'Refresh LockOnData
                    MissileLockMap.Data(MisX, MisY).V = False
                Case 3  'Right
                    'Under Attack
                    If PositionMap.Data(MisX, MisY) > 1 Then
                        If MisP + 1 <> PositionMap.Data(MisX, MisY) Then    'Under Not Self Fire
                            If PositionMap.Data(MisX, MisY) = 2 Then
                                'Destroy this Missile
                                Call GameOver("P1 was destoryed.", 1)
                            Else
                                'Destroy this Missile
                                Call GameOver("P2 was destoryed.", 2)
                            End If
                        End If
                    End If
                    'Physic Attack
                    If PositionMap.Data(MisX, MisY) = OrMap.Y Then
                        MissileMap.Obc(Mis).Id = 0
                    Else
                        'Move Missile
                        If MisX + 1 = OrMap.Y Then     ' Run out of Range
                            MissileMap.Obc(Mis).Id = 0
                        Else
                            MissileMap.Obc(Mis).Y = MissileMap.Obc(Mis).Y + 1
                        End If
                    End If
                    'Refresh LockOnData
                    MissileLockMap.Data(MisX, MisY).V = False
            End Select
            
        End If
    Next Mis
    
End Sub
