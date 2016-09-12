Attribute VB_Name = "ModuleRadar"
Option Explicit

Public Sub Radar(ByVal Playerk As Integer)

    Dim NX As Integer
    Dim NY As Integer
    Dim i As Integer
    Dim j As Integer
    
    NX = PositionMap.Obc(Playerk).X
    NY = PositionMap.Obc(Playerk).Y
    
    For i = -2 To 2
        For j = -2 To 2
            If NX + i < 1 Then
                Player(Playerk).MapV(i + 3, j + 3) = -1
            Else
                If NX + i > OrMap.X Then
                    Player(Playerk).MapV(i + 3, j + 3) = -1
                Else
                    If NY + j < 1 Then
                        Player(Playerk).MapV(i + 3, j + 3) = -1
                    Else
                        If NY + j > OrMap.Y Then
                            Player(Playerk).MapV(i + 3, j + 3) = -1
                        Else
                            Player(Playerk).MapV(i + 3, j + 3) = PositionMap.Data(NX + i, NY + j)
                            If Player(Playerk).MapV(i + 3, j + 3) >= 2 Then Player(Playerk).MapV(i + 3, j + 3) = Player(Playerk).MapV(i + 3, j + 3) + 3  '转换成对抗对象
                            If Player(Playerk).MapV(i + 3, j + 3) = 0 Then
                                Player(Playerk).MapV(i + 3, j + 3) = 1  '转换成实际物理地图
                            Else
                                If Player(Playerk).MapV(i + 3, j + 3) = 1 Then
                                    Player(Playerk).MapV(i + 3, j + 3) = 0
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next j
    Next i
    
    'MsgBox Playerk
    'For i = 1 To 5
    '    For j = 1 To 5
    '        MsgBox i & "," & j & " " & Player(Playerk).MapV(i, j)
    '    Next j
    'Next i

End Sub
