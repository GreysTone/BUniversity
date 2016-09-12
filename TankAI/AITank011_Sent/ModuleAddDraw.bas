Attribute VB_Name = "ModuleAddDraw"
Option Explicit

Public Sub AddLayerDrawMap(ByVal Ob As Object, ByVal X As Integer, ByVal Y As Integer, ByVal DW As Integer)    'Check Over
    
    Dim i As Integer
    Dim j As Integer
    
    Ob.Cls
    Ob.AutoRedraw = True
    Ob.DrawWidth = DW
    
    Ob.Scale (0, 0)-(X, Y)
    
    ' Draw Physics
    For i = 0 To X - 1
        For j = 0 To Y - 1
            Select Case OrMap.Info(i + 1, j + 1)
                Case 0  'Road
                    'Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
                Case 1  'Stone
                    Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
            End Select
        Next j
    Next i
    
    'Draw PositionMap
    i = PositionMap.Obc(1).X
    j = PositionMap.Obc(1).Y
    Ob.Line (j - 1, i - 1)-(j, i), RGB(Player(1).Color.R, Player(1).Color.G, Player(1).Color.b), BF
    
    i = PositionMap.Obc(2).X
    j = PositionMap.Obc(2).Y
    Ob.Line (j - 1, i - 1)-(j, i), RGB(Player(2).Color.R, Player(2).Color.G, Player(2).Color.b), BF
    
    'Draw LockOn    255 171 171
    For i = 0 To X - 1
        For j = 0 To Y - 1
            Select Case MissileLockMap.Data(i + 1, j + 1).V
                Case 0  'False
                    'Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
                Case 1  'True
                    Ob.Line (j, i)-(j + 1, i + 1), RGB(255, 171, 171), BF
            End Select
        Next j
    Next i
    
    'Draw Missile
    For i = 1 To 10
        If MissileMap.Obc(i).Id <> 0 Then
            Ob.Line (MissileMap.Obc(i).Y - 1, MissileMap.Obc(i).X - 1)-(MissileMap.Obc(i).Y, MissileMap.Obc(i).X), RGB(255, 255, 255), BF
        End If
    Next i
    
End Sub
