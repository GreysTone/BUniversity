Attribute VB_Name = "ModuleMap"
Option Explicit

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type Position2_forObject
    Id As Integer
    X As Integer
    Y As Integer
    Belong As Integer
    FireDir As Integer
End Type

Public Type Map
    Names As String
    X As Integer
    Y As Integer
    Players As Integer
    Info(110, 110) As Integer
    Pos(10) As Position
    MsPos As Boolean
End Type

Public Type SpcMap
    Data(110, 110) As Integer
    Obc(10) As Position2_forObject
End Type

Public Type Boolean2
    V As Boolean
    Id As Integer
End Type

Public Type SpcMap2
    Data(110, 110) As Boolean
End Type

Public Type SpcMap3
    Data(110, 110) As Boolean2
End Type

Public Sub DrawMap(ByVal Ob As Object, ByVal X As Integer, ByVal Y As Integer, ByVal DW As Integer)    'Check Over

    Dim i As Integer
    Dim j As Integer
    
    Ob.Cls
    Ob.AutoRedraw = True
    Ob.DrawWidth = DW
    
    Ob.Scale (0, 0)-(X, Y)
    For i = 0 To X - 1
        For j = 0 To Y - 1
            Select Case OrMap.Info(i + 1, j + 1)
                Case 0  'Road
                    'Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
                Case 1  'Stone
                    Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
                Case 5  'P1
                    Ob.Line (j, i)-(j + 1, i + 1), RGB(Player(1).Color.R, Player(1).Color.G, Player(1).Color.B), BF
                Case 6  'P2
                     Ob.Line (j, i)-(j + 1, i + 1), RGB(Player(2).Color.R, Player(2).Color.G, Player(2).Color.B), BF
                Case Else
                    Ob.Line (j, i)-(j + 1, i + 1), vbRed, BF
            End Select
        Next j
    Next i
    
    'Ob.Scale (0, 0)-(2 * X, 2 * Y)
    'For i = 0 To 2 * X Step 2
    '    For j = 0 To 2 * Y Step 2
    '        Select Case OrMap.Info(i / 2 + 1, j / 2 + 1)
    '            Case 0  'Road
    '                'Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
    '            Case 1  'Stone
    '                Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
    '            Case 5  'P1
    '                Ob.Line (j, i)-(j + 1, i + 1), RGB(Player(1).Color.R, Player(1).Color.G, Player(1).Color.B), BF
    '            Case 6  'P2
    '                 Ob.Line (j, i)-(j + 1, i + 1), RGB(Player(2).Color.R, Player(2).Color.G, Player(2).Color.B), BF
    '            Case Else
    '                Ob.Line (j, i)-(j + 1, i + 1), vbRed, BF
    '        End Select
    '    Next j
    'Next i
    
End Sub

Public Sub DrawRadarMap(ByVal Ob As Object, ByVal DW As Integer, ByVal Playerk)    'Check Over

    Dim i As Integer
    Dim j As Integer
    
    Call Radar(Playerk)
    
    Ob.Cls
    Ob.AutoRedraw = True
    Ob.DrawWidth = DW
    
    Ob.Scale (0, 0)-(5, 5)
    For i = 0 To 4
        For j = 0 To 4
            Select Case Player(Playerk).MapV(i + 1, j + 1)
                Case -1 'Null
                    Ob.Line (j, i)-(j + 1, i + 1), vbBlack, BF
                Case 0  'Road
                    'Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
                Case 1  'Stone
                    Ob.Line (j, i)-(j + 1, i + 1), &HE0E0E0, BF
                Case 5  'P1
                    Ob.Line (j, i)-(j + 1, i + 1), RGB(Player(1).Color.R, Player(1).Color.G, Player(1).Color.B), BF
                Case 6  'P2
                     Ob.Line (j, i)-(j + 1, i + 1), RGB(Player(2).Color.R, Player(2).Color.G, Player(2).Color.B), BF
                Case Else
                    Ob.Line (j, i)-(j + 1, i + 1), vbRed, BF
            End Select
        Next j
    Next i
    
End Sub










'############################################################################
'
'        OrMap
'        0 - Road   1 - Stone  2 - Wall   3 - River  4 - Grass  5 - P1  6 - P2 ...
'
'
'         PositionMap
'         0 - OrMapBit   1 - Road   2 - P1  3 - P2
'
'        MissileMap
'        1 - P1   2 - P2
'
'       MapView
'       -1 - OutRange   0 - Road  1- Stone  5-P1  6-P2
'

