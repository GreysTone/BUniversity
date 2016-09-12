Attribute VB_Name = "ModuleFindEnemy"
Option Explicit

Public Function GetFindEnemy(ByVal Playerk As Integer) As String

    Dim NX As Integer
    Dim NY As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ChkOb As Integer
    
    NX = PositionMap.Obc(Playerk).X
    NY = PositionMap.Obc(Playerk).Y
    
    Call Radar(Playerk)
    
    If Playerk = 1 Then
        ChkOb = 2
    Else
        ChkOb = 1
    End If
    
    If Player(Playerk).MapV(1, 3) = ChkOb Then GetFindEnemy = "F"
    If Player(Playerk).MapV(2, 3) = ChkOb Then GetFindEnemy = "F"
    If Player(Playerk).MapV(4, 3) = ChkOb Then GetFindEnemy = "B"
    If Player(Playerk).MapV(5, 3) = ChkOb Then GetFindEnemy = "B"
    If Player(Playerk).MapV(3, 1) = ChkOb Then GetFindEnemy = "L"
    If Player(Playerk).MapV(3, 2) = ChkOb Then GetFindEnemy = "L"
    If Player(Playerk).MapV(3, 4) = ChkOb Then GetFindEnemy = "R"
    If Player(Playerk).MapV(3, 5) = ChkOb Then GetFindEnemy = "R"


End Function
