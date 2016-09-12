Attribute VB_Name = "ModuleGetFreeWay2"
Option Explicit

Public VMap(5, 5) As Integer
Public Count(4) As Integer
Public Last(10) As Integer

'Public Function GetFreeWay(ByVal Playerk As Integer) As String

Public Function GetFreeWay(ByVal Playerk As Integer) As String
    
    Dim i As Integer
    Dim j As Integer
    Dim max As Integer
    Dim Signs As Integer
    Dim ChangeCount As Integer
    
    'Init VMap
    Call InitVMap
    Call LoadP(Playerk)
    
    For i = 1 To 4
        Count(i) = 0
    Next i
    
    max = -1
    Signs = -1
    ChangeCount = 0
    
    Call CalcCount(Playerk)
    
    If Player(Playerk).MapV(2, 3) = Last(Playerk) Then Count(1) = Count(1) / 2
    If Player(Playerk).MapV(4, 3) = Last(Playerk) Then Count(2) = Count(2) / 2
    If Player(Playerk).MapV(3, 2) = Last(Playerk) Then Count(3) = Count(3) / 2
    If Player(Playerk).MapV(3, 4) = Last(Playerk) Then Count(4) = Count(4) / 2
    
    If Player(Playerk).MapV(2, 3) = -1 Or Player(Playerk).MapV(2, 3) = 1 Or Player(Playerk).MapV(2, 3) = 5 Or Player(Playerk).MapV(2, 3) = 6 Then Count(1) = 0
    If Player(Playerk).MapV(4, 3) = -1 Or Player(Playerk).MapV(4, 3) = 1 Or Player(Playerk).MapV(4, 3) = 5 Or Player(Playerk).MapV(4, 3) = 6 Then Count(2) = 0
    If Player(Playerk).MapV(3, 2) = -1 Or Player(Playerk).MapV(3, 2) = 1 Or Player(Playerk).MapV(3, 2) = 5 Or Player(Playerk).MapV(3, 2) = 6 Then Count(3) = 0
    If Player(Playerk).MapV(3, 4) = -1 Or Player(Playerk).MapV(3, 4) = 1 Or Player(Playerk).MapV(3, 4) = 5 Or Player(Playerk).MapV(3, 4) = 6 Then Count(4) = 0
    
    For i = 1 To 4
        If Count(i) > max Then
            max = Count(i)
            Signs = i
        End If
    Next i
    
    Select Case Signs
        Case -1
            GetFreeWay = "S"
        Case 1
            GetFreeWay = "F"
        Case 2
            GetFreeWay = "B"
        Case 3
            GetFreeWay = "L"
        Case 4
            GetFreeWay = "R"
    End Select
    
    'Set Change
    For i = 1 To 3000
        If (Int((3000 - 0 + 1) * Rnd) > 1500) Then
            ChangeCount = ChangeCount + 1
        End If
    Next i
    If ChangeCount > 650 Then
        If GetFreeWay = "F" Then GetFreeWay = Change("F")
        If GetFreeWay = "B" Then GetFreeWay = Change("B")
        If GetFreeWay = "L" Then GetFreeWay = Change("L")
        If GetFreeWay = "R" Then GetFreeWay = Change("R")
    End If
    
    If GetFreeWay = "" Then GetFreeWay = "F"
    
End Function

