Attribute VB_Name = "ModuleGetFreeWay2Link"
Option Explicit

Public Sub InitVMap()

    VMap(1, 1) = 2
    VMap(1, 2) = 2
    VMap(1, 3) = 2
    VMap(1, 4) = 2
    VMap(1, 5) = 2
    
    VMap(2, 1) = 2
    VMap(2, 2) = 1
    VMap(2, 3) = 1
    VMap(2, 4) = 1
    VMap(2, 5) = 2
    
    VMap(3, 1) = 2
    VMap(3, 2) = 1
    VMap(3, 3) = 5
    VMap(3, 4) = 1
    VMap(3, 5) = 2
    
    VMap(4, 1) = 2
    VMap(4, 2) = 1
    VMap(4, 3) = 1
    VMap(4, 4) = 1
    VMap(4, 5) = 2
    
    VMap(5, 1) = 2
    VMap(5, 2) = 2
    VMap(5, 3) = 2
    VMap(5, 4) = 2
    VMap(5, 5) = 2
    
End Sub


Public Sub LoadP(ByVal Playerk As Integer)
    
    Dim i As Integer
    Dim j As Integer
    
    Call Radar(Playerk)
    For i = 1 To 5
        For j = 1 To 5
            If Player(Playerk).MapV(i, j) = -1 Then VMap(i, j) = 0
            If Player(Playerk).MapV(i, j) = 1 Then VMap(i, j) = 0
            If Player(Playerk).MapV(i, j) = 5 Then VMap(i, j) = 5
            If Player(Playerk).MapV(i, j) = 6 Then VMap(i, j) = 5
        Next j
    Next i

End Sub

Public Sub CalcCount(ByVal Playerk As Integer)
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To 3
        For j = 1 To 5
            Count(1) = Count(1) + VMap(i, j)
        Next j
    Next i
    
    For i = 3 To 5
        For j = 1 To 5
            Count(2) = Count(2) + VMap(i, j)
        Next j
    Next i
    
    For i = 1 To 5
        For j = 1 To 3
            Count(3) = Count(3) + VMap(i, j)
        Next j
    Next i
    
    For i = 1 To 5
        For j = 3 To 5
            Count(4) = Count(4) + VMap(i, j)
        Next j
    Next i
    
End Sub

Public Function Change(ByVal Org As String) As String
        
        Dim a As Integer
        Dim b As Integer
        Dim c As Integer
        
        a = b = c = 0
        a = Int((3000 - 0 + 1) * Rnd)
        b = Int((3000 - 0 + 1) * Rnd)
        c = Int((3000 - 0 + 1) * Rnd)
        
        If Org = "F" Then
            If (a > b) And (b > c) Then Change = "B"
            If (a > c) And (c > b) Then Change = "L"
            If (b > a) And (a > c) Then Change = "R"
            If (b > c) And (c > a) Then Change = "R"
            If (c > a) And (a > b) Then Change = "L"
            If (c > b) And (b > a) Then Change = "B"
        End If
        If Org = "B" Then
            If (a > b) And (b > c) Then Change = "L"
            If (a > c) And (c > b) Then Change = "F"
            If (b > a) And (a > c) Then Change = "R"
            If (b > c) And (c > a) Then Change = "R"
            If (c > a) And (a > b) Then Change = "F"
            If (c > b) And (b > a) Then Change = "L"
        End If
        If Org = "L" Then
            If (a > b) And (b > c) Then Change = "B"
            If (a > c) And (c > b) Then Change = "R"
            If (b > a) And (a > c) Then Change = "F"
            If (b > c) And (c > a) Then Change = "F"
            If (c > a) And (a > b) Then Change = "R"
            If (c > b) And (b > a) Then Change = "B"
        End If
        If Org = "R" Then
            If (a > b) And (b > c) Then Change = "F"
            If (a > c) And (c > b) Then Change = "L"
            If (b > a) And (a > c) Then Change = "B"
            If (b > c) And (c > a) Then Change = "B"
            If (c > a) And (a > b) Then Change = "L"
            If (c > b) And (b > a) Then Change = "F"
        End If
        
End Function

