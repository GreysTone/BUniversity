Attribute VB_Name = "ModuleGetFireDirection"
Option Explicit

Public Function GetFireDirection(ByVal Playerk As Integer) As String
    
    Dim NX As Integer
    Dim NY As Integer
    Dim MisId As Integer
    
    NX = PositionMap.Obc(Playerk).X
    NY = PositionMap.Obc(Playerk).Y
    
    If MissileLockMap.Data(NX, NY).V = True Then
        'GetFireDirection = "1"
        MisId = MissileLockMap.Data(NX, NY).Id
        Select Case MissileMap.Obc(MisId).FireDir
            Case 0
                GetFireDirection = "1"
            Case 1
                GetFireDirection = "0"
            Case 2
                GetFireDirection = "3"
            Case 3
                GetFireDirection = "2"
        End Select
    Else
        GetFireDirection = "-1"
    End If
    
End Function

