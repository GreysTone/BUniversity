Attribute VB_Name = "ModuleGetFreeWay"
Option Explicit

'Public Type Way
'    Ws As String
'    Value As Integer
'    Data As Integer
'    Use As Boolean
'End Type

'Public W(5) As Way
'Public LastFreeway(10) As String

Public Function GetFreeWay2(ByVal Playerk As Integer) As String
    
    'Randomize
    
    Dim NX As Integer
    Dim NY As Integer
    
    NX = PositionMap.Obc(Playerk).X
    NY = PositionMap.Obc(Playerk).Y
    
    'GetFreeWay = "S"
    
    'GetFreeWay = RandWay(NX, NY, (Int((1000 - 0 + 1) * Rnd + 0) Mod 4 + 1))
    
    'LastFreeway(Playerk) = GetFreeWay
    
End Function

Public Function RandWay(ByVal NX As Integer, ByVal NY As Integer, ByVal Id As Integer) As String
    
    Select Case Id
        Case 1
            If (NX - 1 > 0) And PositionMap.Data(NX - 1, NY) = 1 Then
                RandWay = "F"
                Exit Function
            End If
            If NX + 1 < OrMap.X And PositionMap.Data(NX + 1, NY) = 1 Then
                RandWay = "B"
                Exit Function
            End If
            If NY - 1 > 0 And PositionMap.Data(NX, NY - 1) = 1 Then
                RandWay = "L"
                Exit Function
            End If
            If NY + 1 < OrMap.Y And PositionMap.Data(NX, NY + 1) = 1 Then
                RandWay = "R"
                Exit Function
            End If
        Case 2
            If NX + 1 < OrMap.X And PositionMap.Data(NX + 1, NY) = 1 Then
                RandWay = "B"
                Exit Function
            End If
            If NY - 1 > 0 And PositionMap.Data(NX, NY - 1) = 1 Then
                RandWay = "L"
                Exit Function
            End If
            If NY + 1 < OrMap.Y And PositionMap.Data(NX, NY + 1) = 1 Then
                RandWay = "R"
                Exit Function
            End If
            If NX - 1 > 0 And PositionMap.Data(NX - 1, NY) = 1 Then
                RandWay = "F"
                Exit Function
            End If
        Case 3
            If NY - 1 > 0 And PositionMap.Data(NX, NY - 1) = 1 Then
                RandWay = "L"
                Exit Function
            End If
            If NY + 1 < OrMap.Y And PositionMap.Data(NX, NY + 1) = 1 Then
                RandWay = "R"
                Exit Function
            End If
            If NX - 1 > 0 And PositionMap.Data(NX - 1, NY) = 1 Then
                RandWay = "F"
                Exit Function
            End If
            If NX + 1 < OrMap.X And PositionMap.Data(NX + 1, NY) = 1 Then
                RandWay = "B"
                Exit Function
            End If
        Case 4
            If NY + 1 < OrMap.Y And PositionMap.Data(NX, NY + 1) = 1 Then
                RandWay = "R"
                Exit Function
            End If
            If NX - 1 > 0 And PositionMap.Data(NX - 1, NY) = 1 Then
                RandWay = "F"
                Exit Function
            End If
            If NX + 1 < OrMap.X And PositionMap.Data(NX + 1, NY) = 1 Then
                RandWay = "B"
                Exit Function
            End If
            If NY - 1 > 0 And PositionMap.Data(NX, NY - 1) = 1 Then
                RandWay = "L"
                Exit Function
            End If
    End Select
    
End Function

