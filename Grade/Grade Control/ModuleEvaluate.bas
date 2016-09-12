Attribute VB_Name = "ModuleEvaluate"
Option Explicit

Public Function GetGradesId(ByVal TId As Integer, ByVal CId As Integer) As Long
    
    Dim i  As Long
    
    For i = 1 To 10000000
        If Grades(i).TermID = TId Then
            If Grades(i).ClassID = CId Then
                GetGradesId = i
                Exit For
            End If
        End If
    Next i
    
End Function
