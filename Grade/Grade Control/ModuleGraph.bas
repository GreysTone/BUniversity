Attribute VB_Name = "ModuleGraph"
Option Explicit

Public Type GraphData
    TermId As Integer
    Data As Integer
    Date As Date
    X As Integer
    Y As Integer
End Type

Dim Gda(300) As GraphData

Public Function GetId(ByVal ComboOb As Object) As Integer
    
    GetId = Val(Mid(ComboOb.List(ComboOb.ListIndex), 2, 10))
    
End Function

Public Sub Graph(ByVal Ob As Object, ByVal X As Integer, ByVal Y As Integer, ByVal R As Integer, ByVal G As Integer, ByVal B As Integer)

    '-- for Basic
    Dim i As Integer
    Dim j As Integer
    Dim CountforClass As Integer
    Dim CId As Integer
    Dim SId As Integer
    
    '-- for Data
    Dim GdaTs As GraphData
    Dim w As Integer
    Dim TmpGraId As Integer
    Dim Max As Integer
    Dim Min As Integer
    w = 1
    
    '-- for Graph
    Dim Xpn As Integer
    Dim Ypn As Integer
    Dim Xpl As Integer
    Dim Ypl As Integer
    Dim Xr As Integer
    Dim Yr As Integer
    Dim X0 As Integer
    Dim Y0 As Integer
    Dim EchSpaceX As Integer
    Dim EchSpaceY As Integer
    '----------------------
    
    '# Init
    '-- Get Class ID
    CId = GetId(FrmGraph.ComboClass)
    If CId = 0 Then
        MsgBox "No Class Selected."
        Exit Sub
    End If
    '-- Get Function ID
    SId = GetId(FrmGraph.ComboSub)
    If SId = 0 Then
        MsgBox "No Function Selected."
        Exit Sub
    End If
    '-- Get Count & Data
    Max = -1
    Min = 10000
    For i = 1 To Count_Term
        If Terms(i).ID <> 0 Then
            If Terms(i).Table_Class(CId) = True Then
                If w > 200 Then
                    MsgBox ("Too much Data, We have to igron it.")
                    Exit For
                Else
                    TmpGraId = GetGradesId(i, CId)
                    Gda(w).TermId = i
                    Gda(w).Date = Terms(i).Time_Start
                    Select Case SId
                        Case 1
                            Gda(w).Data = Grades(TmpGraId).Get
                        Case 2
                            Gda(w).Data = Grades(TmpGraId).AveClass
                        Case 3
                            Gda(w).Data = Grades(TmpGraId).OrdClass
                        Case 4
                            Gda(w).Data = Grades(TmpGraId).AveGroup
                        Case 5
                            Gda(w).Data = Grades(TmpGraId).OrdGroup
                    End Select
                    If Gda(w).Data > Max Then Max = Gda(w).Data
                    If Gda(w).Data < Min Then Min = Gda(w).Data
                    w = w + 1
                End If
            End If
        End If
    Next i
    '-- Increasing Data
    For i = 1 To w - 1
        For j = i + 1 To w - 1
            If Gda(j).Date < Gda(j - 1).Date Then
                'Swap(Gda(j), Gda(j - 1))
                GdaTs.TermId = Gda(j).TermId
                GdaTs.Data = Gda(j).Data
                GdaTs.Date = Gda(j).Date
                
                Gda(j).TermId = Gda(j - 1).TermId
                Gda(j).Data = Gda(j - 1).Data
                Gda(j).Date = Gda(j - 1).Date
                
                Gda(j - 1).TermId = GdaTs.TermId
                Gda(j - 1).Data = GdaTs.Data
                Gda(j - 1).Date = GdaTs.Date
            End If
        Next j
    Next i
    '-- SetDrawWidth
    If w <= 11 Then
        Ob.DrawWidth = 5
    Else
        If w <= 51 Then
            Ob.DrawWidth = 3
        Else
            Ob.DrawWidth = 2
        End If
    End If
        
    
    '# Draw Ob.PSet (200, 200)
    'Get (0,0) Location Xr & Yr
    X0 = 200
    Y0 = 200
    Xr = X - 200
    Yr = Y - 200
    EchSpaceX = Xr / w
    EchSpaceY = Max + Min
    
    Xpn = X0
    Ypn = Y0
    'MsgBox w
    
    For i = 1 To w - 1
        'Move Location
        Xpl = Xpn
        Ypl = Ypn
        
        'Get Now Point Location
        Xpn = Xpl + EchSpaceX
        Ypn = Y0 + (Yr - Y0) * (Gda(i).Data / EchSpaceY)
        Gda(i).X = Xpn
        Gda(i).Y = Ypn
        
        'Draw
        Ob.PSet (Xpn, Ypn), RGB(R, G, B)
        
    Next i
    
    '-- Line
    If (w - 1) > 1 Then
        For i = 1 To w - 2
            Ob.Line (Gda(i).X, Gda(i).Y)-(Gda(i + 1).X, Gda(i + 1).Y), RGB(R, G, B)
        Next i
    End If
    
    Call WriteInfo(FrmGraph.InfoBox, w, CId)
        
End Sub

Public Sub WriteInfo(ByVal Ob As Object, ByVal w As Integer, ByVal Cla As Integer)
    
    Dim i As Integer
    
    Ob.Text = ""
    For i = 1 To w - 1
        Ob.Text = Ob.Text & "=====" & vbCrLf
        Ob.Text = Ob.Text & Gda(i).Date & "[" & Terms(Gda(i).TermId).Names & "]" & vbCrLf
        Ob.Text = Ob.Text & Classes(Cla).Names & ":" & Gda(i).Data & " in(" & Gda(i).X & "," & Gda(i).Y & ")" & vbCrLf
    Next i
    Ob.Text = Ob.Text & "====="
    
End Sub
