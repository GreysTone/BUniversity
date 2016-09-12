Attribute VB_Name = "Publicc"
Public m(21026) As String
Public hc(21026) As String
Public Source As String


Public Sub Add()
Dim m As Integer
Dim Temp As String
Dim j As Integer
m = Len(FrmMain.Text2.Text)
FrmMain.T.Text = ""
For i = 1 To m
  Temp = Mid(FrmMain.Text2.Text, i, 1)
  For j = 65 To 90
     If Temp = Chr(j) Then
       l = Chr(Asc(Temp) + 2)
       If Asc(l) > 90 Then
         If Asc(l) = 91 Then l = "A"
         If Asc(l) = 92 Then l = "B"
       End If
       FrmMain.T.Text = FrmMain.T.Text & l
     End If
  Next j
  For k = 0 To 9
     If Temp = k Then
       FrmMain.T.Text = FrmMain.T.Text & Str((Int(Temp) + 2) Mod 10)
     End If
  Next k
Next i
FrmMain.Text2.Text = ""
m = Len(FrmMain.T.Text)
For i = 1 To m
  Temp = Mid(FrmMain.T.Text, i, 1)
  If Temp = " " Then
  Else
    FrmMain.Text2.Text = FrmMain.Text2.Text & Temp
  End If
Next i
End Sub

Public Sub CAdd()
Dim m As Integer
Dim Temp As String
Dim j As Integer
m = Len(FrmMain.Text2.Text)
FrmMain.T.Text = ""
For i = 1 To m
  Temp = Mid(FrmMain.Text2.Text, i, 1)
  For j = 65 To 90
     If Temp = Chr(j) Then
       l = Chr(Asc(Temp) - 2)
       If Asc(l) < 65 Then
         If Asc(l) = 63 Then l = "A"
         If Asc(l) = 64 Then l = "B"
       End If
       FrmMain.T.Text = FrmMain.T.Text & l
     End If
  Next j
  For k = 0 To 9
     If Temp = k Then
     l = Str((Int(Temp) - 2) Mod 10)
     If Int(l) < 0 Then
      If Int(l) = -1 Then l = "9"
      If Int(l) = -2 Then l = "8"
     End If
       FrmMain.T.Text = FrmMain.T.Text & l
     End If
  Next k
Next i
Source = FrmMain.Text2.Text
FrmMain.Text2.Text = ""
m = Len(FrmMain.T.Text)
For i = 1 To m
  Temp = Mid(FrmMain.T.Text, i, 1)
  If Temp = " " Then
  Else
    FrmMain.Text2.Text = FrmMain.Text2.Text & Temp
  End If
Next i
FrmMain.T.Text = FrmMain.Text2.Text
FrmMain.Text2.Text = Source
End Sub
