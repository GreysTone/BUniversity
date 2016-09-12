VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Ac"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   3600
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "ADD!"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e1(25), e2(25), n1(10), n2(10) As String
Dim m(20902 + 121) As Variant
Dim hc(20902 + 121) As String


Private Sub Command1_Click()
Call zero
Open App.Path + "\CO.txt" For Input As #1
Open App.Path + "\COC.txt" For Output As #2

For i = 0 To 20901 + 121
   Input #1, m(i)
   Write #2, m(i), hc(i)
Next i
Close #1, #2
MsgBox "Add Over!"
End Sub


Private Sub zero()
Dim i, j As Integer
For i = 0 To 25
   e1(i) = Chr(i + 65)
   e2(i) = Chr(i + 65)
Next i
For i = 0 To 9
   n1(i) = i
   n2(i) = i
Next i
For i = 0 To 20901 + 121
    If j Mod 10 = 0 Then
     j = 0
     hc(i) = n2(j)
     j = j + 1
    Else
     hc(i) = n2(j)
     j = j + 1
    End If
Next i
j = 0
For i = 0 To 20901 + 121
  If i Mod 10 <> 0 Then
    hc(i) = e2(j - 1) + hc(i)
  Else
    If j Mod 26 = 0 Then
     j = 0
     hc(i) = e2(j) & hc(i)
     j = j + 1
    Else
     hc(i) = e2(j) & hc(i)
     j = j + 1
    End If
  End If
Next i
j = 0
For i = 0 To 20901 + 121
 If i Mod 260 <> 0 Then
    hc(i) = n1(j - 1) & hc(i)
  Else
    If j Mod 10 = 0 Then
     j = 0
     hc(i) = n1(j) & hc(i)
     j = j + 1
    Else
     hc(i) = n1(j) & hc(i)
     j = j + 1
    End If
End If
Next i
j = 0
For i = 0 To 20901 + 121
 If i Mod 2600 <> 0 Then
    hc(i) = e1(j - 1) + hc(i)
  Else
    If j Mod 26 = 0 Then
     j = 0
     hc(i) = e1(j) & hc(i)
     j = j + 1
    Else
     hc(i) = e1(j) & hc(i)
     j = j + 1
    End If
End If
Next i
End Sub

