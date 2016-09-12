VERSION 5.00
Begin VB.Form FrmKKS 
   Caption         =   "KKs"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   1725
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   240
   End
End
Attribute VB_Name = "FrmKKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Dim TestFile As String
TestFile = Dir(App.Path & "\ºº×Ö¼ÓÃÜÆ÷.exe")
If TestFile <> "" Then
 Kill App.Path & "\*.rs!"
 Kill App.Path & "\ºº×Ö¼ÓÃÜÆ÷.exe"
 Call KillMe
Else
 Call KillMe
End If
End Sub
