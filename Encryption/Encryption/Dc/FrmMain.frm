VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "RCH"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助"
      Height          =   855
      Left            =   2880
      Picture         =   "FrmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CmdUpData 
      Caption         =   "升级"
      Height          =   855
      Left            =   1560
      Picture         =   "FrmMain.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CmdSet 
      Caption         =   "选项"
      Height          =   855
      Left            =   240
      Picture         =   "FrmMain.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton CmdUDC 
      Caption         =   "解密"
      Height          =   735
      Left            =   3240
      Picture         =   "FrmMain.frx":17A8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton CmdADC 
      Caption         =   "加密"
      Height          =   735
      Left            =   3240
      Picture         =   "FrmMain.frx":1D32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   0
      Picture         =   "FrmMain.frx":22BC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdADC_Click()
Dim i, j As Integer
Dim ml As Integer
Dim Temp As String
ml = Len(Text1.Text)
Text2.Text = ""
For i = 1 To ml
  Temp = Mid(Text1.Text, i, 1)
  For j = 0 To 21021
    If Temp = m(j) Then
      Text2.Text = Text2.Text + hc(j)
    Else
      Text2.Text = Text2.Text + Temp
    End If
  Next j
Next i
End Sub

Private Sub CmdSet_Click()
Dim k As Integer
Dim pass As String
k = MsgBox("选项命令为高手设计，您是否需要进入？", vbYesNo)
If k = 6 Then
  pass = InputBox("进入设置需要通行证，请输入您的通行证的ID序列：", "通行证ID序列需求")
  If pass = "RossStudio-Song" Then
     FrmSetting.Show
  End If
Else
  Exit Sub
End If
End Sub

Private Sub CmdUDC_Click()
Dim i, j As Integer
Dim ml As Integer
Dim Temp As String
ml = Len(Text2.Text)
Text1.Text = ""
For i = 1 To ml Step 4
  Temp = Mid(Text2.Text, i, 4)
  For j = 0 To 21021
    If Temp = hc(j) Then
      Text1.Text = Text1.Text + m(j)
    Else
      Text2.Text = Text2.Text + Temp
    End If
  Next j
Next i
End Sub

Private Sub Form_Load()
Call MDLDS.LoadFile
End Sub
