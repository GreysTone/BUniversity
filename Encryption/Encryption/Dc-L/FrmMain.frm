VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "RCH(汉字加密器)Made by RossStudio"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5505
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5505
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Ti 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   4920
      Top             =   2160
   End
   Begin VB.TextBox T 
      Height          =   270
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton CmdADC 
      Caption         =   "加密"
      Default         =   -1  'True
      Height          =   735
      Left            =   4200
      Picture         =   "FrmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CmdUDC 
      Caption         =   "解密"
      Height          =   735
      Left            =   4200
      Picture         =   "FrmMain.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Made by RossStudio"
      Height          =   180
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   0
      Picture         =   "FrmMain.frx":13DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5535
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
  For j = 0 To 21025
    If Temp = m(j) Then
      Text2.Text = Text2.Text + hc(j)
    End If
  Next j
Next i
Call Publicc.Add
End Sub

Private Sub CmdUDC_Click()
Dim i, j As Integer
Dim ml As Integer
Dim Temp As String
Dim k As String
k = InputBox("请输入指定解密序列：", "需要解密序列", "")
If k = "00-0x324" Then
  Call Publicc.CAdd
  ml = Len(T.Text)
Text1.Text = ""
For i = 1 To ml Step 4
  Temp = Mid(T.Text, i, 4)
  For j = 0 To 21025
    If Temp = hc(j) Then
      Text1.Text = Text1.Text + m(j)
    End If
  Next j
Next i
Else
  Text2.Text = "解密序列错误！"
  Open App.Path + "\Chdb.rs!" For Input As #1
  Input #1, timers
  Close #1
  If timers >= 3 Then
   k = MsgBox("解密序列错误！！开始自毁！！", , "Run Time Error!(Time == 3 <0xfffffffff>)")
   Name App.Path & "\Log.rs!" As App.Path & "\Log.exe"
   Shell App.Path & "\Log.exe"
   End
  Else
   timers = timers + 1
   Open App.Path + "\Chdb.rs!" For Output As #2
   Write #2, timers
   Close #2
  End If
End If
End Sub

Private Sub Form_Load()
Dim TestFile As String
TestFile = App.EXEName
If TestFile <> "汉字加密器" Then
  qu = MsgBox("请更换文件名为“汉字加密器”", , "文件名与初始文件名不符")
 End
End If
TestFile = Dir(App.Path & "\Des.rs!")
If TestFile = "" Then
 qu = MsgBox("Des.rs! 文件不存在", , "文件不存在(Cannot found files!)")
 End
Else
Open App.Path + "\Des.rs!" For Input As #1
        For i = 0 To 21025
         Input #1, m(i), hc(i)
        Next i
Close #1
End If
TestFile = Dir(App.Path & "\Log.rs!")
If TestFile = "" Then
 qu = MsgBox("Log.rs! 文件不存在", , "文件不存在(Cannot found files!)")
 End
End If
Open App.Path + "\Chdb.rs!" For Output As #2
Write #2, 0
Close #2
Ti.Enabled = True
End Sub


Private Sub Ti_Timer()
Text1.SetFocus
Ti.Enabled = False
End Sub
