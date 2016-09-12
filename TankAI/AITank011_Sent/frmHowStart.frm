VERSION 5.00
Begin VB.Form frmHowStart 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "如何开始"
   ClientHeight    =   3270
   ClientLeft      =   9840
   ClientTop       =   3585
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label3 
      Caption         =   "       是的，您需要创建自己的AI坦克，您可以在【资源】菜单中找到有关内容。更多帮助内容详见附带文档。"
      Height          =   915
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   3915
   End
   Begin VB.Label Label2 
      Caption         =   $"frmHowStart.frx":0000
      Height          =   1275
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "如何开始游戏？ 很简单！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2715
   End
End
Attribute VB_Name = "frmHowStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
