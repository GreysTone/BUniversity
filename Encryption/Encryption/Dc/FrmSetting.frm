VERSION 5.00
Begin VB.Form FrmSetting 
   Caption         =   "选项"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   Icon            =   "FrmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4575
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   480
      TabIndex        =   1
      Text            =   "App.path\Des.rs!"
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "字典文件目录："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   0
      Picture         =   "FrmSetting.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
