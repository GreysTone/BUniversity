VERSION 5.00
Begin VB.Form FrmClock_Plug 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plugins _ Clock"
   ClientHeight    =   2010
   ClientLeft      =   12465
   ClientTop       =   3390
   ClientWidth     =   3030
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTime"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblDate"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "FrmClock_Plug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    
    lblDate.Caption = Date
    lblTime.Caption = Time
    
End Sub
