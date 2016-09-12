VERSION 5.00
Begin VB.Form FormAbout 
   BorderStyle     =   0  'None
   Caption         =   "RStudio's Copyright"
   ClientHeight    =   2535
   ClientLeft      =   8850
   ClientTop       =   3810
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin RSGradeControl.ACPRibbon ACPRibbon1 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   1296
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   0
      Picture         =   "FormAbout.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2007 - 2010 RStudio."
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3945
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RS. Grade Control Program"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   4065
   End
   Begin VB.Label lblRS 
      BackStyle       =   0  'Transparent
      Caption         =   "RStudio Present."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Line Line3 
      X1              =   4920
      X2              =   4920
      Y1              =   720
      Y2              =   2520
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4920
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   2520
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    Set ACPRibbon1.Picture = Image1.Picture
    
    Me.Caption = "RStudio's Copyright"
    
    '# Show Caption of Form
    ACPRibbon1.Caption = Me.Caption
    
    ACPRibbon1.Refresh
    
    lblRS.Caption = "RStudio Present"
    lblTitle.Caption = "RS. Grade Control Program"
    lblCopy.Caption = "Copyright (C) 2007 - 2010 RStudio."

End Sub

Private Sub lblCopy_DblClick()

    MsgBox "Thanks ACPRibbon."
    
End Sub
