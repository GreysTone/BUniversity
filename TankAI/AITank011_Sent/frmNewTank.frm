VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewTank 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "¥¥Ω®–¬AIÃπøÀ"
   ClientHeight    =   7845
   ClientLeft      =   6825
   ClientTop       =   3015
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Œ¢»Ì—≈∫⁄"
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
   ScaleHeight     =   7845
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin VB.CommandButton cmdCleanList 
      Caption         =   "«Âø’¡–±Ì"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   29
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton cmdClean 
      Caption         =   "«Âø’»´≤ø"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   28
      Top             =   5400
      Width           =   1935
   End
   Begin VB.ListBox ListIns 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   5
      ItemData        =   "frmNewTank.frx":0000
      Left            =   240
      List            =   "frmNewTank.frx":0002
      TabIndex        =   27
      Top             =   7080
      Width           =   5295
   End
   Begin VB.ListBox ListIns 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   4
      ItemData        =   "frmNewTank.frx":0004
      Left            =   240
      List            =   "frmNewTank.frx":0006
      TabIndex        =   26
      Top             =   6480
      Width           =   5295
   End
   Begin VB.ListBox ListIns 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   3
      ItemData        =   "frmNewTank.frx":0008
      Left            =   240
      List            =   "frmNewTank.frx":000A
      TabIndex        =   25
      Top             =   5880
      Width           =   5295
   End
   Begin VB.ListBox ListIns 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   2
      ItemData        =   "frmNewTank.frx":000C
      Left            =   240
      List            =   "frmNewTank.frx":000E
      TabIndex        =   24
      Top             =   5280
      Width           =   5295
   End
   Begin VB.ListBox ListIns 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Index           =   1
      ItemData        =   "frmNewTank.frx":0010
      Left            =   240
      List            =   "frmNewTank.frx":0012
      TabIndex        =   23
      Top             =   4680
      Width           =   5295
   End
   Begin VB.CommandButton cmdOutput 
      Caption         =   "…˙≥…"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   21
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton CmdIns 
      Caption         =   "ªÒ»°∑∂Œßƒ⁄µ–»ÀŒª÷√"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   20
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton CmdIns 
      Caption         =   "»Œ“‚ø…––Õ®¬∑"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5640
      TabIndex        =   19
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton CmdIns 
      Caption         =   "ªÒ»°µºµØ¿¥‘¥"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5640
      TabIndex        =   18
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton CmdIns 
      Caption         =   "ªÒ»° «∑Ò‘⁄µºµØπ•ª˜œﬂ"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5640
      TabIndex        =   17
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton CmdIns 
      Caption         =   "π•ª˜"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6840
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CmdIns 
      Caption         =   "“∆∂Ø"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CmdIns 
      Caption         =   "≈–∂œ"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   14
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ListBox ListIns 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   5295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   200
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdRand 
      Caption         =   "ÀÊª˙…˙≥…"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   600
      Width           =   1935
   End
   Begin VB.HScrollBar HSB 
      Height          =   255
      Left            =   3720
      Max             =   255
      TabIndex        =   10
      Top             =   1290
      Width           =   2175
   End
   Begin VB.HScrollBar HSG 
      Height          =   255
      Left            =   3720
      Max             =   255
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.HScrollBar HSR 
      Height          =   255
      Left            =   3720
      Max             =   255
      TabIndex        =   8
      Top             =   630
      Width           =   2175
   End
   Begin VB.TextBox txtB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   7
      Text            =   "B - 0"
      Top             =   1290
      Width           =   975
   End
   Begin VB.TextBox txtG 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   6
      Text            =   "G - 0"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtR 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   5
      Text            =   "R - 0"
      Top             =   630
      Width           =   975
   End
   Begin VB.TextBox txtColor 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   3
      Top             =   630
      Width           =   975
   End
   Begin VB.TextBox txtSavePath 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "App.path\Tank\"
      Top             =   270
      Width           =   4455
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   1
      Top             =   270
      Width           =   2055
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "ºÏ≤‚(‘›≤ªø™∑≈)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   8085
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÃπøÀ—’…´"
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   960
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "ÃπøÀ√˚≥∆"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   255
      X2              =   8040
      Y1              =   1815
      Y2              =   1815
   End
End
Attribute VB_Name = "frmNewTank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
    '# Need
End Sub

Private Sub cmdClean_Click()    'Check Over
    
    Dim i As Integer
    
    Call CleanAll
    
    For i = 0 To 5
        ListIns(i).Clear
    Next i
    
    ProgressBar1.Value = 0
    
End Sub

Private Sub cmdCleanList_Click()    'Check Over
    
    Dim i As Integer
    Dim j As Integer
    Dim Clist As String
    Dim NClist As Integer
    
    Clist = InputBox("—°‘Ò¡–±Ì 0 - 5")
    If Clist = "" Then Exit Sub
    
    NClist = Val(Clist)
    If NClist < 0 Or NClist > 5 Then Exit Sub
    
    Select Case NClist
        Case 0
            Call cmdClean_Click
        Case 1
            For i = 0 To 60
                NewTank.List1(i).InsID = 0
                NewTank.List1(i).R = ""
                For j = 0 To 10
                    NewTank.List1(i).V(j) = ""
                Next j
            Next i
            ProgressBar1.Value = ProgressBar1.Value - NewTank.CountFile(NClist)
            NewTank.CountFile(NClist) = 0
        Case 2
            For i = 0 To 60
                NewTank.List2(i).InsID = 0
                NewTank.List2(i).R = ""
                For j = 0 To 10
                    NewTank.List2(i).V(j) = ""
                Next j
            Next i
            ProgressBar1.Value = ProgressBar1.Value - NewTank.CountFile(NClist)
            NewTank.CountFile(NClist) = 0
        Case 3
            For i = 0 To 60
                NewTank.List3(i).InsID = 0
                NewTank.List3(i).R = ""
                For j = 0 To 10
                    NewTank.List3(i).V(j) = ""
                Next j
            Next i
            ProgressBar1.Value = ProgressBar1.Value - NewTank.CountFile(NClist)
            NewTank.CountFile(NClist) = 0
        Case 4
            For i = 0 To 60
                NewTank.List4(i).InsID = 0
                NewTank.List4(i).R = ""
                For j = 0 To 10
                    NewTank.List4(i).V(j) = ""
                Next j
            Next i
            ProgressBar1.Value = ProgressBar1.Value - NewTank.CountFile(NClist)
            NewTank.CountFile(NClist) = 0
        Case 5
            For i = 0 To 60
                NewTank.List5(i).InsID = 0
                NewTank.List5(i).R = ""
                For j = 0 To 10
                    NewTank.List5(i).V(j) = ""
                Next j
            Next i
            ProgressBar1.Value = ProgressBar1.Value - NewTank.CountFile(NClist)
            NewTank.CountFile(NClist) = 0
    End Select
    
    ListIns(Clist).Clear
    
End Sub

Private Sub CmdIns_Click(Index As Integer)    'Check Over
    
    Select Case Index
        Case 0
            DialogAdd.Caption = "≈–∂œ√¸¡Ó"
            DialogAdd.lblName = "IF"
            DialogAdd.lblFormat = "√¸¡Ó∏Ò Ω£∫ IF (±Ì¥Ô Ω1,¬ﬂº≠≈–∂œ∑˚,±Ì¥Ô Ω2,’Ê,ºŸ)"
            DialogAdd.ComboRe.Enabled = False
            DialogAdd.Combo1.Enabled = True
            DialogAdd.Combo2.Enabled = True
            DialogAdd.Combo3.Enabled = True
            DialogAdd.ComboTrue.Enabled = True
            DialogAdd.ComboFalse.Enabled = True
        Case 1
            DialogAdd.Caption = "“∆∂Ø√¸¡Ó"
            DialogAdd.lblName = "Move"
            DialogAdd.lblFormat = "√¸¡Ó∏Ò Ω£∫ Move (±Ì¥Ô Ω1)"
            DialogAdd.ComboRe.Enabled = False
            DialogAdd.Combo1.Enabled = True
            DialogAdd.Combo2.Enabled = False
            DialogAdd.Combo3.Enabled = False
            DialogAdd.ComboTrue.Enabled = False
            DialogAdd.ComboFalse.Enabled = False
        Case 2
            DialogAdd.Caption = "π•ª˜√¸¡Ó"
            DialogAdd.lblName = "Attack"
            DialogAdd.lblFormat = "√¸¡Ó∏Ò Ω£∫ Attack (±Ì¥Ô Ω1)"
            DialogAdd.ComboRe.Enabled = False
            DialogAdd.Combo1.Enabled = True
            DialogAdd.Combo2.Enabled = False
            DialogAdd.Combo3.Enabled = False
            DialogAdd.ComboTrue.Enabled = False
            DialogAdd.ComboFalse.Enabled = False
        Case 3
            DialogAdd.Caption = "ªÒ»° «∑Ò‘⁄µºµØπ•ª˜œﬂ√¸¡Ó"
            DialogAdd.lblName = "GetLockOn"
            DialogAdd.lblFormat = "√¸¡Ó∏Ò Ω£∫ = GetLockOn"
            DialogAdd.ComboRe.Enabled = True
            DialogAdd.Combo1.Enabled = False
            DialogAdd.Combo2.Enabled = False
            DialogAdd.Combo3.Enabled = False
            DialogAdd.ComboTrue.Enabled = False
            DialogAdd.ComboFalse.Enabled = False
        Case 4
            DialogAdd.Caption = "ªÒ»°µºµØ¿¥‘¥√¸¡Ó"
            DialogAdd.lblName = "GetFireFrom"
            DialogAdd.lblFormat = "√¸¡Ó∏Ò Ω£∫ = GetFireFrom"
            DialogAdd.ComboRe.Enabled = True
            DialogAdd.Combo1.Enabled = False
            DialogAdd.Combo2.Enabled = False
            DialogAdd.Combo3.Enabled = False
            DialogAdd.ComboTrue.Enabled = False
            DialogAdd.ComboFalse.Enabled = False
        Case 5
            DialogAdd.Caption = "»Œ“‚ø…––Õ®¬∑√¸¡Ó"
            DialogAdd.lblName = "GetFreeWay"
            DialogAdd.lblFormat = "√¸¡Ó∏Ò Ω£∫ = GetFreeWay"
            DialogAdd.ComboRe.Enabled = True
            DialogAdd.Combo1.Enabled = False
            DialogAdd.Combo2.Enabled = False
            DialogAdd.Combo3.Enabled = False
            DialogAdd.ComboTrue.Enabled = False
            DialogAdd.ComboFalse.Enabled = False
        Case 6
            DialogAdd.Caption = "ªÒ»°∑∂Œßƒ⁄µ–»ÀŒª÷√√¸¡Ó"
            DialogAdd.lblName = "GetFind"
            DialogAdd.lblFormat = "√¸¡Ó∏Ò Ω£∫ = GetFind"
            DialogAdd.ComboRe.Enabled = True
            DialogAdd.Combo1.Enabled = False
            DialogAdd.Combo2.Enabled = False
            DialogAdd.Combo3.Enabled = False
            DialogAdd.ComboTrue.Enabled = False
            DialogAdd.ComboFalse.Enabled = False
    End Select
    DialogAdd.Visible = True
    InsCode = Index
    
End Sub

Private Sub cmdOutput_Click()   'Check Over

    Dim i As Integer
    Dim TestFile As String
    
    If txtName.Text = "" Then
        MsgBox "√˚◊÷≤ªƒ‹Œ™ø’"
        Exit Sub
    End If
    
    TestFile = Dir(App.Path & "\Tank\" & txtName.Text & "\")
    If TestFile <> "" Then
        MsgBox "Tank [" & txtName.Text & "] Exist." & vbCrLf & " Please check.", vbOKOnly
        Exit Sub
    End If
    
    TestFile = Dir(App.Path & "\Tank\" & txtName.Text & "\", vbDirectory)
    If TestFile = "." Or TestFile = ".." Then
        MsgBox "Tank [" & txtName.Text & "] Exist." & vbCrLf & " Please check.", vbOKOnly
        Exit Sub
    End If
    
    MkDir App.Path & "\Tank\" & txtName.Text
    
    '# RS Tank Basic Data
    Open App.Path & "\Tank\" & txtName.Text & "\tank.rtb" For Append As #1
    Write #1, "[RS TANK BASIC DATA]"
    Write #1, txtName.Text
    Write #1, HSR.Value, HSG.Value, HSB.Value
    Close #1
    
    '# RS Tank Main Data
    Open App.Path & "\Tank\" & txtName.Text & "\tank.rtm" For Append As #1
    Write #1, "[RS TANK DATA]"
    Write #1, NewTank.CountFile(0)
        For i = 1 To NewTank.CountFile(0)
            Write #1, NewTank.ListMain(i).InsID, NewTank.ListMain(i).R, NewTank.ListMain(i).V(1), NewTank.ListMain(i).V(2), NewTank.ListMain(i).V(3), NewTank.ListMain(i).V(4), NewTank.ListMain(i).V(5)
        Next i
    Close #1
    
    '# RS File1
    Open App.Path & "\Tank\" & txtName.Text & "\file1.rf" For Append As #1
    Write #1, "[RS TANK DATA]"
    Write #1, NewTank.CountFile(1)
        For i = 1 To NewTank.CountFile(1)
            Write #1, NewTank.List1(i).InsID, NewTank.List1(i).R, NewTank.List1(i).V(1), NewTank.List1(i).V(2), NewTank.List1(i).V(3), NewTank.List1(i).V(4), NewTank.List1(i).V(5)
        Next i
    Close #1
    
    '# RS File2
    Open App.Path & "\Tank\" & txtName.Text & "\file2.rf" For Append As #1
    Write #1, "[RS TANK DATA]"
    Write #1, NewTank.CountFile(2)
        For i = 1 To NewTank.CountFile(2)
            Write #1, NewTank.List2(i).InsID, NewTank.List2(i).R, NewTank.List2(i).V(1), NewTank.List2(i).V(2), NewTank.List2(i).V(3), NewTank.List2(i).V(4), NewTank.List2(i).V(5)
        Next i
    Close #1
    
    '# RS File3
    Open App.Path & "\Tank\" & txtName.Text & "\file3.rf" For Append As #1
    Write #1, "[RS TANK DATA]"
    Write #1, NewTank.CountFile(3)
        For i = 1 To NewTank.CountFile(3)
            Write #1, NewTank.List3(i).InsID, NewTank.List3(i).R, NewTank.List3(i).V(1), NewTank.List3(i).V(2), NewTank.List3(i).V(3), NewTank.List3(i).V(4), NewTank.List3(i).V(5)
        Next i
    Close #1
    
    '# RS File4
    Open App.Path & "\Tank\" & txtName.Text & "\file4.rf" For Append As #1
    Write #1, "[RS TANK DATA]"
    Write #1, NewTank.CountFile(4)
        For i = 1 To NewTank.CountFile(4)
            Write #1, NewTank.List4(i).InsID, NewTank.List4(i).R, NewTank.List4(i).V(1), NewTank.List4(i).V(2), NewTank.List4(i).V(3), NewTank.List4(i).V(4), NewTank.List4(i).V(5)
        Next i
    Close #1
    
    '# RS File5
    Open App.Path & "\Tank\" & txtName.Text & "\file5.rf" For Append As #1
    Write #1, "[RS TANK DATA]"
    Write #1, NewTank.CountFile(5)
        For i = 1 To NewTank.CountFile(5)
            Write #1, NewTank.List5(i).InsID, NewTank.List5(i).R, NewTank.List5(i).V(1), NewTank.List5(i).V(2), NewTank.List5(i).V(3), NewTank.List5(i).V(4), NewTank.List5(i).V(5)
        Next i
    Close #1
    
    MsgBox "Output Over."
    
    frmMain.Visible = True
    Unload Me
    
End Sub

Private Sub cmdRand_Click()    'Check Over
    
    HSR.Value = Int((255 - 0 + 1) * Rnd)
    HSG.Value = Int((255 - 0 + 1) * Rnd)
    HSB.Value = Int((255 - 0 + 1) * Rnd)
    
End Sub

Private Sub Form_Load()     'Check Over
    
    Randomize
    Call cmdClean_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)  'Check Over
    
    frmMain.Visible = True
    
End Sub

Private Sub HSB_Change()    'Check Over
    
    txtB.Text = "B - " & HSB.Value
    txtB.BackColor = RGB(0, 0, HSB.Value)
    txtColor.BackColor = RGB(HSR.Value, HSG.Value, HSB.Value)
    NewTank.Color.R = HSR.Value
    NewTank.Color.G = HSG.Value
    NewTank.Color.B = HSB.Value
    
End Sub

Private Sub HSG_Change()    'Check Over
    
    txtG.Text = "G - " & HSG.Value
    txtG.BackColor = RGB(0, HSG.Value, 0)
    txtColor.BackColor = RGB(HSR.Value, HSG.Value, HSB.Value)
    NewTank.Color.R = HSR.Value
    NewTank.Color.G = HSG.Value
    NewTank.Color.B = HSB.Value
    
End Sub

Private Sub HSR_Change()    'Check Over
    
    txtR.Text = "R - " & HSR.Value
    txtR.BackColor = RGB(HSR.Value, 0, 0)
    txtColor.BackColor = RGB(HSR.Value, HSG.Value, HSB.Value)
    NewTank.Color.R = HSR.Value
    NewTank.Color.G = HSG.Value
    NewTank.Color.B = HSB.Value
    
End Sub

Private Sub txtName_Change()    'Check Over
    
    txtSavePath = "App.path\Tank\" & txtName.Text & "\"
    NewTank.Names = txtName.Text
    
End Sub
