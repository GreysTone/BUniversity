VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AI TANK - ÑÝÊ¾ / Copyright (C) 2007 - 2010 RStudio"
   ClientHeight    =   6690
   ClientLeft      =   -15
   ClientTop       =   570
   ClientWidth     =   9120
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
   ScaleHeight     =   6690
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   360
      Max             =   600
      TabIndex        =   19
      Top             =   6240
      Width           =   6255
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   5400
   End
   Begin VB.CommandButton cmdBreak 
      Caption         =   "Break"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   18
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "<..Start..>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   5520
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   1695
      Left            =   8760
      TabIndex        =   7
      Top             =   240
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2990
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   5400
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5400
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   4560
      Width           =   2055
   End
   Begin VB.PictureBox P2Sight 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   5400
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.PictureBox P1Sight 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   5400
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.ListBox ListProgress 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      ItemData        =   "frmMain.frx":0000
      Left            =   360
      List            =   "frmMain.frx":0007
      TabIndex        =   1
      Top             =   5280
      Width           =   4935
   End
   Begin VB.PictureBox PhMap 
      BackColor       =   &H00C0C0C0&
      Height          =   4935
      Left            =   360
      ScaleHeight     =   4875
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
   Begin MSComctlLib.ProgressBar PB2 
      Height          =   1695
      Left            =   8760
      TabIndex        =   8
      Top             =   2400
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2990
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Speed  10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6720
      TabIndex        =   20
      Top             =   6240
      Width           =   840
   End
   Begin VB.Label P2Ata 
      AutoSize        =   -1  'True
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   16
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label P2FireDir 
      AutoSize        =   -1  'True
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   15
      Top             =   3240
      Width           =   600
   End
   Begin VB.Label P2Dir 
      AutoSize        =   -1  'True
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   14
      Top             =   3000
      Width           =   600
   End
   Begin VB.Label P1Ata 
      AutoSize        =   -1  'True
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   13
      Top             =   1320
      Width           =   600
   End
   Begin VB.Label P1FireDir 
      AutoSize        =   -1  'True
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   12
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label P1Dir 
      AutoSize        =   -1  'True
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   11
      Top             =   840
      Width           =   600
   End
   Begin VB.Label lblP2Name 
      AutoSize        =   -1  'True
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   10
      Top             =   2760
      Width           =   600
   End
   Begin VB.Label lblP2 
      AutoSize        =   -1  'True
      Caption         =   "P2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7560
      TabIndex        =   6
      Top             =   2400
      Width           =   360
   End
   Begin VB.Label lblP1 
      AutoSize        =   -1  'True
      Caption         =   "P1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7560
      TabIndex        =   5
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblP1Name 
      AutoSize        =   -1  'True
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7560
      TabIndex        =   9
      Top             =   600
      Width           =   600
   End
   Begin VB.Menu MenuAction 
      Caption         =   "¶ÔÕ½"
      Begin VB.Menu M_NewAction 
         Caption         =   "ÐÂ¶ÔÕ½"
         Shortcut        =   ^N
      End
      Begin VB.Menu Shelf1 
         Caption         =   "-"
      End
      Begin VB.Menu M_EndProgram 
         Caption         =   "½áÊø"
      End
   End
   Begin VB.Menu MenuSource 
      Caption         =   "×ÊÔ´"
      Begin VB.Menu M_NewTank 
         Caption         =   "´´½¨ÐÂAIÌ¹¿Ë"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "°ïÖú"
      Begin VB.Menu M_Start 
         Caption         =   "ÈçºÎ¿ªÊ¼"
      End
      Begin VB.Menu Shelf3 
         Caption         =   "-"
      End
      Begin VB.Menu M_Introduction 
         Caption         =   "°ïÖú/½éÉÜ"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim h As Integer
Dim m As Integer
Dim s As Integer
Dim StartAt As Integer

Private Sub cmdBreak_Click()
    
    If State_GameStart = True Then
        Timer1.Enabled = False
        State_BreakGame = True
        WinGame = 99
    End If
    
    Picture1.Cls
    Picture1.Print "Broken Game.."
        
    cmdStart.Enabled = False
    'cmdLoad.Enabled = False
    cmdBreak.Enabled = False
    State_P1Load = False
    State_P2Load = False
    State_MapLoad = False
    
End Sub

Private Sub MainProgress()
    
    ' Get WhoisFirst
    StartAt = WhoisFirst
    
    If StartAt = 1 Then
        lblP1.BackColor = vbWhite
    Else
        lblP2.BackColor = vbWhite
    End If
    
    WinGame = -1
    Count_Combo = 0
    
    'Draw Radar
    Call DrawRadarMap(P1Sight, 1, 1)
    Call DrawRadarMap(P2Sight, 1, 2)
    
    frmMain.P1Dir.Caption = "None"
    frmMain.P2Dir.Caption = "None"
    frmMain.P1FireDir.Caption = "None"
    frmMain.P2FireDir.Caption = "None"
    frmMain.P1Ata.Caption = "None"
    frmMain.P2Ata.Caption = "None"
    
    Timer2.Enabled = True
    
End Sub

Private Sub cmdStart_Click()
    
    State_GameStart = True
    State_BreakGame = False
    
    'Set Time
    Timer1.Enabled = True
    
    Call MainProgress
    Call WriteInfo(0, 99, 0)
    
End Sub

Private Sub Form_Load()    'Check Over
    
    Call ModuleInstruction.InitInstructions
    
    Randomize
    HScroll1.Value = 10
    
End Sub

Private Sub Form_Unload(Cancel As Integer)    'Check Over
    
    End
    
End Sub

Private Sub HScroll1_Change()
    
    lblValue.Caption = "Speed  " & HScroll1.Value
    Timer2.Interval = HScroll1.Value
    
End Sub

Private Sub lblValue_DblClick()
    
    HScroll1.max = Val(InputBox("Max of HScroll", , "600"))
    
End Sub

Private Sub M_EndProgram_Click()    'Check Over
    
    End
    
End Sub

Private Sub M_Introduction_Click()    'Check Over
    
    frmAbout.Visible = True
    
End Sub

Private Sub M_NewAction_Click()    'Check Over
    
    frmNewAction.Visible = True
    frmMain.Visible = False
    
End Sub

Private Sub M_NewTank_Click()    'Check Over
    
    frmNewTank.Visible = True
    frmMain.Visible = False
    
End Sub

Private Sub M_Start_Click()    'Check Over
    
    frmHowStart.Visible = True
        
End Sub

Private Sub PhMap_DblClick()    'Check Over
    
    Call M_NewAction_Click
    frmMain.Visible = False

End Sub

Private Sub Timer1_Timer()  'Check Over
    
    Picture1.Cls
    Picture1.Print "Game Start" & vbCrLf & Format(Str(h), "00") & ":" & Format(Str(m), "00") & ":" & Format(Str(s), "00")
    s = s + 1
    If s >= 59 Then
        s = 0
        m = m + 1
        If m >= 59 Then
            m = 0
            h = h + 1
        End If
    End If
    
End Sub

Private Sub Timer2_Timer()
    
    If State_BreakGame <> False Or WinGame <> -1 Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        
        'Support Game Info
        Select Case WinGame
            Case 99
                Picture1.Cls
                Picture1.Print "End Game" & vbCrLf & Format(Str(h), "00") & ":" & Format(Str(m), "00") & ":" & Format(Str(s), "00")
                Picture1.Print GameOverInfo
            Case 1
                Picture1.Cls
                Picture1.Print "End Game" & vbCrLf & Format(Str(h), "00") & ":" & Format(Str(m), "00") & ":" & Format(Str(s), "00")
                Picture1.Print GameOverInfo & vbCrLf & "Player1 Win!"
            Case 2
                Picture1.Cls
                Picture1.Print "End Game" & vbCrLf & Format(Str(h), "00") & ":" & Format(Str(m), "00") & ":" & Format(Str(s), "00")
                Picture1.Print GameOverInfo & vbCrLf & "Player2 Win!"
        End Select
        Call AddLayerDrawMap(frmMain.PhMap, OrMap.X, OrMap.Y, 1)
        Call WriteInfo(0, 100, 0)
        
        cmdStart.Enabled = False
        cmdBreak.Enabled = False
        lblP1.BackColor = &H8000000F
        lblP2.BackColor = &H8000000F
        
        Exit Sub
    End If
        
    Count_Combo = Count_Combo + 1
        
    NowTurn = WhoisNext(Count_Combo, StartAt)
    Call PlayerDone(NowTurn)
    Call ModuleMissileDone.MissileDone
    If Count_Combo Mod 2 = 0 Then
        Comboes = Comboes + 1
            
        'Refresh Move = False / Attack = False
        P1Move = False
        P2Move = False
        P1Attack = False
        P2Attack = False
    End If
    
    'Refresh P1 & P2 's Sight (PictureBox)
    Call DrawRadarMap(P1Sight, 1, 1)
    Call DrawRadarMap(P2Sight, 1, 2)
    
    'Call DrawAddMap
    Call AddLayerDrawMap(PhMap, OrMap.X, OrMap.Y, 1)
    
End Sub
