VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmGraph 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "FrmGraph"
   ClientHeight    =   6045
   ClientLeft      =   5265
   ClientTop       =   2475
   ClientWidth     =   8610
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
   ScaleHeight     =   6045
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSaveInfo 
      Caption         =   "Save Info"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddGraph 
      Caption         =   "Add Graph"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdGraph 
      Caption         =   "New Graph"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox InfoBox 
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmGraph.frx":0000
   End
   Begin VB.PictureBox Picbox 
      Height          =   2415
      Left            =   360
      ScaleHeight     =   2355
      ScaleWidth      =   7635
      TabIndex        =   5
      Top             =   2280
      Width           =   7695
   End
   Begin VB.ComboBox ComboSub 
      Height          =   435
      Left            =   5040
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.ComboBox ComboClass 
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1320
      Width           =   4575
   End
   Begin RSGradeControl.ACPRibbon ACPRibbon1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   1296
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5040
      TabIndex        =   3
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label lblClass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "FrmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim R As Integer
Dim G As Integer
Dim B As Integer

Private Sub cmdAddGraph_Click()
    
    '# (Range 200,200)
    Call Graph(Picbox, Picbox.Width, Picbox.Height, R, G, B)
    
End Sub

Private Sub cmdGraph_Click()
    
    '# Init
    Picbox.Cls
    
    '# (Range 200,200)
    Call Graph(Picbox, Picbox.Width, Picbox.Height, R, G, B)
    
End Sub

Private Sub cmdSaveInfo_Click()
    
    Dim FileName As String
    
    FileName = InputBox("FileName to Save")
    
    If FileName = "" Then
        Exit Sub
    Else
        Open App.Path & "\" & FileName & ".txt" For Append As #1
            Write #1, InfoBox.Text
        Close #1
        MsgBox "Save Over." & vbCrLf & "[" & App.Path & "\" & FileName & "_Data" & ".txt" & "]"
    End If
    
End Sub

Private Sub Form_GotFocus()
    
    Call Load
    
End Sub

Private Sub Load()
    
    Dim i As Integer
    
    If Count_Class = 0 Then
        MsgBox "No Class! Please add Class first."
        Me.Visible = False
    End If
    
    ComboClass.Clear
    
    For i = 0 To Count_Class
        If Classes(i).ID <> 0 Then ComboClass.AddItem ("<" & Classes(i).ID & "> " & Classes(i).Names)
    Next i
    
End Sub

Private Sub Form_Load()
    
    '# Show Caption of Form
    ACPRibbon1.Caption = "Graph"
    
    '# Repaint Ribbon
    ACPRibbon1.Refresh
    
    Call Load
    
    '# Load ComboSub
    ComboSub.Clear
    
    ComboSub.AddItem "<1> Get"
    ComboSub.AddItem "<2> Ave of Class"
    ComboSub.AddItem "<3> Order of Class"
    ComboSub.AddItem "<4> Ave of Grade"
    ComboSub.AddItem "<5> Order of Grade"
    
End Sub

