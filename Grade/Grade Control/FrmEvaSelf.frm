VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEvaSelf 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   7530
   ClientTop       =   4575
   ClientWidth     =   6255
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
   ScaleHeight     =   4245
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ComboBox ComboTerm 
      Height          =   435
      Left            =   480
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton cmdEva 
      Caption         =   "Evaluate"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProBar1 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox Resbox 
      Height          =   1335
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2355
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmEvaSelf.frx":0000
   End
   Begin RSGradeControl.ACPRibbon ACPRibbon1 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   1296
   End
   Begin VB.Label lblTermE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Term Source:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1875
   End
   Begin VB.Label lblSign 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class Count : No Progress"
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   2955
   End
End
Attribute VB_Name = "FrmEvaSelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEva_Click()
    
    Dim i As Integer
    Dim TId As Integer
    Dim Count As Integer
    Dim EachVal As Integer
    Dim GradeTemp As Single
    Dim GId As Integer
    
    'Init
    Resbox.Text = ""
    ProBar1.Value = 0
    ProBar1.Min = 0
    ProBar1.Max = 110
    ProBar1.Visible = True
    Count = 0
    
    'Count
    TId = Val(Mid(ComboTerm.List(ComboTerm.ListIndex), 2, 10))
    For i = 1 To 180
        If Terms(TId).Table_Class(i) = True Then Count = Count + 1
        If i Mod 5 = 0 Then ProBar1.Value = ProBar1.Value + 1
    Next i  ' Value = 34
    lblSign.Caption = "Class Count : " & Count
    ProBar1.Value = ProBar1.Value + 1   ' Value = 35
    
    'Clac
    EachVal = (ProBar1.Max - ProBar1.Value) / Count
    For i = 1 To 180
        If Terms(TId).Table_Class(i) = True Then
            GId = GetGradesId(TId, i)
            Resbox.Text = Resbox.Text & "========" & vbCrLf
            Resbox.Text = Resbox.Text & Classes(i).ID & "  " & Classes(i).Names & vbCrLf
            GradeTemp = Val(InputBox("Self Evaluate Data for " & Classes(i).Names))
            If GradeTemp <> 0 Then
                GradeTemp = GradeTemp * 0.01 + Grades(GId).AveClass * 0.001 + Grades(GId).AveGroup * 0.002
            Else
                GradeTemp = 0
            End If
            GradeTemp = Int(GradeTemp * 100)
            If GradeTemp > Grades(GId).All Then GradeTemp = Grades(GId).All
            Resbox.Text = Resbox.Text & Str(GradeTemp) & vbCrLf & "[" & Grades(GId).AveClass * 0.001 & "][" & Grades(GId).AveGroup * 0.002 & "]" & vbCrLf
            ProBar1.Value = ProBar1.Value + EachVal
        End If
    Next i
    Resbox.Text = Resbox.Text & "========" & vbCrLf
    
    'Close
    ProBar1.Value = ProBar1.Max
    ProBar1.Visible = False
    
End Sub

Private Sub Form_GotFocus()
    
    Call Load
    
End Sub

Private Sub Load()
    
    Dim i As Integer
    
    If Count_Term = 0 Then
        MsgBox "No Term! Please add term first."
        Me.Visible = False
    End If
    
    ComboTerm.Clear
    
    For i = 0 To Count_Term
        If Terms(i).ID <> 0 Then ComboTerm.AddItem ("<" & Terms(i).ID & "> " & Terms(i).Names)
    Next i
    
End Sub

Private Sub cmdSave_Click()
    
    Dim FileName As String
    
    FileName = InputBox("FileName to Save")
    
    If FileName = "" Then
        Exit Sub
    Else
        Open App.Path & "\" & FileName & ".txt" For Append As #1
            Write #1, Resbox.Text
        Close #1
        MsgBox "Save Over." & vbCrLf & "[" & App.Path & "\" & FileName & "_Self" & ".txt" & "]"
    End If
    
End Sub

Private Sub Form_Load()
    
    '# Show Caption of Form
    ACPRibbon1.Caption = "Evaluate from Self."
    
    '# Repaint Ribbon
    ACPRibbon1.Refresh
    
    Call Load
    
End Sub
