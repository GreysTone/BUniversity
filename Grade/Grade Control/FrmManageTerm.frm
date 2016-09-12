VERSION 5.00
Begin VB.Form FrmManageTerm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   3390
   ClientTop       =   5715
   ClientWidth     =   6495
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
   ScaleHeight     =   3255
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ListBox List_Exist 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Double Click to Load on Frame Modify"
      Top             =   1200
      Width           =   1575
   End
   Begin RSGradeControl.ACPRibbon ACPRibbon1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6480
      _ExtentX        =   11456
      _ExtentY        =   1296
   End
   Begin VB.Frame famModify 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modify"
      Height          =   2295
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   3135
      Begin VB.TextBox txtClasses 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   11
         Text            =   "Classes"
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtEndDate 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   10
         Text            =   "End Date"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtStartDate 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   9
         Text            =   "Start Date"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtNames 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         TabIndex        =   8
         Text            =   "TermName"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   6480
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   6480
      Y1              =   720
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   3240
   End
   Begin VB.Label lblS_Select 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblS_Have 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exist"
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   675
   End
End
Attribute VB_Name = "FrmManageTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    
    Call ModelPublic.AddTerm
    Load
    
End Sub

Private Sub cmdDelete_Click()
    
    Call ModelPublic.DeleteTerm(Val(Mid(List_Exist.List(List_Exist.ListIndex), 2, 10)))
    Load

End Sub

Private Sub Form_Load()
    
     '# Show Caption of Form
    ACPRibbon1.Caption = "Manage Term"
    
    '# Repaint Ribbon
    ACPRibbon1.Refresh
    
    Load
    
End Sub

Private Sub Load()
    Dim i As Integer
    
    List_Exist.Clear

    For i = 1 To Count_Term
        If Terms(i).ID <> 0 Then List_Exist.AddItem ("<" & Terms(i).ID & "> " & Terms(i).Names)
    Next i
    
    Call ModelPublic.SaveData
    
End Sub

Private Sub List_Exist_Click()

    lblS_Select.Caption = Terms(Val(Mid(List_Exist.List(List_Exist.ListIndex), 2, 10))).Names

End Sub

Private Sub List_Exist_DblClick()

    Dim i As Integer
    
    lblID.Caption = Terms(Val(Mid(List_Exist.List(List_Exist.ListIndex), 2, 10))).ID
    txtNames.Text = Terms(Val(Mid(List_Exist.List(List_Exist.ListIndex), 2, 10))).Names
    txtStartDate.Text = Terms(Val(Mid(List_Exist.List(List_Exist.ListIndex), 2, 10))).Time_Start
    txtEndDate.Text = Terms(Val(Mid(List_Exist.List(List_Exist.ListIndex), 2, 10))).Time_End
    txtClasses.Text = ""
    For i = 1 To Count_Class
        If Terms(Val(Mid(List_Exist.List(List_Exist.ListIndex), 2, 10))).Table_Class(i) = True Then
            txtClasses.Text = txtClasses.Text & i & ","
        End If
    Next i
    If txtClasses.Text <> "" Then
        txtClasses.Text = Left(txtClasses.Text, Len(txtClasses.Text) - 1)
    Else
        txtClasses.Text = "Null"
    End If
    
    txtNames.Enabled = True
    txtStartDate.Enabled = True
    txtEndDate.Enabled = True
    txtClasses.Enabled = True

End Sub

Private Sub txtEndDate_DblClick()
    
    If IsDate(txtEndDate.Text) = False Then
        MsgBox "We found wrong format." & vbCrLf & vbCrLf & "You need rewrite it"
        Call ModelPublic.ErrorLog("FrmManageTerm", "WrongDateFormat", "ReWrite")
    Else
        If IsDate(txtStartDate.Text) = False Then
            MsgBox "We found wrong format." & vbCrLf & vbCrLf & "You need rewrite it"
            Call ModelPublic.ErrorLog("FrmManageTerm", "WrongDateFormat", "ReWrite")
        Else
            If txtStartDate.Text > txtEndDate.Text Then
                MsgBox "We found the end date is earlier than the start date." & vbCrLf & "[" & txtStartDate.Text & "][" & txtEndDate.Text & "]" & vbCrLf & vbCrLf & "You need redo it."
                Call ModelPublic.ErrorLog("FrmManageTerm", "End<Start", "Delete")
            Else
                Terms(Val(lblID.Caption)).Time_End = txtEndDate.Text
                txtEndDate.Enabled = False
                MsgBox "Modify Over", , "Notification"
                Load
            End If
        End If
    End If
    
End Sub

Private Sub txtNames_DblClick()
    
        Terms(Val(lblID.Caption)).Names = txtNames.Text
        txtNames.Enabled = False
        MsgBox "Modify Over", , "Notification"
        Load

End Sub

Private Sub txtStartDate_DblClick()
    
    If IsDate(txtStartDate.Text) = False Then
        MsgBox "We found wrong format." & vbCrLf & vbCrLf & "You need rewrite it"
        Call ModelPublic.ErrorLog("FrmManageTerm", "WrongDateFormat", "ReWrite")
    Else
        If IsDate(txtEndDate.Text) = False Then
            MsgBox "We found wrong format." & vbCrLf & vbCrLf & "You need rewrite it"
            Call ModelPublic.ErrorLog("FrmManageTerm", "WrongDateFormat", "ReWrite")
        Else
            If txtStartDate.Text > txtEndDate.Text Then
                MsgBox "We found the end date is earlier than the start date." & vbCrLf & "[" & txtStartDate.Text & "][" & txtEndDate.Text & "]" & vbCrLf & vbCrLf & "You need redo it."
                Call ModelPublic.ErrorLog("FrmManageTerm", "End<Start", "Delete")
            Else
                Terms(Val(lblID.Caption)).Time_Start = txtStartDate.Text
                txtStartDate.Enabled = False
                MsgBox "Modify Over", , "Notification"
                Load
            End If
        End If
    End If
    
End Sub
