VERSION 5.00
Begin VB.Form FrmManageClass 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "ManageClass"
   ClientHeight    =   3600
   ClientLeft      =   10170
   ClientTop       =   5340
   ClientWidth     =   6915
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
   ScaleHeight     =   3600
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox List_Class 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Double Click to Load on Frame Modify"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
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
      Height          =   2160
      ItemData        =   "FrmManageClass.frx":0000
      Left            =   3720
      List            =   "FrmManageClass.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Double Click to Load on Frame Modify"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin RSGradeControl.ACPRibbon ACPRibbon1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   1296
   End
   Begin VB.Label lblS_Class 
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
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lblS_Term 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Term"
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
      TabIndex        =   5
      Top             =   840
      Width           =   750
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
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   675
   End
End
Attribute VB_Name = "FrmManageClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Load()

    Dim i As Integer
    
    If Count_Term = 0 Then
        MsgBox "No Term! Please add term first."
        Me.Visible = False
    End If
    
    Combo1.Clear
    
    For i = 0 To Count_Term
        If Terms(i).ID <> 0 Then Combo1.AddItem ("<" & Terms(i).ID & "> " & Terms(i).Names)
    Next i
    
End Sub

Private Sub Refresh_List_Class()
    
    Dim i As Integer
    
    List_Class.Clear
    
    For i = 1 To Count_Class
        If Classes(i).ID <> 0 Then List_Class.AddItem ("<" & Classes(i).ID & ">" & Classes(i).Names)
    Next i
    
End Sub

Private Sub Refresh_List_Exist(ByVal TermID As Integer)
    
    Dim i As Integer
    
    List_Exist.Clear
    
    For i = 1 To 180
        If Terms(TermID).Table_Class(i) = True Then List_Exist.AddItem ("<" & Classes(i).ID & ">" & Classes(i).Names)
    Next i
    
End Sub

Private Sub cmdAdd_Click()
    Dim i As Integer
    
    Dim TempID1 As Integer
    Dim TempID2 As Integer
    TempID1 = Val(Mid(List_Class.List(List_Class.ListIndex), 2, 10))
    
    If Val(Mid(Combo1.List(Combo1.ListIndex), 2, 10)) = 0 Then
        MsgBox "None Term Choise"
        Exit Sub
    End If
    
    For i = 0 To List_Exist.ListCount - 1
        TempID2 = Val(Mid(List_Exist.List(i), 2, 10))

        If (TempID1 = TempID2) Then
            MsgBox "Exist!"
            Exit Sub
        End If
    Next i
    
    List_Exist.AddItem ("<" & Classes(TempID1).ID & ">" & Classes(TempID1).Names)
    Terms(Val(Mid(Combo1.List(Combo1.ListIndex), 2, 10))).Table_Class(TempID1) = True
    
    Call ModelPublic.SaveData
    Call Refresh_List_Exist(Val(Mid(Combo1.List(Combo1.ListIndex), 2, 10)))
    
End Sub

Private Sub cmdDelete_Click()
    Dim TempID As Integer
    
    TempID = Val(Mid(List_Class.List(List_Class.ListIndex), 2, 10))
    
    If TempID <> 0 Then
        Call ModelPublic.DeleteClass(TempID)
    End If
    Call Refresh_List_Class
    Refresh_List_Exist (Val(Mid(Combo1.List(Combo1.ListIndex), 2, 10)))
    
End Sub

Private Sub cmdNew_Click()
    
    Call ModelPublic.AddClass
    Call Refresh_List_Class
    
End Sub

Private Sub cmdRemove_Click()
    Dim i As Integer
    
    Dim TempID2 As Integer
    TempID2 = Val(Mid(List_Exist.List(List_Exist.ListIndex), 2, 10))
    
    If TempID2 <> 0 Then
        List_Exist.RemoveItem (List_Exist.ListIndex)
        Terms(Val(Mid(Combo1.List(Combo1.ListIndex), 2, 10))).Table_Class(TempID2) = False
    End If
    
    Call ModelPublic.SaveData
    Call Refresh_List_Exist(Val(Mid(Combo1.List(Combo1.ListIndex), 2, 10)))
    
End Sub

Private Sub Combo1_Click()
    
    Dim TempID As String
    
    TempID = Combo1.List(Combo1.ListIndex)
    Call Refresh_List_Exist(Val(Mid(TempID, 2, 10)))
    
End Sub

Private Sub Form_GotFocus()
    
    Call Load
    Call Refresh_List_Class
    
End Sub

Private Sub Form_Load()
    
    '# Show Caption of Form
    ACPRibbon1.Caption = "Manage Class"
    
    '# Repaint Ribbon
    ACPRibbon1.Refresh
    
    Call Load
    Call Refresh_List_Class
    
End Sub
