VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "RStudio Grade Control"
   ClientHeight    =   5055
   ClientLeft      =   3015
   ClientTop       =   2280
   ClientWidth     =   14160
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
   ScaleHeight     =   5055
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List_Grade 
      Height          =   2265
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   13815
   End
   Begin RSGradeControl.ACPRibbon ACPRibbon1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   3836
   End
   Begin VB.Image Image12 
      Height          =   480
      Left            =   9480
      Picture         =   "FrmMain.frx":0000
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image11 
      Height          =   480
      Left            =   8880
      Picture         =   "FrmMain.frx":08CA
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   8280
      Picture         =   "FrmMain.frx":1194
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image Image8 
      Height          =   435
      Left            =   7680
      Picture         =   "FrmMain.frx":1739
      Top             =   2160
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   6960
      Picture         =   "FrmMain.frx":1CEF
      Top             =   2160
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image7 
      Height          =   210
      Left            =   6480
      Picture         =   "FrmMain.frx":227A
      Top             =   2280
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   5760
      Picture         =   "FrmMain.frx":23C1
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4920
      Picture         =   "FrmMain.frx":274B
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4200
      Picture         =   "FrmMain.frx":2A55
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3360
      Picture         =   "FrmMain.frx":331F
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2640
      Picture         =   "FrmMain.frx":3BE9
      Top             =   2160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblState 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "State"
      Height          =   315
      Left            =   12000
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label lblTerm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Term: UnSet"
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
      Top             =   2160
      Width           =   1770
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   0
      Picture         =   "FrmMain.frx":3F73
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TestFile As String
Dim Selects As Integer

Private Sub Form_Load()
    
    Call ModelPublic.Initiation
    
    '# Set Circle Menu Button Picture
    Set ACPRibbon1.Picture = Image1.Picture

    '# Show Caption of Form
    ACPRibbon1.Caption = Me.Caption & "     |      Welcome <" & ModelPublic.UserName & ">"

    '# Show Button to Customize Menu
    ACPRibbon1.ShowCustomMenu = False

    '# Add TopButtons ---   ID - Capt. - Icons
    ACPRibbon1.AddTopButton "1", "New Grade", Image7.Picture
    ACPRibbon1.AddTopButton "2", "New Term", Image7.Picture
    ACPRibbon1.AddTopButton "3", "New Class", Image7.Picture

    '# Add Tabs ---   ID - Caption
    ACPRibbon1.AddTab "1", "Grade Manage"
    ACPRibbon1.AddTab "2", "Grade Evaluate"
    ACPRibbon1.AddTab "3", "Grade Grapher"
    ACPRibbon1.AddTab "4", "Designer"
    ACPRibbon1.AddTab "5", "Plugins"
    ACPRibbon1.AddTab "6", "Help"

    '# Add Cats ---   ID - Tab - Caption - ShowDialogButton
    ACPRibbon1.AddCat "1", "1", "Modify", False
    ACPRibbon1.AddCat "2", "1", "Delete", False
    ACPRibbon1.AddCat "3", "2", "Evaluate", False
    ACPRibbon1.AddCat "4", "3", "Grapher", False
    ACPRibbon1.AddCat "5", "4", "Manage", False
    ACPRibbon1.AddCat "6", "5", "System", False
    ACPRibbon1.AddCat "7", "5", "Users", False
    ACPRibbon1.AddCat "8", "6", "Copyright", False

    '# Add Button ---    ID - Cat - Capt. - Icons -   More Arrow   - ToolTip            12
    ACPRibbon1.AddButton "1", "1", "Add", Image10.Picture, False, "Add new grades."
    ACPRibbon1.AddButton "2", "1", "Modify", Image3.Picture, False, "Modify existent grades."
    
    ACPRibbon1.AddButton "3", "2", "Delete", Image12.Picture, False, "Delete existent grades."
    
    ACPRibbon1.AddButton "4", "3", "from Data", Image9.Picture, False, "Evaluate grades from existent data."
    ACPRibbon1.AddButton "5", "3", "from Self.", Image4.Picture, False, "Evaluate grades from self evaluate."
    
    ACPRibbon1.AddButton "6", "4", "Start Grapher", Image8.Picture, True, "Start Grapher Model."
    
    ACPRibbon1.AddButton "7", "5", "Term", Image2.Picture, True, "Manage Terms."
    ACPRibbon1.AddButton "8", "5", "Class", Image6.Picture, True, "Manage Classs."
    
    ACPRibbon1.AddButton "11", "6", "Refresh", Image5.Picture, False, "Plugins - Manual Refresh State."
    ACPRibbon1.AddButton "9", "7", "Clock", Image5.Picture, False, "Plugins - Clock ."
    
    ACPRibbon1.AddButton "10", "8", "Copy_RStudio.", Image11.Picture, False, "RStudio's Copyright."

    '# Repaint Ribbon
    ACPRibbon1.Refresh
    
    
    '# Check First Start
    TestFile = Dir(App.Path & "\Fllog.rs!")
    If TestFile = "" Then
        MsgBox "Table <Designer> First", , "The First Time"
        Open App.Path & "\Fllog.rs!" For Output As #1
            Write #1, "FS_Time Checked!"
        Close #1
        State_Firstrun = True
    Else
        State_Firstrun = False
    End If
    
    '# Dim Var
    Dim i As Integer
    Dim j As Integer
    Dim Temp_Count As Integer
    Dim Temp_TermID As Integer
    
    '# Input Term.rs!
    TestFile = Dir(App.Path & "\Term.rs!")
    If TestFile <> "" And State_Firstrun = True Then
        Selects = MsgBox("We found some data dovetail into another data." & vbCrLf & "On [Term.rs!]" & vbCrLf & vbCrLf & "Do you want to ignore it?", vbYesNo, "Error")
        If Selects = 7 Then
            MsgBox "Program will not go on until this problem won't happend."
            Call ModelPublic.ErrorLog("FrmMain", "Data Check Unpass", "EndProgram")
            End
        Else
            MsgBox "Ignore the problem."
            Call ModelPublic.ErrorLog("FrmMain", "Data Check Unpass", "Ignore")
            Open App.Path & "\Term.rs!" For Append As #1
            Close #1
        End If
    End If
    
    If TestFile <> "" Then
        Open App.Path & "\Term.rs!" For Input As #1
            Input #1, Count_Term
            For i = 1 To Count_Term
                Input #1, Terms(i).Names, Terms(i).ID
                Input #1, Terms(i).Time_Start, Terms(i).Time_End
            Next i
        Close #1
        
        If Count_Term > 0 Then State_TermHave = True
    Else
        Count_Term = 0
        State_TermHave = False
    End If
    
    
    '# Input Class.rs!
    TestFile = Dir(App.Path & "\Class.rs!")
    If TestFile <> "" And State_Firstrun = True Then
        Selects = MsgBox("We found some data dovetail into another data." & vbCrLf & "On [Class.rs!]" & vbCrLf & vbCrLf & "Do you want to ignore it?", vbYesNo, "Error")
        If Selects = 7 Then
            MsgBox "Program will not go on until this problem won't happend."
            Call ModelPublic.ErrorLog("FrmMain", "Data Check Unpass", "EndProgram")
            End
        Else
            MsgBox "Ignore the problem."
            Call ModelPublic.ErrorLog("FrmMain", "Data Check Unpass", "Ignore")
            Open App.Path & "\Class.rs!" For Append As #1
            Close #1
        End If
    End If
    
    If TestFile <> "" Then
        Open App.Path & "\Class.rs!" For Input As #1
            Input #1, Count_Class
            For i = 1 To Count_Class
                Input #1, Classes(i).Names, Classes(i).ID
                Input #1, Temp_Count
                For j = 1 To Temp_Count
                    Input #1, Temp_TermID
                    Terms(Temp_TermID).Table_Class(Classes(i).ID) = True
                Next j
            Next i
        Close #1
        
        If Count_Class > 0 Then State_ClassHave = True
    Else
        Count_Class = 0
        State_ClassHave = False
    End If
    
    '# Input Grade
    TestFile = Dir(App.Path & "\Grade.rs!")
    If TestFile <> "" And State_Firstrun = True Then
        Selects = MsgBox("We found some data dovetail into another data." & vbCrLf & "On [Grade.rs!]" & vbCrLf & vbCrLf & "Do you want to ignore it?", vbYesNo, "Error")
        If Selects = 7 Then
            MsgBox "Program will not go on until this problem won't happend."
            Call ModelPublic.ErrorLog("FrmMain", "Data Check Unpass", "EndProgram")
            End
        Else
            MsgBox "Ignore the problem."
            Call ModelPublic.ErrorLog("FrmMain", "Data Check Unpass", "Ignore")
            Open App.Path & "\Grade.rs!" For Append As #1
            Close #1
        End If
    End If
    
    If TestFile <> "" Then
        Open App.Path & "\Grade.rs!" For Input As #1
            Input #1, Count_Grade
            For i = 1 To Count_Grade
                Input #1, Grades(i).ID, Grades(i).TermID, Grades(i).ClassID
                Input #1, Grades(i).Get, Grades(i).All, Grades(i).AveClass, Grades(i).OrdClass
                Input #1, Grades(i).AveGroup, Grades(i).OrdGroup
            Next i
        Close #1
        
        If Count_Grade > 0 Then State_GradeHave = True
    Else
        Count_Grade = 0
        State_GradeHave = False
    End If
    
    '# RefreshData
    Call ModelPublic.RefeshData
    Call Refresh_List_Grade
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End
    
End Sub


Private Sub ACPRibbon1_MainMenuClick()
    
    '# This Event occurs on click in Main Button Menu
    FormAbout.Visible = True

End Sub

Private Sub ACPRibbon1_CustomClick()

    '# This Event occurs on click in Custom Button Menu
    'MsgBox "Custom Click"

End Sub

Private Sub ACPRibbon1_MenuClick(ByVal ID As String, ByVal Caption As String)

    '# This Event occurs when click on each Menu Button
    'MsgBox "MenuClick: " & ID & "--" & Caption
    
    Select Case Val(ID)
        Case 1
            If State_TermHave = True And State_ClassHave = True Then
                Call ModelPublic.AddGrade
                Call Refresh_List_Grade
            Else
                MsgBox "AddTerm or AddClass First"
            End If
        Case 2
            Call ModelPublic.AddTerm
        Case 3
            If State_TermHave = True Then
                Call ModelPublic.AddClass
            Else
                MsgBox "AddTerm First"
            End If
    End Select

End Sub

Private Sub ACPRibbon1_CatClick(ByVal ID As String, ByVal Caption As String)

    '# This Event occurs when click on each ShowDialogButton for each Categorie
    'MsgBox "ShowDialogClick: " & ID & "--" & Caption
    
    Select Case Val(ID)
    'Case 3
    '    MsgBox "# Evaluate"
    'Case 4
    '    MsgBox "# Grapher"
    'Case 6
    '    MsgBox "# System"
    'Case 7
    '    MsgBox "# User"
    End Select

End Sub
Private Sub Refresh_List_Grade()
    
    Dim i As Integer
    
    List_Grade.Clear
    
    For i = 1 To Count_Grade
        If Grades(i).ID <> 0 Then
            If Grades(i).TermID = NowTerm Then
                List_Grade.AddItem ("<" & i & ">     ID-" & Grades(i).ID & "     " & Terms(Grades(i).TermID).Names & "     " & Classes(Grades(i).ClassID).Names & "     Get/All-" & Grades(i).Get & "/" & Grades(i).All & "     Order/Ave[Class]-" & Grades(i).OrdClass & "/" & Grades(i).AveClass & "     Order/Dft[Grade]-" & Grades(i).OrdGroup & "/" & Grades(i).AveGroup)
            End If
        End If
    Next i
    
End Sub

Private Sub ACPRibbon1_ButtonClick(ByVal ID As String, ByVal Caption As String)

    '# This Event occurs when click on each Button
    'MsgBox "ButtonClick: " & ID & "--" & Caption
    Dim TempID As Integer
    
    Select Case Val(ID)
        Case 1
            If State_TermHave = True And State_ClassHave = True Then
                '# Add Grades
                Call ModelPublic.AddGrade
                Call Refresh_List_Grade
            Else
                MsgBox "AddTerm or AddClass First"
            End If
        Case 2
            If State_GradeHave = True Then
                '#Modify
                TempID = Val(InputBox("Enter GradeID of Modify", "Enter"))
                Call ModelPublic.ModifyGrade(TempID)
                Call Refresh_List_Grade
            Else
                MsgBox "No available grades"
            End If
        Case 3
            If State_GradeHave = True Then
                '#Delete
                TempID = Val(InputBox("Enter GradeID of Delete", "Enter"))
                Call ModelPublic.DeleteGrade(TempID)
                Call Refresh_List_Grade
            Else
                MsgBox "No available grades"
            End If
        Case 4
            If State_GradeHave = True Then
                '# from Data
                FrmEvaluate.Visible = True
            Else
                MsgBox "No available grades"
            End If
        Case 5
            If State_GradeHave = True Then
                '# from Self.
                FrmEvaSelf.Visible = True
            Else
                MsgBox "No available grades"
            End If
        Case 6
            If State_GradeHave = True Then
                '#  Grapher
                FrmGraph.Visible = True
            Else
                MsgBox "No available grades"
            End If
        Case 7
            FrmManageTerm.Visible = True
        Case 8
            If State_TermHave = True Then
                FrmManageClass.Visible = True
            Else
                MsgBox "AddTerm First"
            End If
        Case 9
            FrmClock_Plug.Visible = True
        Case 10
            FormAbout.Visible = True
        Case 11
            MsgBox "Manual Refresh State's Information.", , "Notification"
            Call ModelPublic.RefeshData
            MsgBox "Refresh over.", , "Notification"
    End Select

End Sub
