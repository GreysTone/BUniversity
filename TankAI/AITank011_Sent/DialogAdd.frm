VERSION 5.00
Begin VB.Form DialogAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "对话框标题"
   ClientHeight    =   3195
   ClientLeft      =   8325
   ClientTop       =   4725
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox ComboFile 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox ComboFalse 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox ComboTrue 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox ComboRe 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblFalse 
      AutoSize        =   -1  'True
      Caption         =   "假"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label lblTrue 
      AutoSize        =   -1  'True
      Caption         =   "真"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      Caption         =   "命令参数"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      Caption         =   "命令格式："
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   900
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "命令"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblReturn 
      AutoSize        =   -1  'True
      Caption         =   "接受返回值变量"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   885
      Width           =   1260
   End
End
Attribute VB_Name = "DialogAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()    'Check Over
    
    Unload Me
    
End Sub

Private Sub Form_Load()    'Check Over
    
    Dim i As Integer
    
    ComboRe.Clear
    For i = 1 To 10
        ComboRe.AddItem "变量" & i
    Next i
    
    Combo1.Clear
    For i = 1 To 10
        Combo1.AddItem "变量" & i
    Next i
    Combo1.AddItem "方向 <前>"
    Combo1.AddItem "方向 <后>"
    Combo1.AddItem "方向 <左>"
    Combo1.AddItem "方向 <右>"
    Combo1.AddItem "方向 <保持>"
    
    Combo2.Clear
    Combo2.AddItem "="
    Combo2.AddItem ">"
    Combo2.AddItem ">="
    Combo2.AddItem "<"
    Combo2.AddItem "<="
    Combo2.AddItem "<>/ !"
    
    Combo3.Clear
    For i = 1 To 10
        Combo3.AddItem "变量" & i
    Next i
    Combo3.AddItem "True"
    Combo3.AddItem "False"
    'Combo3.AddItem "数值"
    
    ComboTrue.Clear
    For i = 1 To 5
        ComboTrue.AddItem "文件" & i
    Next i
    
    ComboFalse.Clear
    For i = 1 To 5
        ComboFalse.AddItem "文件" & i
    Next i
    
    ComboFile.Clear
    ComboFile.AddItem "主体"
    For i = 1 To 5
        ComboFile.AddItem "文件" & i
    Next i
    ComboFile.ListIndex = 0
    
End Sub
Private Sub AddLink(ByVal InstructionCode)  'Check Over
    
    Dim AdId As Integer
    
    AdId = ModuleInstruction.AddNewAiList(ComboFile.ListIndex)
    If AdId = -1 Then
        MsgBox ComboFile.List(ComboFile.ListIndex) & " is Full."
        Exit Sub
    Else
        Select Case ComboFile.ListIndex
            Case 0
                NewTank.ListMain(AdId).InsID = InstructionCode
                NewTank.ListMain(AdId).V(1) = Combo1.ListIndex
                NewTank.ListMain(AdId).V(2) = Combo2.ListIndex
                NewTank.ListMain(AdId).V(3) = Combo3.ListIndex
                NewTank.ListMain(AdId).V(4) = ComboTrue.ListIndex
                NewTank.ListMain(AdId).V(5) = ComboFalse.ListIndex
                NewTank.ListMain(AdId).R = ComboRe.ListIndex
                Call AddInstToList(ComboFile.ListIndex, InstructionCode, AdId)
            Case 1
                NewTank.List1(AdId).InsID = InstructionCode
                NewTank.List1(AdId).V(1) = Combo1.ListIndex
                NewTank.List1(AdId).V(2) = Combo2.ListIndex
                NewTank.List1(AdId).V(3) = Combo3.ListIndex
                NewTank.List1(AdId).V(4) = ComboTrue.ListIndex
                NewTank.List1(AdId).V(5) = ComboFalse.ListIndex
                NewTank.List1(AdId).R = ComboRe.ListIndex
                Call AddInstToList(ComboFile.ListIndex, InstructionCode, AdId)
            Case 2
                NewTank.List2(AdId).InsID = InstructionCode
                NewTank.List2(AdId).V(1) = Combo1.ListIndex
                NewTank.List2(AdId).V(2) = Combo2.ListIndex
                NewTank.List2(AdId).V(3) = Combo3.ListIndex
                NewTank.List2(AdId).V(4) = ComboTrue.ListIndex
                NewTank.List2(AdId).V(5) = ComboFalse.ListIndex
                NewTank.List2(AdId).R = ComboRe.ListIndex
                Call AddInstToList(ComboFile.ListIndex, InstructionCode, AdId)
            Case 3
                NewTank.List3(AdId).InsID = InstructionCode
                NewTank.List3(AdId).V(1) = Combo1.ListIndex
                NewTank.List3(AdId).V(2) = Combo2.ListIndex
                NewTank.List3(AdId).V(3) = Combo3.ListIndex
                NewTank.List3(AdId).V(4) = ComboTrue.ListIndex
                NewTank.List3(AdId).V(5) = ComboFalse.ListIndex
                NewTank.List3(AdId).R = ComboRe.ListIndex
                Call AddInstToList(ComboFile.ListIndex, InstructionCode, AdId)
            Case 4
                NewTank.List4(AdId).InsID = InstructionCode
                NewTank.List4(AdId).V(1) = Combo1.ListIndex
                NewTank.List4(AdId).V(2) = Combo2.ListIndex
                NewTank.List4(AdId).V(3) = Combo3.ListIndex
                NewTank.List4(AdId).V(4) = ComboTrue.ListIndex
                NewTank.List4(AdId).V(5) = ComboFalse.ListIndex
                NewTank.List4(AdId).R = ComboRe.ListIndex
                Call AddInstToList(ComboFile.ListIndex, InstructionCode, AdId)
            Case 5
                NewTank.List5(AdId).InsID = InstructionCode
                NewTank.List5(AdId).V(1) = Combo1.ListIndex
                NewTank.List5(AdId).V(2) = Combo2.ListIndex
                NewTank.List5(AdId).V(3) = Combo3.ListIndex
                NewTank.List5(AdId).V(4) = ComboTrue.ListIndex
                NewTank.List5(AdId).V(5) = ComboFalse.ListIndex
                NewTank.List5(AdId).R = ComboRe.ListIndex
                Call AddInstToList(ComboFile.ListIndex, InstructionCode, AdId)
        End Select
    End If
    
End Sub

Private Sub OKButton_Click()    'Check Over

    Dim AdId As Integer
    
    Select Case InsCode
        Case 0 'IF
            If Combo1.ListIndex < 0 Or Combo2.ListIndex < 0 Or Combo3.ListIndex < 0 Or ComboTrue.ListIndex < 0 Or ComboFalse.ListIndex < 0 Or ComboFile.ListIndex < 0 Then
                MsgBox "None Select."
                Exit Sub
            End If
            If ComboTrue.ListIndex = ComboFalse.ListIndex Then
                MsgBox "Same File."
                Exit Sub
            End If
            Call AddLink(InsCode)
        Case 1 'Move
            If Combo1.ListIndex < 0 Or ComboFile.ListIndex < 0 Then
                MsgBox "None Select."
                Exit Sub
            End If
            Call AddLink(InsCode)
        Case 2 'Attack
            If Combo1.ListIndex < 0 Or ComboFile.ListIndex < 0 Then
                MsgBox "None Select."
                Exit Sub
            End If
            Call AddLink(InsCode)
        Case Else 'Need Return Options
            If ComboRe.ListIndex < 0 Or ComboFile.ListIndex < 0 Then
                MsgBox "None Select."
                Exit Sub
            End If
            Call AddLink(InsCode)
    End Select
    
    Unload Me
    
End Sub
