VERSION 5.00
Begin VB.Form frmNewAction 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "新建对战"
   ClientHeight    =   5820
   ClientLeft      =   5685
   ClientTop       =   4530
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdInit 
      Caption         =   "初始化系统"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   4920
      Width           =   4335
   End
   Begin VB.PictureBox MapVeiw 
      BackColor       =   &H00C0C0C0&
      Height          =   2055
      Left            =   2280
      ScaleHeight     =   1995
      ScaleWidth      =   1995
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.ComboBox ComboMode 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmNewAction.frx":0000
      Left            =   1920
      List            =   "frmNewAction.frx":0002
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   360
      Width           =   2655
   End
   Begin VB.PictureBox P1Sight 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   2040
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox P2Sight 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   3720
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblTank 
      AutoSize        =   -1  'True
      Caption         =   "载入坦克："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   9
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Label lblMapPath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "- 双击右侧框载入"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "载入地图："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label lblModel 
      AutoSize        =   -1  'True
      Caption         =   "对战模式："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   1200
   End
   Begin VB.Label lblP1 
      AutoSize        =   -1  'True
      Caption         =   "P1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   3240
      Width           =   300
   End
   Begin VB.Label lblP2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "P2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4380
      TabIndex        =   2
      Top             =   3240
      Width           =   300
   End
End
Attribute VB_Name = "frmNewAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Path As String
Dim TestFile As String

Private Sub cmdInit_Click()
    
      If ComboMode.ListIndex <> 0 Then  ' Only "UnderAttack"
        MsgBox "Mode haven't been selected."
        Exit Sub
    End If
    
    If State_MapLoad = False Then
        MsgBox "Map haven't been loaded."
        Exit Sub
    End If
    
    If State_P1Load = False Then
        MsgBox "Player1 haven't been loaded."
        Exit Sub
    End If
    
    If State_P2Load = False Then
        MsgBox "Player2 haven't been loaded."
        Exit Sub
    End If
    
    Mode = 1    ' Under Attack
    Call InitGame
    
    frmMain.Visible = True
    Unload Me
    
End Sub

Private Sub Form_Load()    'Check Over

    ComboMode.Clear
    ComboMode.AddItem "遭受攻击"
    ComboMode.ListIndex = 0
    State_P1Load = False
    State_P2Load = False
    State_MapLoad = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)  'Check Over
    
    frmMain.Visible = True
    Unload Me
    
End Sub

Private Sub MapVeiw_DblClick()    'Check Over
    
    Dim CheckHead As String
    Dim i As Integer
    Dim j As Integer
    
    Path = InputBox("App.path\Map\", "Load Map", "")
    If Path = "" Then Exit Sub
    
    TestFile = Dir(App.Path & "\Map\" & Path)
    If TestFile = "" Then
        MsgBox "Cannot find this Map [" & Path & "]." & vbCrLf & " Please check.", vbOKOnly
        Exit Sub
    Else
        '# Load Map
        Open App.Path & "\Map\" & Path For Input As #1
        
        If EOF(1) <> True Then  '# End of file?
            Input #1, CheckHead
            If CheckHead <> "[RS TANK MAP DATA]" Then
                MsgBox "Cannot read this Map [" & Path & "]." & vbCrLf & " Please check.", vbOKOnly
                Close #1
                Exit Sub
            Else '# Read X, Y, Players
                If EOF(1) <> True Then '# End of file?
                    Input #1, OrMap.X, OrMap.Y, OrMap.Players
                    '# PlayerSet
                    If OrMap.Players < 2 Then
                        State_P2_On = False
                        Count_Player = OrMap.Players
                        lblP2.Caption = "Null"
                        P2Sight.BackColor = RGB(255, 255, 255)
                        P2Sight.Enabled = False
                    Else
                        Count_Player = OrMap.Players
                    End If
                    
                    '# Big Map
                    If OrMap.X > 100 Or OrMap.Y > 100 Then
                        MsgBox "Sizeof(Map) is too big."
                        Close #1
                        Exit Sub
                    End If
                    '# Read Info
                    For i = 1 To OrMap.X
                        For j = 1 To OrMap.Y
                            If EOF(1) <> True Then
                                Input #1, OrMap.Info(i, j)
                                If OrMap.Info(i, j) = 0 Then
                                    PositionMap.Data(i, j) = 1
                                Else
                                    PositionMap.Data(i, j) = 0
                                End If
                                '# Load Tank Init Position
                                If OrMap.Info(i, j) > 4 Then
                                    OrMap.Pos(OrMap.Info(i, j) - 4).X = i
                                    OrMap.Pos(OrMap.Info(i, j) - 4).Y = j
                                    PositionMap.Data(i, j) = OrMap.Info(i, j) - 3
                                    PositionMap.Obc(PositionMap.Data(i, j) - 1).X = i
                                    PositionMap.Obc(PositionMap.Data(i, j) - 1).Y = j
                                    '# 是否为指定起始点地图
                                    OrMap.MsPos = True
                                End If
                            Else
                                MsgBox "Cannot read this Map [" & Path & "]." & vbCrLf & " Please check.", vbOKOnly
                                Close #1
                                Exit Sub
                            End If
                        Next j
                    Next i
                    Close #1
                Else
                    MsgBox "Cannot read this Map [" & Path & "]." & vbCrLf & " Please check.", vbOKOnly
                    Close #1
                End If
            End If
        Else
            MsgBox "Cannot read this Map [" & Path & "]." & vbCrLf & " Please check.", vbOKOnly
            Close #1
        End If
    End If
    
    'For i = 1 To OrMap.Players
    '    MsgBox "P" & i & "<X,Y> <" & OrMap.Pos(i).X & "," & OrMap.Pos(i).Y & ">"
    'Next i
    
    OrMap.Names = Path
    lblMapPath.Caption = "- " & UCase(OrMap.Names)
    If OrMap.MsPos = True Then lblMapPath.Caption = lblMapPath.Caption & vbCrLf & "给定初始坐标"
    lblP1.Caption = "P1 <" & OrMap.Pos(1).X & "," & OrMap.Pos(1).Y & ">"
    lblP2.Caption = "<" & OrMap.Pos(2).X & "," & OrMap.Pos(2).Y & "> P2"

    '# Draw Map
    Call ModuleMap.DrawMap(frmNewAction.MapVeiw, OrMap.X, OrMap.Y, 1)
    
    State_MapLoad = True
    
End Sub

Private Sub P1Sight_DblClick()  'Check Over
    
    Dim Path As String
    Dim Temp As String
    Dim CheckUnder As Integer
    
    Path = InputBox("App.path\Tank\", "Load AI Tank", "")
    If Path = "" Then Exit Sub
    CheckUnder = CheckReadAiTank(Path, 1) 'Check & Load
    
    If CheckUnder = -1 Then
        MsgBox "Cannot load this tank."
        Exit Sub
    End If
    
    Player(1).PId = 1
    P1Sight.BackColor = RGB(Player(1).Color.R, Player(1).Color.G, Player(1).Color.b)
    Temp = lblP1.Caption
    lblP1.Caption = "重绘地图"
    Call ModuleMap.DrawMap(frmNewAction.MapVeiw, OrMap.X, OrMap.Y, 1)
    lblP1.Caption = Temp
    State_P1Load = True
    
End Sub

Private Sub P2Sight_DblClick()
    
    Dim Path As String
    Dim Temp As String
    Dim CheckUnder As Integer
    
    Path = InputBox("App.path\Tank\", "Load AI Tank", "")
    If Path = "" Then Exit Sub
    CheckUnder = CheckReadAiTank(Path, 2) 'Check & Load
    
    If CheckUnder = -1 Then
        MsgBox "Cannot load this tank."
        Exit Sub
    End If
    
    Player(1).PId = 2
    P2Sight.BackColor = RGB(Player(2).Color.R, Player(2).Color.G, Player(2).Color.b)
    Temp = lblP2.Caption
    lblP2.Caption = "重绘地图"
    Call ModuleMap.DrawMap(frmNewAction.MapVeiw, OrMap.X, OrMap.Y, 1)
    lblP2.Caption = Temp
    State_P2Load = True
    
End Sub
