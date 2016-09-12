Attribute VB_Name = "ModuleGeneral"
Option Explicit

Public Mode As Integer
Public InsCode As Integer

Public Player(2) As Tanks
Public NewTank As Tanks

Public OrMap As Map

Public Count_Player As Integer

Public State_P2_On As Boolean
Public State_MapLoad As Boolean
Public State_P1Load As Boolean
Public State_P2Load As Boolean

Public Sub InitGame()

    'frmMain.cmdLoad.Enabled = True
    frmMain.cmdStart.Enabled = True
    frmMain.cmdBreak.Enabled = True
    'frmMain.txtDraw.Locked = False
    
    frmMain.lblP1Name.Caption = Player(1).Names
    frmMain.lblP2Name.Caption = Player(2).Names
    frmMain.P1Dir.Caption = "Wait..."
    frmMain.P2Dir.Caption = "Wait..."
    frmMain.P1FireDir.Caption = "Wait..."
    frmMain.P2FireDir.Caption = "Wait..."
    frmMain.P1Ata.Caption = "Wait..."
    frmMain.P2Ata.Caption = "Wait..."
    
    
    Call ModuleMap.DrawMap(frmMain.PhMap, OrMap.X, OrMap.Y, 1)
    P1Move = False
    P2Move = False
    P1Attack = False
    P2Attack = False
    
    frmMain.Picture1.Print "Load Over." & vbCrLf; "Wait for Instruction<Start>"
    
End Sub
