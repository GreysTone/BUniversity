Attribute VB_Name = "ModuleGetLockOn"
Option Explicit

Public Function GetLockOn(ByVal Playerk As Integer) As String
    
    Dim NX As Integer
    Dim NY As Integer
    
    NX = PositionMap.Obc(Playerk).X
    NY = PositionMap.Obc(Playerk).Y
    
    If MissileLockMap.Data(NX, NY).V = True Then
        GetLockOn = "1"
        If Playerk = 1 Then
            frmMain.P1Ata.BackColor = vbRed
            frmMain.P1Ata.ForeColor = vbWhite
            frmMain.P1Ata.Caption = "Under Attack!"
        Else
            frmMain.P2Ata.BackColor = vbRed
            frmMain.P2Ata.ForeColor = vbWhite
            frmMain.P2Ata.Caption = "Under Attack!"
        End If
    Else
        GetLockOn = "0"
        If Playerk = 1 Then
            frmMain.P1Ata.BackColor = &H8000000F
            frmMain.P1Ata.ForeColor = vbBlack
            frmMain.P1Ata.Caption = "None"
        Else
            frmMain.P2Ata.BackColor = &H8000000F
            frmMain.P2Ata.ForeColor = vbBlack
            frmMain.P2Ata.Caption = "None"
        End If
    End If
    
End Function
