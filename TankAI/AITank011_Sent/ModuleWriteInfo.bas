Attribute VB_Name = "ModuleWriteInfo"
Option Explicit

Public Sub WriteInfo(ByVal PId As Integer, ByVal Cases As Integer, ByVal Files As Integer)
    
    Select Case Cases
        Case 0
            frmMain.ListProgress.AddItem "Player" & PId & " " & "判断"
        Case 1
            frmMain.ListProgress.AddItem "Player" & PId & " " & "移动"
        Case 2
            frmMain.ListProgress.AddItem "Player" & PId & " " & "攻击"
        Case 3
            frmMain.ListProgress.AddItem "Player" & PId & " " & "获取被攻击"
        Case 4
            frmMain.ListProgress.AddItem "Player" & PId & " " & "获取导弹来源"
        Case 5
            frmMain.ListProgress.AddItem "Player" & PId & " " & "获取任意可行通路"
        Case 6
            frmMain.ListProgress.AddItem "Player" & PId & " " & "获取近地对抗对象"
        Case 99
            frmMain.ListProgress.Clear
            frmMain.ListProgress.AddItem "-----Game Start-----"
        Case 100
            frmMain.ListProgress.AddItem "-----End Game-----"
    End Select
    
End Sub
