Attribute VB_Name = "ModuleWriteInfo"
Option Explicit

Public Sub WriteInfo(ByVal PId As Integer, ByVal Cases As Integer, ByVal Files As Integer)
    
    Select Case Cases
        Case 0
            frmMain.ListProgress.AddItem "Player" & PId & " " & "�ж�"
        Case 1
            frmMain.ListProgress.AddItem "Player" & PId & " " & "�ƶ�"
        Case 2
            frmMain.ListProgress.AddItem "Player" & PId & " " & "����"
        Case 3
            frmMain.ListProgress.AddItem "Player" & PId & " " & "��ȡ������"
        Case 4
            frmMain.ListProgress.AddItem "Player" & PId & " " & "��ȡ������Դ"
        Case 5
            frmMain.ListProgress.AddItem "Player" & PId & " " & "��ȡ�������ͨ·"
        Case 6
            frmMain.ListProgress.AddItem "Player" & PId & " " & "��ȡ���ضԿ�����"
        Case 99
            frmMain.ListProgress.Clear
            frmMain.ListProgress.AddItem "-----Game Start-----"
        Case 100
            frmMain.ListProgress.AddItem "-----End Game-----"
    End Select
    
End Sub
