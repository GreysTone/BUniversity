Attribute VB_Name = "ModuleInstruction"
Option Explicit

Public Type Instructions
    Id As Integer
    'Cd As Integer
    Father As Integer           ' 0 for Logic / 1 for Action / 2 for Info
    Vars As Integer
    Return As Boolean
End Type

Public Type InList
    InsID As Integer
    V(10) As String
    R As String
End Type

Public Ins(10) As Instructions

Public Sub InitInstructions()    'Check Over

    '0 /IF
    Ins(0).Id = 0
    Ins(0).Father = 1
    Ins(0).Vars = 5
    Ins(0).Return = False
    
    '1 /Move
    Ins(1).Id = 1
    Ins(1).Father = 1
    Ins(1).Vars = 1
    Ins(1).Return = False
    
    '2 /Attack
    Ins(2).Id = 2
    Ins(2).Father = 1
    Ins(2).Vars = 1
    Ins(2).Return = False
    
    '3 /GetLockOn
    Ins(3).Id = 3
    Ins(3).Father = 2
    Ins(3).Vars = 0
    Ins(3).Return = True
    
    '4 /GetFireDirection
    Ins(4).Id = 4
    Ins(4).Father = 2
    Ins(4).Vars = 0
    Ins(4).Return = True
    
    '5 /GetFreeWay
    Ins(5).Id = 5
    Ins(5).Father = 2
    Ins(5).Vars = 0
    Ins(5).Return = True
    
    '6 /FindEnermy
    Ins(6).Id = 6
    Ins(6).Father = 2
    Ins(6).Vars = 0
    Ins(6).Return = True
    
End Sub
Public Sub AddNewlistGeneral()    'Check Over

    frmNewTank.ProgressBar1.Value = frmNewTank.ProgressBar1 + 1
    
End Sub

Public Sub CleanAll()    'Check Over

    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 5
        NewTank.CountFile(i) = 0
    Next i
    
    For i = 0 To 110
        NewTank.ListMain(i).InsID = 0
        NewTank.ListMain(i).R = ""
        For j = 0 To 10
            NewTank.ListMain(i).V(j) = ""
        Next j
    Next i
    
    For i = 0 To 60
        NewTank.List1(i).InsID = 0
        NewTank.List1(i).R = ""
        For j = 0 To 10
            NewTank.List1(i).V(j) = ""
        Next j
    Next i
    
    For i = 0 To 60
        NewTank.List2(i).InsID = 0
        NewTank.List2(i).R = ""
        For j = 0 To 10
            NewTank.List2(i).V(j) = ""
        Next j
    Next i
    
    For i = 0 To 60
        NewTank.List3(i).InsID = 0
        NewTank.List3(i).R = ""
        For j = 0 To 10
            NewTank.List3(i).V(j) = ""
        Next j
    Next i
    
    For i = 0 To 60
        NewTank.List4(i).InsID = 0
        NewTank.List4(i).R = ""
        For j = 0 To 10
            NewTank.List4(i).V(j) = ""
        Next j
    Next i
    
    For i = 0 To 60
        NewTank.List5(i).InsID = 0
        NewTank.List5(i).R = ""
        For j = 0 To 10
            NewTank.List5(i).V(j) = ""
        Next j
    Next i
    
End Sub

Public Function AddNewAiList(ByVal ListNumber)    'Check Over

    If ListNumber = 0 Then
        If NewTank.CountFile(0) + 1 > 100 Then
            AddNewAiList = -1
        Else
            Call AddNewlistGeneral
            NewTank.CountFile(0) = NewTank.CountFile(0) + 1
            AddNewAiList = NewTank.CountFile(0)
        End If
    Else
        If NewTank.CountFile(ListNumber) + 1 > 50 Then
            AddNewAiList = -1
        Else
            Call AddNewlistGeneral
            NewTank.CountFile(ListNumber) = NewTank.CountFile(ListNumber) + 1
            AddNewAiList = NewTank.CountFile(ListNumber)
        End If
    End If
        
End Function

Public Sub AddInstToList(ByVal List As Integer, ByVal Inst As Integer, ByVal Line)  'Check Over

    Select Case Inst
        Case 0
            frmNewTank.ListIns(List).AddItem Line & " IF <" & DialogAdd.Combo1.List(DialogAdd.Combo1.ListIndex) & "> " & DialogAdd.Combo2.List(DialogAdd.Combo2.ListIndex) & " <" & DialogAdd.Combo3.List(DialogAdd.Combo3.ListIndex) & "> , True:" & DialogAdd.ComboTrue.List(DialogAdd.ComboTrue.ListIndex) & ", False:" & DialogAdd.ComboFalse.List(DialogAdd.ComboFalse.ListIndex)
        Case 1
            frmNewTank.ListIns(List).AddItem Line & " Move <" & DialogAdd.Combo1.List(DialogAdd.Combo1.ListIndex) & ">"
        Case 2
            frmNewTank.ListIns(List).AddItem Line & " Attack <" & DialogAdd.Combo1.List(DialogAdd.Combo1.ListIndex) & ">"
        Case 3
            frmNewTank.ListIns(List).AddItem Line & " <" & DialogAdd.ComboRe.List(DialogAdd.ComboRe.ListIndex) & "> = GetLockOn"
        Case 4
            frmNewTank.ListIns(List).AddItem Line & " <" & DialogAdd.ComboRe.List(DialogAdd.ComboRe.ListIndex) & "> = GetFireFrom"
        Case 5
            frmNewTank.ListIns(List).AddItem Line & " <" & DialogAdd.ComboRe.List(DialogAdd.ComboRe.ListIndex) & "> = GetFreeWay"
        Case 6
            frmNewTank.ListIns(List).AddItem Line & " <" & DialogAdd.ComboRe.List(DialogAdd.ComboRe.ListIndex) & "> = GetFind"
    End Select
    
End Sub
