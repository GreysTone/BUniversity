Attribute VB_Name = "ModuleLogicIFLink"
Option Explicit

Public Sub IFLink(ByVal Id As Integer, ByVal i As Integer, ByVal Vid As Integer, ByVal From As Integer)
    
    Dim Value As Integer
    
    Select Case From
        Case 0
            Value = Val(Player(Id).ListMain(i).V(Vid))
        Case 1
            Value = Val(Player(Id).List1(i).V(Vid))
        Case 2
            Value = Val(Player(Id).List2(i).V(Vid))
        Case 3
            Value = Val(Player(Id).List3(i).V(Vid))
        Case 4
            Value = Val(Player(Id).List4(i).V(Vid))
        Case 5
            Value = Val(Player(Id).List5(i).V(Vid))
    End Select
    
    Select Case Value
        Case 0  'File1
            Call File1Done(Id)
        Case 1  'File2
            Call File2Done(Id)
        Case 2  'File3
            Call File3Done(Id)
        Case 3  'File4
            Call File4Done(Id)
        Case 4  'File5
            Call File5Done(Id)
    End Select
    
End Sub
