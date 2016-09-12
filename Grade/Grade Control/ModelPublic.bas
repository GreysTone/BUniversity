Attribute VB_Name = "ModelPublic"

Public UserName As String

Public Type Class
    ID As Integer
    Names As String
End Type

Public Type Term
    ID As Integer
    Names As String
    Time_Start As Date
    Time_End As Date
    Table_Class(200) As Boolean
End Type

Public Type Grade
    ID As Integer
    TermID As Integer
    ClassID As Integer
    Get As Integer
    All As Integer
    AveClass As Integer
    OrdClass As Integer
    AveGroup As Integer
    OrdGroup As Integer
End Type

Public Count_Class As Integer
Public Count_Term As Integer
Public Count_Grade As Long

Public State_Firstrun As Boolean
Public State_TermHave As Boolean
Public State_ClassHave As Boolean
Public State_GradeHave As Boolean

Public Terms(2100) As Term
Public Classes(200) As Class
Public Grades(10000100) As Grade

Public NowTerm As Integer

Public Sub Initiation()

    State_Firstrun = True
    State_TermHave = False
    State_ClassHave = False
    State_GradeHave = False
    
End Sub

Public Sub AddTerm()
    Dim IDs As Integer
    Dim Buffer As String
    Dim buffer1 As String
    Dim buffer2 As String
    
    IDs = GetFreeID(0, Count_Term)
    If IDs = 9999999 Then
        Exit Sub
    End If
    
    Terms(IDs).ID = IDs
    Terms(IDs).Names = InputBox("Enter Term's Name:", "Enter")
    Buffer = InputBox("Enter the Start Date," & vbCrLf & "Format: YYYY/MM/DD", "Enter", Date)
    If IsDate(Buffer) Then
        Terms(IDs).Time_Start = Buffer
    Else
        MsgBox "We found wrong format." & vbCrLf & vbCrLf & "You need redo it"
        Call ModelPublic.ErrorLog("Mod_Public", "WrongDateFormat", "Delete")
        Call DeleteTerm(IDs)
        Exit Sub
    End If
    Buffer = InputBox("Enter the Start Date," & vbCrLf & "Format: YYYY/MM/DD", "Enter", Date)
    If IsDate(Buffer) Then
        Terms(IDs).Time_End = Buffer
    Else
        MsgBox "We found wrong format." & vbCrLf & vbCrLf & "You need redo it"
        Call ModelPublic.ErrorLog("Mod_Public", "WrongDateFormat", "Delete")
        Call DeleteTerm(IDs)
        Exit Sub
    End If
    
    If Terms(IDs).Time_End < Terms(IDs).Time_Start Then
        MsgBox "We found the end date is earlier than the start date." & vbCrLf & "[" & Terms(IDs).Time_Start & "][" & Terms(IDs).Time_End & "]" & vbCrLf & vbCrLf & "You need redo it."
        Call ModelPublic.ErrorLog("Mod_Public", "End<Start", "Delete")
        Call ModelPublic.DeleteTerm(IDs)
        Exit Sub
    End If
    MsgBox "Add Class"
    State_TermHave = True
    
    SaveData
    Call ModelPublic.RefeshData
    
End Sub

Public Sub DeleteTerm(ByVal ID As Integer)
    
    Dim i As Integer
    Dim Temp As Integer
    
    Terms(ID).ID = 0
    SaveData
    
    If Count_Term = 0 Then State_TermHave = False
    
    Temp = 0
    For i = 1 To Count_Term
        If Terms(i).ID <> 0 Then Temp = Temp + 1
    Next i
    If Temp = 0 Then
        State_TermHave = False
    End If
    
    Call ModelPublic.RefeshData
    
End Sub

Public Sub AddClass()
    
    Dim IDs As Integer
    
    IDs = GetFreeID(1, Count_Class)
    If IDs = 9999999 Then
        Exit Sub
    End If

    Classes(IDs).ID = IDs
    Classes(IDs).Names = InputBox("Enter Class' Name:", "Enter")
    State_ClassHave = True
    
    SaveData
    Call ModelPublic.RefeshData
    
End Sub

Public Sub DeleteClass(ByVal ID As Integer)
    
    Dim i As Integer
    Dim Temp As Integer
    
    Classes(ID).ID = 0
    For i = 1 To Count_Term
        Terms(i).Table_Class(ID) = False
    Next i
    
    SaveData
    If Count_Class = 0 Then State_ClassHave = False
    
    Temp = 0
    For i = 1 To Count_Class
        If Classes(i).ID <> 0 Then Temp = Temp + 1
    Next i
    If Temp = 0 Then
        State_ClassHave = False
    End If
    
    Call ModelPublic.RefeshData
    
End Sub

Public Sub AddGrade()
    
    Dim IDs As Integer
    Dim Confirm As Boolean
    Dim k As Integer
    
    IDs = GetFreeID(2, Count_Grade)
    If IDs = 9999999 Then
        Exit Sub
    End If
    
    Grades(IDs).ID = IDs
    
    '# Term
    Confirm = False
    While (Confirm = False)
        Grades(IDs).TermID = Val(InputBox("Enter Term ID:", "Enter"))
        If Grades(IDs).TermID > Count_Term Then
            MsgBox "Out of Order (Term)"
        Else
            If Terms(Grades(IDs).TermID).ID = 0 Then
                MsgBox "0 Order (Term)"
            Else
                k = MsgBox(Terms(Grades(IDs).TermID).Names & "?", vbYesNo, "Confirm")
                If k = 6 Then
                    Confirm = True
                End If
            End If
        End If
    Wend
    
    '# Class
    Confirm = False
    While (Confirm = False)
        Grades(IDs).ClassID = Val(InputBox("Enter class ID:", "Enter"))
        If Grades(IDs).ClassID > Count_Class Then
            MsgBox "Out of Order (Class)"
        Else
            If Classes(Grades(IDs).ClassID).ID = 0 Then
                MsgBox "0 Order (Class)"
            Else
                k = MsgBox(Classes(Grades(IDs).TermID).Names & "?", vbYesNo, "Confirm")
                If k = 6 Then
                    Confirm = True
                End If
            End If
        End If
    Wend
    
    MsgBox "0 for Null."
    
    Grades(IDs).Get = Val(InputBox("Enter Get Grades:", , "Enter"))
    Grades(IDs).All = Val(InputBox("Enter Full Grades:", , "Enter"))
    Grades(IDs).AveClass = Val(InputBox("Enter Average Grades:", , "Enter"))
    Grades(IDs).OrdClass = Val(InputBox("Enter Order of Class:", , "Enter"))
    Grades(IDs).AveGroup = Val(InputBox("Enter Difficlut Setting:", , "Enter"))
    Grades(IDs).OrdGroup = Val(InputBox("Enter Order of Grade:", , "Enter"))
    
    State_GradeHave = True
    
    Call SaveGrades
    
End Sub

Public Sub DeleteGrade(ByVal ID As Integer)
    
    Dim i As Integer
    Dim Temp As Integer
    
    Grades(i).ID = 0
    
     Call SaveGrades
     
    If Count_Class = 0 Then State_ClassHave = False
    
    Temp = 0
    For i = 1 To Count_Grade
        If Grades(i).ID <> 0 Then Temp = Temp + 1
    Next i
    If Temp = 0 Then
        State_GradeHave = False
    End If
    
End Sub

Public Sub ModifyGrade(ByVal ID As Integer)
    
    Dim Confirm As Boolean
    
    '# Term
    Confirm = False
    While (Confirm = False)
        Grades(ID).TermID = Val(InputBox("Enter Term ID:", "Enter"))
        If Grades(ID).TermID > Count_Term Then
            MsgBox "Out of Order (Term)"
        Else
            If Terms(Grades(ID).TermID).ID = 0 Then
                MsgBox "0 Order (Term)"
            Else
                k = MsgBox(Terms(Grades(ID).TermID).Names & "?", vbYesNo, "Confirm")
                If k = 6 Then
                    Confirm = True
                End If
            End If
        End If
    Wend
    
    '# Class
    Confirm = False
    While (Confirm = False)
        Grades(ID).ClassID = Val(InputBox("Enter class ID:", "Enter"))
        If Grades(ID).ClassID > Count_Class Then
            MsgBox "Out of Order (Class)"
        Else
            If Classes(Grades(ID).ClassID).ID = 0 Then
                MsgBox "0 Order (Class)"
            Else
                k = MsgBox(Classes(Grades(ID).TermID).Names & "?", vbYesNo, "Confirm")
                If k = 6 Then
                    Confirm = True
                End If
            End If
        End If
    Wend
    
    MsgBox "0 for Null."
    
    Grades(ID).Get = Val(InputBox("Enter Get Grades:", , "Enter"))
    Grades(ID).All = Val(InputBox("Enter Full Grades:", , "Enter"))
    Grades(ID).AveClass = Val(InputBox("Enter Average Grades:", , "Enter"))
    Grades(ID).OrdClass = Val(InputBox("Enter Order of Class:", , "Enter"))
    Grades(ID).AveGroup = Val(InputBox("Enter Difficlut Setting:", , "Enter"))
    Grades(ID).OrdGroup = Val(InputBox("Enter Order of Grade:", , "Enter"))
    
    Call SaveGrades
    
End Sub

Public Sub SaveGrades()
    
    Dim i As Integer
    
    Open App.Path & "\Grade.rs!" For Output As #1
        Write #1, Count_Grade
        For i = 1 To Count_Grade
            Write #1, Grades(i).ID, Grades(i).TermID, Grades(i).ClassID
            Write #1, Grades(i).Get, Grades(i).All, Grades(i).AveClass, Grades(i).OrdClass
            Write #1, Grades(i).AveGroup, Grades(i).OrdGroup
        Next i
    Close #1
        
End Sub

Public Sub RefeshData()
Dim i As Integer

    FrmMain.lblState.Caption = "Refreshing Data"
    FrmMain.lblState.Visible = True
    
    If State_TermHave = False Then
        FrmMain.lblTerm = "Term: UnSet"
    Else
        '# Change lblState
        For i = 1 To Count_Term
            If Terms(i).ID <> 0 And Terms(i).Time_Start <= Date And Terms(i).Time_End >= Date Then
                FrmMain.lblTerm = "Term  <" & Terms(i).Names & ">"
                NowTerm = i
                Exit For
            End If
        Next i
    End If
    
    FrmMain.lblState.Visible = False
End Sub

Public Sub SaveData()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Temp_Count As Integer
    
    Open App.Path & "\Term.rs!" For Output As #1
            Write #1, Count_Term
            For i = 1 To Count_Term
                Write #1, Terms(i).Names, Terms(i).ID
                Write #1, Terms(i).Time_Start, Terms(i).Time_End
            Next i
        Close #1
    
    Open App.Path & "\Class.rs!" For Output As #1
            Write #1, Count_Class
            For i = 1 To Count_Class
                Write #1, Classes(i).Names, Classes(i).ID
                '# Get Temp_Count
                Temp_Count = Get_Temp_Count(i)
                Write #1, Temp_Count
                For j = 1 To Temp_Count
                    'Write #1, Temp_TermID
                    For k = 1 To Count_Term
                        If (Terms(k).ID <> 0) Then
                            If Terms(k).Table_Class(Classes(i).ID) = True Then
                                Write #1, k
                            End If
                        End If
                    Next k
                Next j
            Next i
        Close #1
        
End Sub

Private Function Get_Temp_Count(ByVal ClassID)
    Dim i As Integer
    
    Get_Temp_Count = 0
    
    For i = 1 To Count_Term
        If (Terms(i).ID <> 0) Then
            If Terms(i).Table_Class(Classes(ClassID).ID) = True Then
                Get_Temp_Count = Get_Temp_Count + 1
            End If
        End If
    Next i
    
End Function

Public Sub ErrorLog(ByVal SubName As String, ByVal ErrosInfo As String, ByVal Result As String)
    Dim FileNum As Integer
    FileNum = FreeFile
    Open App.Path & "\Error.rsg" For Append As #FileNum
        Write #FileNum, "[" & Date & "-" & Time & "]" & vbCrLf & "[System Catch] [" & ErrosInfo & "] On [" & SubName & "]" & " By [" & Result & "]"
    Close #FileNum
End Sub

Public Function GetFreeID(ByVal Subject As Integer, ByVal All_Count As Integer) As Integer
'# Subject = 0 - Term / 1 - Class / 2 - Grade
    Dim Counts(10000100) As Boolean
    Dim Selected As Boolean
    
    Dim i As Integer
    
    Select Case Subject
        Case 0      '# Term
            For i = 1 To All_Count
                Counts(Terms(i).ID) = True
            Next i
            
            For i = 1 To All_Count
                If Counts(i) = False Then
                    GetFreeID = i
                    Selected = True
                End If
            Next i
            
            If Selected = False Then
                GetFreeID = All_Count + 1
                Count_Term = Count_Term + 1
                
                '# FlowOver Control
                If Count_Term > 2000 Then
                    MsgBox "Too more Terms!"
                    GetFreeID = 9999999
                End If
            End If
            
        Case 1      '# Class
            For i = 1 To All_Count
                Counts(Classes(i).ID) = True
            Next i
            
            For i = 1 To All_Count
                If Counts(i) = False Then
                    GetFreeID = i
                    Selected = True
                End If
            Next i
            
            If Selected = False Then
                GetFreeID = All_Count + 1
                Count_Class = Count_Class + 1
                
                '# FlowOver Control
                If Count_Term > 180 Then
                    MsgBox "Too more Classes!"
                    GetFreeID = 9999999
                End If
            End If
            
        Case 2      '# Grade
            For i = 1 To All_Count
                Counts(Grades(i).ID) = True
            Next i
            
            For i = 1 To All_Count
                If Counts(i) = False Then
                    GetFreeID = i
                    Selected = True
                End If
            Next i
            
            If Selected = False Then
                GetFreeID = All_Count + 1
                Count_Grade = Count_Grade + 1
                
                '# FlowOver Control
                If Count_Term > 10000000 Then
                    MsgBox "Too more Grades!"
                    GetFreeID = 9999999
                End If
            End If
    End Select
End Function
