Attribute VB_Name = "MDLDS"
Public Dictionaryfile As String
Public m() As Variant
Public hc() As String
Dim i As Integer
Public db As Integer

Public Sub LoadFile()
Dim Temp As String
  Open App.Path + "\RCH.rdb" For Input As #1
  Do While Not EOF(1)
  
  Loop
  MsgBox ("配置文件“Des.rs!”文件不存在，请查看帮助。")
  End
End Sub

