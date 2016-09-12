Attribute VB_Name = "ModKillMe"
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long '获取自己的PID
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Sub KillMe()
Dim MyFilename As String
Dim tmp As String * 255
Dim l As Integer
l = GetModuleFileName(0, tmp, 255)
MyFilename = Mid(tmp, 1, l)
Shell "cmd /c ping 127.0.0.1 -n 1 && del """ & MyFilename & """", vbHide:     ExitProcess (0)
End Sub


