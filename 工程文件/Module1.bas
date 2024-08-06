Attribute VB_Name = "API_32"
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal strPath As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Sub SleepApi32test()
    MsgBox "接下来将静止3秒！"
    Sleep 3000
    MsgBox "静止解除！"
End Sub

Sub 发送消息测试()
    Dim myhwnd&
    myhwnd = FindWindow("OpusApp", vbNullString)
    PostMessage myhwnd, &H111, 8, 0
End Sub
