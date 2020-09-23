Attribute VB_Name = "Module1"
Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOSIZE = &H1

Sub TopMost(frm As Form)
Dim WindowHandle As Integer

WindowHandle = frm.hWnd
SetWindowPos WindowHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

