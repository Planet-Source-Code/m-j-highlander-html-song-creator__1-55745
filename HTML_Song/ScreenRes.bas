Attribute VB_Name = "Module1"
Option Explicit
         
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Const BITSPIXEL As Long = 12
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10


Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1


Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Sub TopMost(frm As Form)

SetWindowPos frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

End Sub

Public Function GetScreenBPP() As Integer
   
   'currHRes = GetDeviceCaps(hdc, HORZRES)
   'currVRes = GetDeviceCaps(hdc, VERTRES)
   'currBPP = GetDeviceCaps(hdc, BITSPIXEL)
   
GetScreenBPP = GetDeviceCaps(frmXSongMain.hdc, BITSPIXEL)
   
End Function
