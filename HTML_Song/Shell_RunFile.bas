Attribute VB_Name = "API_RunFile"
Option Explicit

' ***************************************************************
' Shows how to shell from a VB App and open a URL.
' ***************************************************************

Private Declare Function GetActiveWindow Lib _
    "user32" () As Long

Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal _
    lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

    'SW_HIDE = 0
    'SW_NORMAL = 1
    'SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
    'SW_MAXIMIZE = 3
    'SW_SHOWNOACTIVATE = 4
    'SW_SHOW = 5
    'SW_MINIMIZE = 6
    'SW_SHOWMINNOACTIVE = 7
    'SW_SHOWNA = 8
    'SW_RESTORE = 9
Public Const SW_SHOWNORMAL = 1

Public Sub RunFile(ByVal mFile As String, mFilePath As String, RunStyle As Integer)
'sample call:      RunFile File1.FileName, File1.Path, SW_SHOWNORMAL

    Dim temp As Long
    Dim Msg As String
    Dim X As Long

    temp = GetActiveWindow()
    X = ShellExecute(temp, "Open", mFile, "", mFilePath, RunStyle)
    
    If X < 32 Then
        Select Case X
            Case 0
                Msg = "The file could not be run due to insufficient system memory or a corrupt program file"
            Case 2
                Msg = "File Not Found"
            Case 3
                Msg = "Invalid Path"
            Case 5
                Msg = "Sharing or protection error"
            Case 6
                Msg = "Separate data segments are required for each task "
            Case 8
                Msg = "Insufficient memory to run the program"
            Case 10
                Msg = "Incorrect Windows version"
            Case 11
                Msg = "Invalid Program File"
            Case 12
                Msg = "Program file requires a different operating System "
            Case 13
                Msg = "Program requires MS-DOS 4.0"
            Case 14
                Msg = "Unknown program file type"
            Case 15
                Msg = "Windows prgram does not support protected memory mode"
            Case 16
                Msg = "Invalid use of data segments when loading a second instance of a program"
            Case 19
                Msg = "Attempt to run a compressed program file"
            Case 20
                Msg = "Invalid dynamic link library"
            Case 21
                Msg = "Program requires Windows 32-bit extensions"
            Case 31
                Msg = mFilePath
                If Right(Msg, 1) <> "\" Then Msg = Msg + "\"
                Msg = Msg + mFile
                Shell "rundll32.exe shell32.dll,OpenAs_RunDLL " + Msg
        End Select

        If X <> 31 Then MsgBox Msg, vbCritical, "Error Message"

    End If
    
End Sub
Public Sub HyperLink(ByVal sURL As String)
Dim iRet As Long

On Error GoTo URL_Error

iRet = ShellExecute(0, vbNullString, sURL, vbNullString, "c:\", SW_SHOWNORMAL)
Exit Sub

'<error handler>
URL_Error:
    MsgBox "Couldn't open URL", vbCritical, "Open URL Error"
    Err = 0
'</error handler>
End Sub

