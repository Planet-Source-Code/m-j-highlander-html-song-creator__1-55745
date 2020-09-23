Attribute VB_Name = "General_Functions"
Option Explicit

Global Const ATTR_READONLY = 1    'Read-only file
Global Const ATTR_VOLUME = 8  'Volume label
Global Const ATTR_ARCHIVE = 32    'File has changed since last back-up
Global Const ATTR_NORMAL = 0  'Normal files
Global Const ATTR_HIDDEN = 2  'Hidden files
Global Const ATTR_SYSTEM = 4  'System files
Global Const ATTR_DIRECTORY = 16  'Directory

Global Const ATTR_DIR_ALL = ATTR_DIRECTORY + ATTR_READONLY + ATTR_ARCHIVE + ATTR_HIDDEN + ATTR_SYSTEM
Global Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_READONLY Or ATTR_ARCHIVE
Global Const ATTR_ALL_FILES_EXCEPT_READONLY = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_ARCHIVE
Public Sub EnableAll(frmX As Form)
On Error Resume Next
Dim ctrlX As Control

    For Each ctrlX In frmX.Controls
            ctrlX.Enabled = True
    Next ctrlX
    
End Sub

Public Sub DisableAll(frmX As Form)
On Error Resume Next
Dim ctrlX As Control

    For Each ctrlX In frmX.Controls
        Select Case ctrlX.Name
            Case "lblExtracting", "Bar"
                'do nothing
            Case Else
                ctrlX.Enabled = False
        End Select
    Next ctrlX
    
End Sub

Function CreatePath(ByVal DestPath$) As Boolean
Dim BackPos As Integer, forePos As Integer
Dim temp As String
Dim sOriginalCurDir As String
Dim sOriginalCurDrive As String

    ' save current dir
    sOriginalCurDir = CurDir

    ' Add slash to end of path if not there already
    If Right$(DestPath$, 1) <> "\" Then
        DestPath$ = DestPath$ + "\"
    End If
          


    ' Change to the root dir of the drive
    On Error Resume Next
    ChDrive DestPath$
    If Err <> 0 Then GoTo ErrorOut
    ChDir "\"


    ' Attempt to make each directory, then change to it
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do While forePos <> 0
        temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)

        Err = 0
        MkDir temp$
        If Err <> 0 And Err <> 75 Then GoTo ErrorOut

        Err = 0
        ChDir temp$
        If Err <> 0 Then GoTo ErrorOut

        BackPos = forePos
        forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop
    'Restore CurDir & CurDrv
    ChDir sOriginalCurDir
    ChDrive sOriginalCurDir
    'Return
    CreatePath = True
    Exit Function
                 
'<ERROR>
ErrorOut:
    CreatePath = False
    Exit Function
'</ERROR>
End Function


Function RemoveSlash(ByVal sPath As String) As String

sPath = Trim(sPath)

If Right(sPath, 1) = "\" Then
    RemoveSlash = Left(sPath, Len(sPath) - 1)
Else
    RemoveSlash = sPath
End If


End Function

Function DirExists(sDir As String) As Boolean
Dim tmp As String
Dim iResult As Integer

If Trim(sDir) = "" Then
            DirExists = False
            Exit Function
End If


iResult = 0
If Dir$(sDir, ATTR_DIR_ALL) <> "" Then
    iResult = GetAttr(sDir) And ATTR_DIRECTORY
End If

If iResult = 0 Then   'Directory not found, or the passed argument is a filename not a directory
    DirExists = False
Else
    DirExists = True
End If


End Function

Function ExtractDirName(FileName As String) As String

'Extract the Directory name from a full file name
    Dim tmp$
    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractDirName = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    tmp = Left(FileName, PrevPos)
    If Right(tmp, 1) = "\" Then tmp = Left(tmp, Len(tmp) - 1)
    ExtractDirName = tmp
    
End Function

Function FileExists(sFile As String) As Boolean

If Trim(sFile) = "" Then
            FileExists = False
            Exit Function
End If

If Dir$(sFile, ATTR_ALL_FILES) = "" Then
    FileExists = False
Else
    FileExists = True
End If

End Function

Function GetAppPath()

If Len(App.Path) = 3 Then
    GetAppPath = App.Path
Else
    GetAppPath = App.Path + "\"
End If

End Function


