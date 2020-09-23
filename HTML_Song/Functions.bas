Attribute VB_Name = "Functions"
Option Explicit
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Type FontSpecs
    Face As String
    Size As String
    Color As String
    Bold As Boolean
    Italic As Boolean
End Type
Public Type Margins
    LeftMargin As String
    RightMargin As String
    TopMargin As String
End Type

'''''''''''''''''''''''''''''''''''PUBLIC VARS
Public LyricsFont As FontSpecs
Public ArtistFont As FontSpecs
Public SongFont As FontSpecs
Public LyricsMargins As Margins
Public UpMargins As Margins
Public UpperFrameHeight As Integer
Public OutputFolder As String
Public SongTitle As String
Public ArtistName As String
Public LyricsBGColor As String
Public UpBGColor As String
Public SongFile As String
Public MediaCtlWidth As String
Public MediaCtlHeight As String

Public Function CreatePath(ByVal Path As String) As Boolean
On Error Resume Next
    
Dim v As Variant
Dim idx As Integer
Dim sFolder As String
Dim lower As Integer, upper As Integer

If Right$(Path, 1) = "\" Then Path = Left$(Path, Len(Path) - 1)
v = Split(Path, "\")
If IsArray(v) Then
    lower = LBound(v)
    upper = UBound(v)
    sFolder = v(lower) & "\" & v(lower + 1) ' drive + first folder

    MkDir sFolder
    For idx = lower + 2 To upper
        sFolder = sFolder & "\" & v(idx)
        MkDir sFolder
    Next
End If

If DirExists(sFolder) Then
    CreatePath = True
Else
    CreatePath = False
End If

End Function
Function HeuristicSongName(ByVal SongFileName As String) As String()
Dim v As Variant, idx As Integer, i As Integer
Dim sTemp As String, iPos As Integer
Dim sArray(0 To 1) As String

SongFileName = ExtractFileName(SongFileName)

iPos = InStrRev(SongFileName, ".")
If iPos > 0 Then
    SongFileName = Left$(SongFileName, iPos - 1)
End If

v = Split(SongFileName, "-")

sArray(0) = Trim$(v(LBound(v)))
sArray(1) = Trim$(v(UBound(v)))

If sArray(0) = sArray(1) Then sArray(0) = ""
HeuristicSongName = sArray

End Function
Public Sub Main()
'''''''''''''''''''''''''''''''''''DEFAULT VALS
LyricsFont.Face = "Verdana,Tahoma,Arial,Times"
LyricsFont.Size = "2"
LyricsFont.Color = "#000000"
LyricsFont.Bold = False
LyricsFont.Italic = False
ArtistFont.Face = "Tahoma,Verdana,Arial,Times"
ArtistFont.Size = "4"
ArtistFont.Color = "#000000"
ArtistFont.Bold = True
ArtistFont.Italic = False
SongFont.Face = "Times,Verdana,Tahoma,Arial"
SongFont.Size = "6"
SongFont.Color = "#000000"
SongFont.Bold = True
SongFont.Italic = True
LyricsMargins.LeftMargin = "12"
LyricsMargins.RightMargin = "12"
LyricsMargins.TopMargin = "8"
UpMargins.LeftMargin = "12"
UpMargins.RightMargin = "12"
UpMargins.TopMargin = "0"
UpperFrameHeight = "60"
OutputFolder = ""
SongTitle = ""
ArtistName = ""
LyricsBGColor = "#000000"
UpBGColor = "#000000"
SongFile = ""
MediaCtlWidth = "150"
MediaCtlHeight = "45"

Load frmXSongMain

frmXSongMain.Show

End Sub

Public Function RemoveInvalidChars(ByVal FileName As String) As String
'invalid chars: \/:*?"<>|

FileName = Replace$(FileName, "\", "-")
FileName = Replace$(FileName, "/", "-")
FileName = Replace$(FileName, ":", "-")
FileName = Replace$(FileName, "*", "+")
FileName = Replace$(FileName, "?", "!")
FileName = Replace$(FileName, "<", "")
FileName = Replace$(FileName, ">", "")
FileName = Replace$(FileName, "|", "-")
FileName = Replace$(FileName, """", "'")

RemoveInvalidChars = FileName

End Function
Public Function AppPath() As String

If Right$(App.Path, 1) <> "\" Then
    AppPath = App.Path & "\"
Else
    AppPath = App.Path
End If

End Function
Function HexColorToVBHex(RGB_Color As String) As Long
On Error GoTo Err_Hex
Dim HexRGB As String
Dim HexR$, HexG$, HexB$

HexRGB = RGB_Color
If Len(HexRGB) < 6 Then HexRGB = String(6 - Len(HexRGB), "0") + HexRGB


HexR = Right(HexRGB, 2)
HexG = Mid(HexRGB, 3, 2)
HexB = Left(HexRGB, 2)

HexColorToVBHex = CLng("&H" & HexR & HexG & HexB)

Exit Function
Err_Hex:
    Err = 0
    HexColorToVBHex = 0

End Function

Function LongColorToHex(RGB_Color As Long) As String
Dim HexRGB As String
Dim HexR$, HexG$, HexB$

HexRGB = Hex$(RGB_Color)
If Len(HexRGB) < 6 Then HexRGB = String(6 - Len(HexRGB), "0") + HexRGB

'Reverse VB Hex to get HTML Hex
HexR = Right(HexRGB, 2)
HexG = Mid(HexRGB, 3, 2)
HexB = Left(HexRGB, 2)

LongColorToHex = HexR & HexG & HexB

End Function
Function GetWinDir() As String
    
    Dim WinDir As String
    Dim File As String
    Dim Res As Long
    WinDir = Space$(20)
    Res = GetWindowsDirectory(WinDir, 20)
    File = Left$(WinDir, InStr(1, WinDir, Chr$(0)) - 1)
    GetWinDir = Trim$(File) & "\"
    
End Function


Public Function IsHTML(ByVal Text As String) As Boolean
Dim iResult As Long

Text = LCase$(Text)
iResult = 0
iResult = InStr(Text, "<")
If iResult > 0 Then iResult = InStr(Text, ">")

If iResult = 0 Then
    IsHTML = False
Else
    IsHTML = True
End If

End Function
Public Function AddSlash(ByVal Path As String) As String

If Right$(Path, 1) <> "\" Then
    AddSlash = Path & "\"
Else
    AddSlash = Path
End If

End Function
Function ReplaceChars(ByVal astr As String, ByVal ReplaceWith As String, ByVal UnwantedChars As String) As String
' For filenames: ReplaceChars(file_name, "", "\/:*?<>|" + Chr$(34)))

Dim TmpStr As String
Dim ch As String
Dim i As Integer

TmpStr = ""

For i = 1 To Len(UnwantedChars)
    ch = Mid$(UnwantedChars, i, 1)
    If ch = "!" Then ch = ""
    TmpStr = TmpStr + ch
Next i
UnwantedChars = TmpStr

TmpStr = ""
ch = ""

If Left(UnwantedChars, 1) <> "[" Then UnwantedChars = "[" + UnwantedChars
If Right(UnwantedChars, 1) <> "]" Then UnwantedChars = UnwantedChars + "]"

For i = 1 To Len(astr)
    ch = Mid$(astr, i, 1)
    If ch = "!" Then ch = ReplaceWith   '  "!" has special meaning to LIKE
    If ch Like UnwantedChars Then
        ch = ReplaceWith
        If Right$(TmpStr, 1) = ReplaceWith Then ch = ""
    End If
    
    TmpStr = TmpStr + ch
Next i
ReplaceChars = TmpStr

End Function

Function CenterFormUp(frmForm As Form)  'as void
frmForm.Left = (Screen.Width - frmForm.Width) / 2
frmForm.Top = (Screen.Height - frmForm.Height) / 3
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
    tmp = tmp & "\" 'COOL?
    ExtractDirName = tmp
    
End Function
Function ExtractFileName(FileName As String) As String
    
'Extract the File title from a full file name


    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
    ExtractFileName = FileName
    Exit Function
    End If
    
    Do While pos <> 0
    PrevPos = pos
    pos = InStr(pos + 1, FileName, "\")
    Loop

    ExtractFileName = Right(FileName, Len(FileName) - PrevPos)

End Function



Function RevRGB(HexRGB As String) As String
Dim Var1 As String
Dim Var2 As String
Dim Var3 As String

Var1 = Left(HexRGB, 2)
Var2 = Mid(HexRGB, 3, 2)
Var3 = Right(HexRGB, 2)

RevRGB = Var3 & Var2 & Var1

End Function


Function SaveFile(FileName As String, FileContent As String) As Boolean
On Error GoTo Save_Error
Dim FileNum As Integer

FileNum = FreeFile

Open FileName For Output As #FileNum

Print #FileNum, FileContent

Close FileNum
SaveFile = True
Exit Function

Save_Error:
SaveFile = False
Exit Function
End Function


Function LoadFile(FileName As String) As String
'Loads the contents of a file into a string variable

On Error GoTo LoadFile_Error
Dim FF As Integer
Dim FileContents As String

FF = FreeFile
Open FileName For Input As #FF
FileContents = Input(LOF(FF), FF)
Close #FF
LoadFile = FileContents
Exit Function
LoadFile_Error:
    LoadFile = "#ERROR#"
    Exit Function

End Function

Public Function AddBR(ByVal Text As String) As String

Dim sTemp As String
Dim idx As Long
ReDim blines(1 To 1) As String




Text = Replace(Text, "&", "&amp;")
Text = Replace(Text, Chr$(34), "&quot;")
Text = Replace(Text, "<", "&lt;")
Text = Replace(Text, ">", "&gt;")

Text = Replace$(Text, vbCrLf, "<BR>" & vbCrLf)

AddBR = Text

End Function
Function SelectAllText(txtBox As TextBox) 'as void

txtBox.SelStart = 0
txtBox.SelLength = Len(txtBox.Text)

End Function

Sub Text2Lines(Text As String, Lines() As String)
Dim ch As String * 1
Dim cntr As Long
Dim index As Integer
Dim MaxIndex As Integer
Dim NewLine As String * 2

NewLine = Chr(13) + Chr(10)

ReDim Lines(1 To 9000)

index = 1
For cntr = 1 To Len(Text)
    ch = Mid$(Text, cntr, 1)
    Select Case Asc(ch)
        Case 13
            'do nothing
        Case 10     'always after the 13
            index = index + 1
        Case Else

            Lines(index) = Lines(index) + ch
    End Select
Next cntr

MaxIndex = index

ReDim Preserve Lines(1 To MaxIndex)
End Sub

Function ChangeFileExtension(FileName As String, NewExtension As String) As String
Dim OldExt As String
OldExt = ExtractFileExtension(FileName)
ChangeFileExtension = Left$(FileName, Len(FileName) - Len(OldExt)) & NewExtension

End Function

Function ExtractFileExtension(FileName As String) As String

    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, ".")
    If pos = 0 Then
    ExtractFileExtension = ""
    Exit Function
    End If
    
    Do While pos <> 0
    PrevPos = pos
    pos = InStr(pos + 1, FileName, ".")
    Loop

    ExtractFileExtension = Right(FileName, Len(FileName) - PrevPos)

End Function


Function Quote(sText As String) As String
Quote = """ & sText & """
End Function

Function Slasher(Strng As String, flag As String) As String
' Flag could be:
' "\?" to add a slash to the left if it doesn't already exist
' "?\" to add a slash to the right if it doesn't already exist
' "\?\" to enclose the string in slashes
' any other string to strip left and right slashes
' "?" can be any single character.

Dim AString As String

AString = Strng
If flag Like "\?" Then
    'left slash
    If Left(AString, 1) <> "\" Then AString = "\" + AString
ElseIf flag Like "?\" Then
    'right slash
    If Right(AString, 1) <> "\" Then AString = AString + "\"
ElseIf flag Like "\?\" Then
    'right & left slashes
    If Left(AString, 1) <> "\" Then AString = "\" + AString
    If Right(AString, 1) <> "\" Then AString = AString + "\"
Else
    'strip slashes if existing
    If Left(AString, 1) = "\" Then AString = Right(AString, Len(AString) - 1)
    If Right(AString, 1) = "\" Then AString = Left(AString, Len(AString) - 1)
End If

Slasher = AString
End Function


Function Read_ID3Tag(ByVal sFileName As String) As Variant
'On Error GoTo ERROR_UNKNOWN
Dim iFileNum As Integer
Dim Return_Array(1 To 3) As String
Dim sTag As String * 127
Dim Title$, Artist$, Album$
Dim fl As Long
iFileNum = FreeFile

Open sFileName For Binary Access Read Shared As #iFileNum

fl = LOF(1)
Seek #1, (fl - 128) + 1  'VB OFFSET!!!
Get #1, , sTag
Close #1

If Left(sTag, 3) <> "TAG" Then
    Read_ID3Tag = ""
Else
    Title = Trim(Mid(sTag, 4, 30))
    Artist = Trim(Mid(sTag, 34, 30))
    Album = Trim(Mid(sTag, 64, 30))
    Return_Array(1) = RTrimNulls(Title)
    Return_Array(2) = RTrimNulls(Artist)
    Return_Array(3) = RTrimNulls(Album)
    Read_ID3Tag = Return_Array
End If

End Function
Function RTrimNulls(sStr As String) As String
Dim idx As Integer
Dim ch As String * 1
For idx = Len(sStr) To 1 Step -1
    If Asc(Mid(sStr, idx, 1)) <> 0 Then Exit For
Next idx

RTrimNulls = Left(sStr, idx)
End Function
