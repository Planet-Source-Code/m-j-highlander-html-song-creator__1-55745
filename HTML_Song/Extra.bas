Attribute VB_Name = "ExtraFunctions"
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
    GetAppPath = App.Path & "\"
End If

End Function

Function RemoveSlash(ByVal sPath As String) As String

sPath = Trim(sPath)

If Right(sPath, 1) = "\" Then
    RemoveSlash = Left(sPath, Len(sPath) - 1)
Else
    RemoveSlash = sPath
End If


End Function

Function CurDrive$()
    CurDrive$ = Left$(CurDir$, 1)
End Function

Function DoHex(txt As String) As String
Dim ch As String
Dim DH As String
Dim i As Integer
Dim dch As Integer

DH = ""
txt = Trim(txt)

For i = 1 To Len(txt) Step 2
ch = Mid(txt, i, 2)
If Left(ch, 1) <> " " Then
    dch = Hex2Dec(ch)
    DH = DH + Chr(dch)
End If
Next i
DoHex = DH
End Function

Sub Encrypt(Secret$, PassWord$)

' Secret$ = the string you wish to encrypt or decrypt.
' PassWord$ = the password with which to encrypt the string.
' Calling the sub on a string Encrypts it
' Calling the sub on an encrypted  string Decrypts it

Dim l As Long
Dim X As Long
Dim iChar As Integer

    l = Len(PassWord$)
    For X = 1 To Len(Secret$)
        iChar = Asc(Mid$(PassWord$, (X Mod l) - l * ((X Mod l) = 0), 1))
        Mid$(Secret$, X, 1) = Chr$(Asc(Mid$(Secret$, X, 1)) Xor iChar)
    Next X

End Sub

Function ExtractDirNameX(FileName As String) As String

'Extract the Directory name from a full file name
'The return will have "\" appended to the end

    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractDirNameX = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    ExtractDirNameX = Left(FileName, PrevPos)

End Function

Function GetParentDir(sPath As String)
Dim pos%, OldPos%

pos = 0

Do
   OldPos = pos
   pos = InStr(pos + 1, sPath, "\")
   If pos = 0 Then Exit Do
Loop

GetParentDir = Left(sPath, OldPos - 1)

End Function

Function GetTempDir()

    GetTempDir = Environ$("temp")

End Function

Function Hex2Dec(ch As String)
Dim l As String, r As String
Dim rNum As Integer, lNum As Integer

ch = UCase(ch)
l = Left(ch, 1)
r = Right(ch, 1)
Select Case r
    Case "0" To "9"
    rNum = Val(r)
    Case "A" To "F"
    rNum = Asc(r) - 55
End Select

Select Case l
    Case "0" To "9"
    lNum = Val(l)
    Case "A" To "F"
    lNum = Asc(l) - 55
End Select

Hex2Dec = rNum + 16 * lNum

End Function

Function Hex2Int(ByVal HexNum As String) As Integer
Dim ch As String * 1
Dim d As Integer, dd As Integer

HexNum = UCase$(HexNum)
ch = Right$(HexNum, 1)
Select Case ch
    Case "0" To "9"
        d = Val(ch)
    Case "A"
        d = 10
    Case "B"
        d = 11
    Case "C"
        d = 12
    Case "D"
        d = 13
    Case "E"
        d = 14
    Case "F"
        d = 15
End Select

ch = Left$(HexNum, 1)
Select Case ch
    Case "0" To "9"
        dd = Val(ch)
    Case "A"
        dd = 10
    Case "B"
        dd = 11
    Case "C"
        dd = 12
    Case "D"
        dd = 13
    Case "E"
        dd = 14
    Case "F"
        dd = 15
End Select

Hex2Int = d + 16 * dd
End Function

Function iIsValidFileName(FName As String) As Integer
' Check if FName is a valid file name
' Returns True Or False

On Error Resume Next

Const ATTR_NORMAL = 0
Const ATTR_HIDDEN = 2
Const ATTR_SYSTEM = 4

Dim FF As Integer
Dim Exists As Integer
Dim Attr_Mask As Integer

Attr_Mask = ATTR_NORMAL + ATTR_HIDDEN + ATTR_SYSTEM
If Dir$(FName, Attr_Mask) <> "" Then
    Exists = True
Else
    Exists = False
End If
FF = FreeFile
Open FName For Binary As #FF
If Err Then
    'not a valid name
    iIsValidFileName = False
    Exit Function
Else
    iIsValidFileName = True
    Close #FF
    If Exists = False Then Kill FName
End If

End Function

Function InRange(iValue As Integer, LowBound As Integer, UpperBound As Integer) As Integer

Select Case iValue
    Case Is < LowBound
        InRange = LowBound
    Case Is > UpperBound
        InRange = UpperBound
    Case Else
        InRange = iValue
End Select

End Function

Function PseudoCrypt(NumWords As Integer) As String
' Generate random text that appears as if it is
' an encrpted text!
' It is useful in creating a piece of text for testing purposes.

Dim i As Integer
Dim lngth As Integer

Dim tmp As String
Randomize

For i = 1 To NumWords
    lngth = Rnd * 8
    tmp = tmp & " " & RndStr(lngth)
Next i

PseudoCrypt = Trim(tmp)

End Function

Function ReadFile(sFileName As String) As String
On Error GoTo ReadFileError
Dim FF As Integer
Dim TmpStr As String

FF = FreeFile
Open sFileName For Input As #FF
TmpStr = Input$(LOF(FF), FF)
Close #FF
ReadFile = TmpStr
Exit Function

ReadFileError:
    ReadFile = ""
    Exit Function

End Function

Sub RGBSplit(RGB_Color As Long, r%, G%, b%)
Dim HexRGB As String
Dim HexR$, HexG$, HexB$

HexRGB = Hex$(RGB_Color)
If Len(HexRGB) < 6 Then HexRGB = String(6 - Len(HexRGB), "0") + HexRGB

HexR = Right(HexRGB, 2)
HexG = Mid(HexRGB, 3, 2)
HexB = Left(HexRGB, 2)
r% = Hex2Int(HexR)
G% = Hex2Int(HexG)
b% = Hex2Int(HexB)

End Sub

Function RndInt(lower, upper) As Integer
'Returns a random integer greater than or equal to the Lower parameter
'and less than or equal to the Upper parameter.
Randomize Timer
RndInt = Int(Rnd * (upper - lower + 1)) + lower

End Function

Function RndStr(StrLen As Integer) As String
' generate a random string containing small case letters only

Dim idx As Integer
Dim ch As String * 1
Dim tmp As String

For idx = 1 To StrLen
        ch = Chr$(RndInt(97, 122))
        tmp = tmp + ch
Next idx

RndStr = tmp

End Function

Sub ShellSort(SortArray() As String)
'The fastets sort algorithm!

Dim Row As Integer
Dim MaxRow As Integer
Dim MinRow As Integer
Dim swtch As Integer
Dim Limit As Integer
Dim Offset As Integer

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
            If LCase(SortArray(Row)) > LCase(SortArray(Row + Offset)) Then
               Swap SortArray(Row), SortArray(Row + Offset)
               swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = swtch - Offset
      Loop While swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub

Function sIsValidFileName(FName As String) As String
' Check if FName is a valid file name
' Returns the Actual file (might differ from FName) name or an empty string

On Error Resume Next

'Const ATTR_NORMAL = 0
'Const ATTR_HIDDEN = 2
'Const ATTR_SYSTEM = 4

Dim FF As Integer
Dim Exists As Integer
Dim Attr_Mask As Integer

Attr_Mask = ATTR_NORMAL + ATTR_HIDDEN + ATTR_SYSTEM
If Dir$(FName, Attr_Mask) <> "" Then
    Exists = True
Else
    Exists = False
End If
FF = FreeFile
Open FName For Binary As #FF
If Err Then
    'not a valid name
    sIsValidFileName = ""
    Exit Function
Else
    sIsValidFileName = Dir(FName)
    Close #FF
    If Exists = False Then Kill FName
End If

End Function

Function Slash(Strng As String) As String

If Len(Strng) = 0 Then
    Slash = ""
    Exit Function
End If

If Right$(Strng, 1) <> "\" Then
    Slash = Strng + "\"
Else
    Slash = Strng
End If

End Function

Sub SortLines(Text As String)
Dim ch As String * 1
Dim cntr As Long
Dim index As Integer
Dim MaxIndex As Integer
Dim NewLine As String * 2

'Text = XTrim(Text)
NewLine = Chr(13) + Chr(10)

ReDim Lines(1 To 1000) As String

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

'For Cntr = 1 To MaxIndex
'    Lines(Cntr) = XTrim(Lines(Cntr))
'Next Cntr


ShellSort Lines()
Text = ""
For cntr = 1 To MaxIndex
    Lines(cntr) = XTrim(Lines(cntr))
    If Lines(cntr) <> "" Then
        Text = Text + Lines(cntr) + NewLine
    End If

Next cntr

End Sub

Sub SplitFileName(FName As String, TheName As String, TheExt As String)
Dim cntr As Integer
Dim ch As String * 1
Dim ThePos As Integer
ThePos = 0
'In 32 bit Windows, FName may contain more than one "." we want the last one
For cntr = Len(FName) To 1 Step -1
    ch = Mid$(FName, cntr, 1)
    If ch = "." Then
        ThePos = cntr
        Exit For
    End If
Next cntr


If ThePos = 0 Then
    TheName = FName
    TheExt = ""
Else
    TheName = Left$(FName, ThePos - 1)
    TheExt = Right(FName, Len(FName) - ThePos)
End If

End Sub

Function StringToColor(sColor) As Long
Const BLACK = &H0&
Const RED = &HFF&
Const GREEN = &HFF00&
Const YELLOW = &HFFFF&
Const BLUE = &HFF0000
Const MAGENTA = &HFF00FF
Const CYAN = &HFFFF00
Const WHITE = &HFFFFFF
Const GREY = &HC0C0C0

    Select Case LCase(sColor)
        Case "black", "k"
            StringToColor = BLACK
        Case "red", "r"
            StringToColor = RED
        Case "green", "g"
            StringToColor = GREEN
        Case "yellow", "y"
            StringToColor = YELLOW
        Case "blue", "b"
            StringToColor = BLUE
        Case "magenta", "m"
            StringToColor = MAGENTA
        Case "cyan", "c"
            StringToColor = CYAN
        Case "white", "w"
            StringToColor = WHITE
        Case "grey", "e"
            StringToColor = GREY
    Case Else
            StringToColor = BLACK
    End Select
End Function

Sub Swap(Var1, Var2)
Dim tmp As Variant
    tmp = Var1
    Var1 = Var2
    Var2 = tmp

End Sub

Function UpEachFirst(strz As String) As String

Dim OutStr As String
Dim Char As String * 1
Dim ch As String * 1
Dim i%

OutStr = UCase(Left(strz, 1))

For i = 1 To Len(strz) - 1
    ch = Mid(strz, i, 1)
        If ch = " " Then
             Char = UCase(Mid(strz, i + 1, 1))
        Else
             Char = LCase(Mid(strz, i + 1, 1))
        End If
    
    OutStr = OutStr + Char

Next i
UpEachFirst = OutStr
End Function

Function UpFirst(strz As String) As String
Dim FirstLetter$, OtherLetters$

FirstLetter = UCase$(Left(strz, 1))
OtherLetters = LCase(Right(strz, Len(strz) - 1))
UpFirst = FirstLetter + OtherLetters

End Function

Function WriteFile(sFileName As String, sContents As String) As Boolean

Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_DIRECTORY
Dim FF As Integer

On Error GoTo WriteFileError

FF = FreeFile
If Dir(sFileName, ATTR_ALL_FILES) = "" Then
    Open sFileName For Output As #FF
    Print #FF, sContents
    Close #FF
    WriteFile = True
Else       'A file or a folder with the same name already exists
    WriteFile = False
End If

Exit Function

WriteFileError:
    WriteFile = False
    Exit Function

End Function

Function XTrim(sLine As String) As String
'//Is this func OK????????/

Dim ch As String * 1
sLine = Trim$(sLine)
If Right(sLine, 1) = Chr$(13) Or Right(sLine, 1) = Chr$(13) Then
    sLine = Left(sLine, Len(sLine) - 1)
End If
XTrim = sLine
End Function

