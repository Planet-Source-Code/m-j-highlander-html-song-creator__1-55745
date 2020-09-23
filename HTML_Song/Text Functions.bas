Attribute VB_Name = "Text_Functions"
Option Explicit

Public Enum TextArrayOps
    BREACK_ONLY_AFTER_DOTS
End Enum

Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long


'Private ALL_NON_ALPHANUMERIC_CHARS As String
'Private ALL_ALPHANUMERIC_CHARS As String

Public Const Quote = """"
Public Const ALPHANUMERIC_CHARS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Public Const ALL_PRINTABLE_CHARS = ALPHANUMERIC_CHARS & Quote & " !#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
Public Const ALL_TEXT_CHARS = vbTab & vbCrLf & ALL_PRINTABLE_CHARS

Public Enum CharRangeConstants
    [AlphaNumeric Only] = 1  ' Starting from 1 , since VB defaults vars
    [All Printable] = 2      ' to zero, which might cause problems
    [All Text Chars] = 3
End Enum

Public Function RemoveNonAlphaNum4(ByVal Words As String, ByVal CharsToKeep As CharRangeConstants) As String
Dim ItIsValid  As Boolean
Dim Alpha As String
Dim iPos As Long
Dim sChar As String, sWork As String
Dim cntr As Long, idx As Long
Dim chars() As Byte
Dim b() As Byte
Dim c() As Byte

Select Case CharsToKeep
    Case [AlphaNumeric Only]
        Alpha = ALPHANUMERIC_CHARS
    Case [All Printable]
        Alpha = ALL_PRINTABLE_CHARS
    Case [All Text Chars]
        Alpha = ALL_TEXT_CHARS
End Select


ReDim b(0 To Len(Words) - 1)
ReDim c(0 To Len(Words) - 1)
b = StrConv(Words, vbFromUnicode) ' VB Strings are Double-Byte Unicode
idx = 0
For cntr = 0 To Len(Words) - 1
    ItIsValid = False
    If InStr(Alpha, Chr$(b(cntr))) Then
        c(idx) = b(cntr)
        'If idx > 0 Then If c(idx) = 10 And c(idx - 1) <> 13 Then c(idx) = 32
        idx = idx + 1
    End If
Next cntr

RemoveNonAlphaNum4 = FixNewLineChars(Left$(StrConv(c, vbUnicode), idx))


End Function

Public Function HandleTextTrigger(ByVal Text As String, TriggerChars As String) As String

Dim sTemp As String
Dim sResult As String
Dim sTrig As String
Dim sCaption As String

sTemp = Text

Do
    sTrig = FindTrigger(sTemp, TriggerChars)
    If sTrig = "" Then Exit Do
    sCaption = RemoveTriggerChars(sTrig, TriggerChars)
    sResult = InputBox(sCaption, "Enter Data", sCaption)
    sTemp = Replace(sTemp, sTrig, sResult)
Loop

HandleTextTrigger = sTemp

End Function

Private Function FindTrigger(sInputStr As String, sTriggerChar As String) As String

Dim lPos1 As Long
Dim lPos2 As Long

lPos1 = InStr(sInputStr, sTriggerChar)
lPos2 = InStr(lPos1 + 1, sInputStr, sTriggerChar)

If (lPos1 <> 0) And lPos2 <> 0 Then
    FindTrigger = Mid$(sInputStr, lPos1, lPos2 - lPos1 + Len(sTriggerChar))
Else
    FindTrigger = ""
End If

End Function


Private Function RemoveTriggerChars(InputStr As String, CharsToRemove As String) As String
    RemoveTriggerChars = Replace$(InputStr, CharsToRemove, "")
End Function



Public Function CompactBlankLines(ByVal Text As String) As String
' Convert successive blank lines into one line,
' also trims leading and trailing CR,LF. ?

Const DOUBLE_CRLF = vbCrLf & vbCrLf
Const TRIPPLE_CRLF = vbCrLf & vbCrLf & vbCrLf

Dim pos As Long
Dim sWork As String
Dim idx As Long

sWork = Trim$(Text)

'Keep Removing double CRLF's until none found:
pos = InStr(sWork, TRIPPLE_CRLF)
Do While pos > 0
     sWork = Replace$(sWork, TRIPPLE_CRLF, DOUBLE_CRLF)
     pos = InStr(sWork, TRIPPLE_CRLF)
Loop

'Trim right and left of Text
sWork = CrLfTabTrim(sWork) 'could be slow?

CompactBlankLines = sWork & vbCrLf  'it surely has none

End Function

Public Property Get CharAt(Text As String, Position As Long) As String

    If Position > 0 Then
        CharAt = Mid$(Text, Position, 1)
    End If

End Property

Public Property Let CharAt(Text As String, Position As Long, ByVal sNewValue As String)
    
    If Position > 0 Then
        Mid$(Text, Position, 1) = sNewValue
    End If
    
End Property

Public Function CrLfTabTrim(ByVal Text As String) As String
'Trim leading and trailing Cr , Lf , Tab and Space Chars

Dim sTemp As String
Dim ch As String * 1
Dim idx As Long

sTemp = Trim$(Text)

'''Left Trim (Convert Cr,Lf and Tab to Spaces)
For idx = 1 To Len(sTemp)
    ch = CharAt(sTemp, idx)
    If (ch = vbCr Or ch = vbLf Or ch = vbTab) Then
        ch = " "
        CharAt(sTemp, idx) = ch
    Else
        Exit For 'break at first non Cr/Lf/Tab char
    End If
Next idx

'''Right Trim (Convert Cr,Lf and Tab to Spaces)
For idx = Len(sTemp) To 1 Step -1
    ch = CharAt(sTemp, idx)
    If (ch = vbCr Or ch = vbLf Or ch = vbTab) Then
        ch = " "
        CharAt(sTemp, idx) = ch
    Else
        Exit For 'break at first non-Cr/Lf/Tab char
    End If
Next idx

'''Trim Spaces
CrLfTabTrim = Trim$(sTemp)

End Function

Public Function DoEllipses(ByVal Text As String, Optional MaxLength As Long = 0) As String

Dim sTemp As String
Dim pos As Long

'Remove Leading and Trailing Cr,Lf,Tab and Space
sTemp = CrLfTabTrim(Text)

'Ignore text after the first Cr char
pos = InStr(1, sTemp, vbCr)
If pos > 0 Then
    sTemp = Left$(sTemp, pos - 1)
End If

'Take only the specified length (if specified!)
If MaxLength > 0 Then
    sTemp = Left$(sTemp, MaxLength)
End If

'Did we truncate the text? if yes add Ellipses (...)
If Len(sTemp) < Len(Text) Then
    sTemp = sTemp & " ..."
End If

DoEllipses = sTemp

End Function

Public Function GetTextFileContents(ByVal Filename As String) As String
On Error GoTo Error_GetTextFileContents
Dim iFF As Integer

iFF = FreeFile

Open Filename For Input As #iFF
    GetTextFileContents = Input$(LOF(iFF), iFF)
Close #iFF

Exit Function
Error_GetTextFileContents:
    GetTextFileContents = ""
    
End Function
Public Function RemoveNonAlphaNum3(ByVal Words As String, ByVal CharsToKeep As CharRangeConstants) As String
' Remove all non-alphanumeric characters from the Words
' The fastest method.

Dim Alpha As String
Dim iPos As Long
Dim sChar As String, sWork As String
Dim cntr As Long, idx As Long
Dim ch As String * 1

Select Case CharsToKeep
    Case [AlphaNumeric Only]
        Alpha = ALPHANUMERIC_CHARS
    Case [All Printable]
        Alpha = ALL_PRINTABLE_CHARS
    Case [All Text Chars]
        Alpha = ALL_TEXT_CHARS
End Select

sWork = String(Len(Words), " ")

For cntr = 1 To Len(Words)
    ch = Mid$(Words, cntr, 1)
    If InStr(Alpha, ch) <> 0 Then
        idx = idx + 1
        Mid$(sWork, idx, 1) = ch
    End If
Next cntr

sWork = Left$(sWork, idx)

RemoveNonAlphaNum3 = FixNewLineChars(sWork)

End Function

Public Function RemoveNonAlphaNum2(ByVal Words As String) As String
' Remove all non-alphanumeric characters from the Words

Const Alpha = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz "
Const SNG_SPACE = " "
Const DBL_SPACE = "  "

Dim iPos As Long
Dim sChar As String, sWork As String
Dim lCntr As Long

'Set the working variable
sWork = Trim(Words)

'Remove all charaters that are NOT in the ALPHA const
For lCntr = 0 To 255
    If InStr(Alpha, Chr(lCntr)) = 0 Then
        sWork = Replace(sWork, Chr(lCntr), SNG_SPACE)
    End If
Next lCntr

'Remove all double spaces created
iPos = InStr(sWork, DBL_SPACE)
While iPos > 0
     sWork = Replace(sWork, DBL_SPACE, SNG_SPACE)
    iPos = InStr(sWork, DBL_SPACE)
Wend


RemoveNonAlphaNum2 = sWork
    
End Function

Public Function ReverseStr(ByVal Text As String, Optional ByLine As Boolean = True) As String
ReDim sArray(1 To 1) As String
Dim sTemp As String
Dim idx As Long
Dim cidx As Long

If ByLine = False Then  ' Reverse Entire Text
    sTemp = String$(Len(Text), " ")
    For idx = 0 To Len(Text) - 1
        ' Kinda Cool!
        CharAt(sTemp, idx + 1) = CharAt(Text, Len(Text) - idx)
    Next idx
    ' CrLf will also be reversed, so fix it
    sTemp = Replace$(sTemp, vbLf & vbCr, vbCrLf)

Else                    ' Reverse Each Line Alone
    Text2Array Text, sArray
    For idx = LBound(sArray) To UBound(sArray)
            sTemp = String$(Len(sArray(idx)), " ")
            For cidx = 0 To Len(sTemp) - 1
                ' Kinda Cool!
                CharAt(sTemp, cidx + 1) = CharAt(sArray(idx), Len(sTemp) - cidx)
            Next cidx
            sArray(idx) = sTemp
    Next idx
    sTemp = Join$(sArray, vbCrLf)
End If

ReverseStr = sTemp

End Function

Public Function SetTextMaxWidthWords(Text As String, iMax As Integer)

Dim idx As Long
ReDim sTempArray(1 To 1) As String

Text2Array Text, sTempArray


For idx = LBound(sTempArray) To UBound(sTempArray)
        sTempArray(idx) = SetLineMaxWidthWords(sTempArray(idx), iMax)
Next idx

SetTextMaxWidthWords = Join(sTempArray, vbCrLf)


End Function


Function SetLineMaxWidthWords(ByVal TextLine As String, ByVal MaxWidth As Long) As String

Dim HowManyTimes As Long
Dim idx As Long
Dim PosShift As Long
Dim sTemp As String
Dim sReturn As String

HowManyTimes = Len(TextLine) \ MaxWidth
If Len(TextLine) Mod MaxWidth <> 0 Then 'there is a remainder
    HowManyTimes = HowManyTimes + 1
End If

PosShift = MaxWidth

If MaxWidth < Len(TextLine) Then
    sTemp = TextLine
    For idx = 1 To HowManyTimes
        
        Do While PosShift > 0
            If IsChrAlphaNumeric(Mid$(sTemp, PosShift, 1)) Then
                PosShift = PosShift - 1
            Else
                Exit Do
            End If
        Loop
        
        sTemp = InsertString(sTemp, vbCrLf, PosShift)
        PosShift = PosShift + MaxWidth + 2
    Next idx
    SetLineMaxWidthWords = sTemp
Else
    SetLineMaxWidthWords = TextLine
End If

End Function


Public Sub BuildCharLists()
''Case 48 To 57, 65 To 90, 97 To 122   ' 0..9 , A..Z , a..z
'Dim ch As String * 1
'Dim idx As Integer
'
'ALL_ALPHANUMERIC_CHARS = ""
'For idx = 48 To 57
'    ALL_ALPHANUMERIC_CHARS = ALL_ALPHANUMERIC_CHARS & Chr(idx)
'Next idx
'For idx = 65 To 90
'    ALL_ALPHANUMERIC_CHARS = ALL_ALPHANUMERIC_CHARS & Chr(idx)
'Next idx
'For idx = 97 To 122
'    ALL_ALPHANUMERIC_CHARS = ALL_ALPHANUMERIC_CHARS & Chr(idx)
'Next idx
''''''''''''''''''''''''''''''''''''''
'ALL_NON_ALPHANUMERIC_CHARS = ""
'For idx = 0 To 47
'    ALL_NON_ALPHANUMERIC_CHARS = ALL_NON_ALPHANUMERIC_CHARS & Chr(idx)
'Next idx
'For idx = 58 To 64
'    ALL_NON_ALPHANUMERIC_CHARS = ALL_NON_ALPHANUMERIC_CHARS & Chr(idx)
'Next idx
'For idx = 91 To 96
'    ALL_NON_ALPHANUMERIC_CHARS = ALL_NON_ALPHANUMERIC_CHARS & Chr(idx)
'Next idx
'For idx = 123 To 255
'    ALL_NON_ALPHANUMERIC_CHARS = ALL_NON_ALPHANUMERIC_CHARS & Chr(idx)
'Next idx

End Sub



Public Function IsChrAlphaNumeric(ByVal Char As String) As Boolean

Dim cChar As Byte
If Char <> "" Then
    cChar = Asc(Char)
    IsChrAlphaNumeric = CBool(IsCharAlphaNumeric(cChar))
Else
    IsChrAlphaNumeric = False
End If

End Function


Function AddToLine(ByVal TextLine As String, ByVal AddToLeft As String, ByVal AddToRight As String) As String

AddToLine = AddToLeft & TextLine & AddToRight


End Function


Public Function AddLineNumbers(Text As String, ByVal NumStart As Long, ByVal NumStep As Long, ByVal Delimiter As String, NumDigits As Long, IgnoreEmptyLines As Boolean) As String
Dim FormatStr As String
Dim idx As Long
Dim cntr As Long
ReDim sTempArray(1 To 1) As String

FormatStr = String(NumDigits, "0")

Text2Array Text, sTempArray

cntr = NumStart
For idx = LBound(sTempArray) To UBound(sTempArray)
    If IgnoreEmptyLines And sTempArray(idx) = "" Then
        'do nothing
    Else
        sTempArray(idx) = Format(cntr, FormatStr) & Delimiter & sTempArray(idx)
        cntr = cntr + NumStep
    End If
    
Next idx

AddLineNumbers = Join(sTempArray, vbCrLf)

End Function



Function AddToLines(Text As String, AddToLeft As String, AddToRight As String, IgnoreEmptyLines As Boolean) As String
Dim idx As Long
ReDim TheArray(1 To 1) As String

Text2Array Text, TheArray

For idx = LBound(TheArray) To UBound(TheArray)
    If IgnoreEmptyLines And TheArray(idx) = "" Then
        'do nothing
    Else
         TheArray(idx) = AddToLine(TheArray(idx), AddToLeft, AddToRight)
    End If
Next idx

AddToLines = Array2Text(TheArray)

End Function

Function MultiInstrRev(Start As Long, Text As String, LookFor As String, Compare As VbCompareMethod)

Dim iLen As Long
Dim chLookFor As String * 1
Dim idx As Long
Dim iPos As Long
Dim iFirstPos As Long


iLen = Len(LookFor)
iFirstPos = 0

For idx = 1 To iLen
    chLookFor = Mid(LookFor, idx, 1)
    iPos = InStrRev(Text, chLookFor, Start, Compare)
    If (iPos <> 0 And iPos > iFirstPos) Then
         iFirstPos = iPos
'         MsgBox iFirstPos
    End If

Next idx


'If iFirstPos = Len(Text) + 1 Then ' value didn't change / nothing found
'    MultiInstrRev = 0
'    MsgBox "X"
'Else
    MultiInstrRev = iFirstPos
'End If

End Function


Function MultiInstr(Start As Long, Text As String, LookFor As String, Compare As VbCompareMethod)

Dim iLen As Long
Dim chLookFor As String * 1
Dim idx As Long
Dim iPos As Long
Dim iFirstPos As Long


iLen = Len(LookFor)
iFirstPos = Len(Text) + 1 ' out of text boundry // an impossible value

For idx = 1 To iLen
    chLookFor = Mid(LookFor, idx, 1)
    iPos = InStr(Start, Text, chLookFor, Compare)
    If (iPos <> 0 And iPos < iFirstPos) Then
         iFirstPos = iPos
    End If
    
Next idx

If iFirstPos = Len(Text) + 1 Then ' value didn't change / nothing found
    MultiInstr = 0
Else
    MultiInstr = iFirstPos
End If

End Function

Function InsertString(MainStr As String, SubStr As String, Position As Long)
Dim sLeft As String, sRight As String

Select Case Position
    Case Is > Len(MainStr)
        sLeft = MainStr
        sRight = ""
    Case Is <= 0
        sLeft = ""
        sRight = MainStr
    Case Is <= Len(MainStr)
        sLeft = Left(MainStr, Position)
        sRight = Right(MainStr, Len(MainStr) - Position)
End Select
    
InsertString = sLeft & SubStr & sRight

End Function

Function xDeSpace(sText As String) As String
Dim idx As Integer
Dim sTemp As String
Dim sResult As String
Dim ch As String * 1
Dim InSpaces As Boolean
Dim scount As Integer

sTemp = Trim$(sText)
sResult = ""

scount = 0
For idx = 1 To Len(sTemp)
    ch = Mid(sTemp, idx, 1)
        
    Select Case ch
        Case " "
        scount = scount + 1
        Case Else
        InSpaces = False
        scount = 0
    End Select
    
    
    If scount > 1 Then
        InSpaces = True
    End If
    
    If Not (InSpaces) Then
        sResult = sResult & ch
    End If
    
Next idx
xDeSpace = sResult
End Function



Public Function SetTextMaxWidth(Text As String, iMax As Integer)

Dim idx As Long
ReDim sTempArray(1 To 1) As String

Text2Array Text, sTempArray


For idx = LBound(sTempArray) To UBound(sTempArray)
        sTempArray(idx) = SetLineMaxWidth(sTempArray(idx), iMax)
Next idx

SetTextMaxWidth = Join(sTempArray, vbCrLf)


End Function

Function RemoveNonAlphaNum1(Text) As String

Const Pattern = "[a-zA-Z0-9" & vbCrLf & ".,;: ]"

Dim idx As Long
Dim sTemp As String
Dim Char As String * 1


sTemp = ""
    For idx = 1 To Len(Text)
        Char = Mid(Text, idx, 1)
        If Char Like Pattern Then sTemp = sTemp & Char
    Next

RemoveNonAlphaNum1 = sTemp

'0 64,91 96 , 123 255


End Function


Public Function Max(ByVal ValA, ByVal ValB)

Max = IIf(ValA > ValB, ValA, ValB)

End Function


Public Function DelLeftTo(sMainStr As String, sSubStr As String, boolMatchCase As Boolean, boolInclusive As Boolean) As String

Dim pos As Long
If boolMatchCase = True Then
        pos = InStr(1, sMainStr, sSubStr, vbBinaryCompare)
Else
        pos = InStr(1, sMainStr, sSubStr, vbTextCompare)
End If


If (pos = 0) Then
            DelLeftTo = sMainStr
Else
            If boolInclusive = True Then
                        DelLeftTo = Right(sMainStr, Len(sMainStr) - pos - Len(sSubStr) + 1)
            Else
                        DelLeftTo = Right(sMainStr, Len(sMainStr) - pos + 1)
            End If
End If

End Function

Public Function DelRightTo(sMainStr As String, sSubStr As String, boolMatchCase As Boolean, boolInclusive As Boolean) As String

Dim pos As Long
If boolMatchCase = True Then
        pos = InStrRev(sMainStr, sSubStr, -1, vbBinaryCompare)
Else
        pos = InStrRev(sMainStr, sSubStr, -1, vbTextCompare)
End If


If (pos = 0) Then
            DelRightTo = sMainStr
Else
            If boolInclusive = True Then
                        DelRightTo = Left(sMainStr, pos - 1)
            Else
                        DelRightTo = Left(sMainStr, pos + Len(sSubStr) - 1)
            End If
End If

End Function

Function Array2Text(sArray() As String) As String

Array2Text = Join(sArray, vbCrLf)

End Function

Public Function DelLeft(ByVal TextLine As String, Count As Long) As String

If Count <= 0 Then
    DelLeft = TextLine
ElseIf Count > Len(TextLine) Then
    DelLeft = ""
Else
    DelLeft = Right$(TextLine, Len(TextLine) - Count)
End If

End Function

Function DelRight(ByVal TextLine As String, Count As Long) As String

If Count <= 0 Then
    DelRight = TextLine
ElseIf Count > Len(TextLine) Then
    DelRight = ""
Else
    DelRight = Left$(TextLine, Len(TextLine) - Count)
End If

End Function

Function FixNewLineChars(ByVal Text As String) As String

Dim LineFeed As String * 1
Dim CarrigeReturn As String * 1
Dim BeepChar As String * 1

LineFeed = Chr(10)
CarrigeReturn = Chr(13)
BeepChar = Chr(7)
Text = Replace(Text, vbCrLf, BeepChar)
Text = Replace(Text, LineFeed, BeepChar)
Text = Replace(Text, CarrigeReturn, BeepChar)
Text = Replace(Text, LineFeed & CarrigeReturn, BeepChar)

FixNewLineChars = Replace(Text, BeepChar, vbCrLf)
End Function


Function InsertString2(ByVal TextLine As String, ByVal StrToInsert As String, ByVal Position As Integer) As String
Dim sLeft As String, sRight As String

If Position > Len(TextLine) Then
    Position = Len(TextLine)
End If

If TextLine = "" Then
    InsertString2 = ""
Else
    
    sLeft = Left(TextLine, Position)
    sRight = Right(TextLine, Len(TextLine) - Position)
    
    InsertString2 = sLeft & StrToInsert & sRight

End If

End Function

Function RemoveBlankLines(ByVal Text As String) As String
Dim idx As Long
Dim sTemp As String
ReDim sArray(1 To 1) As String

Text2Array Text, sArray

For idx = LBound(sArray) To UBound(sArray)
    If sArray(idx) = "" Then
            sArray(idx) = Chr$(7)
    End If
Next idx

sTemp = Join$(sArray, vbCrLf)
sTemp = Replace$(sTemp, Chr$(7) & vbCrLf, "")

RemoveBlankLines = Replace$(sTemp, Chr$(7), "") 'Trim trailing if exists.

End Function

Function SetLineMaxWidth(ByVal TextLine As String, ByVal MaxWidth As Integer) As String

Dim HowManyTimes As Integer
Dim idx As Integer
Dim PosShift As Long
Dim sTemp As String
Dim sReturn As String

HowManyTimes = Len(TextLine) \ MaxWidth
PosShift = MaxWidth

If MaxWidth < Len(TextLine) Then
    sTemp = TextLine
    For idx = 1 To HowManyTimes
        sTemp = InsertString(sTemp, vbCrLf, PosShift)
'        If sReturn <> "" Then sTemp = sReturn
        PosShift = PosShift + MaxWidth + 2
    Next idx
    SetLineMaxWidth = sTemp
Else
    SetLineMaxWidth = TextLine
End If

End Function

Public Function CompactSpaces(ByVal Text As String) As String
' Convert successive spaces into one space,
' also trims leading and trailing spaces. / Really Fast!

Const SNG_SPACE = " "
Const DBL_SPACE = "  "

Dim pos As Long
Dim sWork As String
Dim idx As Long

sWork = Trim$(Text)

'Keep Removing double spaces until none found:
pos = InStr(sWork, DBL_SPACE)
Do While pos > 0
     sWork = Replace$(sWork, DBL_SPACE, SNG_SPACE)
     pos = InStr(sWork, DBL_SPACE)
Loop

'Trim right and left of each line, don't trim tabs:
sWork = TrimSpaces(sWork, True, True, False)

CompactSpaces = sWork

End Function

Function stringf(Text As String) As String
Dim sTemp As String

sTemp = Replace(Text, "\n", vbCrLf, , , vbTextCompare)
sTemp = Replace(sTemp, "\t", vbTab, , , vbTextCompare)
sTemp = Replace(sTemp, "\\", "\", , , vbTextCompare)

stringf = sTemp

End Function

Function Tab2Spaces(Text As String, NumSpaces As Integer) As String

Dim sTemp As String

If NumSpaces < 1 Then NumSpaces = 1
sTemp = Space$(NumSpaces)

Tab2Spaces = Replace(Text, Chr(9), sTemp)


End Function

Function BreackOnlyAfter(ByVal Text As String, ByVal AfterWhat As String) As String

Dim sTemp As String


sTemp = Replace(Text, vbCrLf, "")
sTemp = Replace(sTemp, ".", "." & vbCrLf)

BreackOnlyAfter = sTemp


End Function


Function TrimSpaces(Text As String, TrimLeft As Boolean, TrimRight As Boolean, TrimTab As Boolean) As String

Dim idx As Long
ReDim sTempArray(1 To 1) As String

Text2Array Text, sTempArray

For idx = LBound(sTempArray) To UBound(sTempArray)
    If (TrimLeft And TrimRight) Then
        sTempArray(idx) = Trim(sTempArray(idx))
        If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), True, True)
     ElseIf TrimLeft Then
        sTempArray(idx) = LTrim(sTempArray(idx))
        If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), True, False)
    ElseIf TrimRight Then
        sTempArray(idx) = RTrim(sTempArray(idx))
       If TrimTab Then sTempArray(idx) = TrimTabs(sTempArray(idx), False, True)
    Else
        'Do Nothing
    End If
Next idx

TrimSpaces = Join(sTempArray, vbCrLf)

End Function


Sub Text2Array(ByVal sText As String, ByRef sArray() As String)
    ' sText should not be Empty:
    ' Check for it in the calling routine.
    
    Dim vTmpArray As Variant
    Dim idx As Long
    
    vTmpArray = Split(sText, vbCrLf)
    ReDim sArray(LBound(vTmpArray) To UBound(vTmpArray))
    
    For idx = LBound(vTmpArray) To UBound(vTmpArray)
        sArray(idx) = vTmpArray(idx)
    Next idx

End Sub
Function TrimTabs(ByVal TextLine As String, TrimLeft As Boolean, TrimRight As Boolean) As String
Dim idx As Long
Dim ch As String * 1

If (TrimLeft And TrimRight) Then
    'Trim both
    For idx = 1 To Len(TextLine)
        ch = Mid(TextLine, idx, 1)
        If ch = Chr(9) Then 'TAB
               Mid(TextLine, idx) = Chr(7)
        Else
                Exit For
        End If
    Next idx
    For idx = Len(TextLine) To 1 Step -1
        ch = Mid(TextLine, idx, 1)
        If ch = Chr(9) Then 'TAB
               Mid(TextLine, idx) = Chr(7)
        Else
                Exit For
        End If
    Next idx

ElseIf TrimLeft Then
    For idx = 1 To Len(TextLine)
        ch = Mid(TextLine, idx, 1)
        If ch = Chr(9) Then 'TAB
               Mid(TextLine, idx) = Chr(7)
        Else
                Exit For
        End If
    Next idx

ElseIf TrimRight Then
    For idx = Len(TextLine) To 1 Step -1
        ch = Mid(TextLine, idx, 1)
        If ch = Chr(9) Then 'TAB
               Mid(TextLine, idx) = Chr(7)
        Else
                Exit For
        End If
    Next idx

End If

TrimTabs = Replace(TextLine, Chr(7), "")

End Function
