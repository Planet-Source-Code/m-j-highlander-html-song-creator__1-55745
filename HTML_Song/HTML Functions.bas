Attribute VB_Name = "HTML_Functions"
Option Explicit

Public Enum HTML_Operation
    Cancel
    DeleteTagAndContent
    DeleteTagKeepContent
    ExtractTagAndContent
End Enum


Type sTag
    Href As String
    Text As String
End Type


Public Function CommentOut_ImageTags(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<img", vbTextCompare)
'MsgBox iOpeningPos

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, ">", vbTextCompare)
        
'MsgBox iClosingPos
    'sTempChars = String(iClosingPos - iOpeningPos + 1, Chr$(7))
    Mid$(sHTML, iOpeningPos) = "<!--"
    Mid$(sHTML, iClosingPos - 2) = "-->"
Loop

CommentOut_ImageTags = sHTML

End Function

Public Function IsURLLocal(ByVal Text As String) As Boolean
Dim pos As Long

pos = 0
pos = InStr(1, Text, "http://", vbTextCompare)
pos = pos + InStr(1, Text, "ftp://", vbTextCompare)
pos = pos + InStr(1, Text, "www.", vbTextCompare)
' if after this pos is still 0 , then none was found

If pos = 0 Then
    IsURLLocal = True
Else
    IsURLLocal = False
End If

'MsgBox Text & " " & IsURLLocal

End Function

Public Function RemovePath(ByVal Attr As String) As String
' INPUT:  Attribute with contents
' Output: Attribute with content after removing file path
' Example - HREF attribute: (Quotes are part of the strings)
' Input   href="file:///cool folder/help.html"
' Output  href="help.html"

Dim sTemp As String
Dim OpenQuote As Long, CloseQuote As Long, LastSlash As Long

CloseQuote = 0

sTemp = Attr
OpenQuote = InStr(1, sTemp, Quote)
CloseQuote = InStr(OpenQuote + 1, sTemp, Quote)

If CloseQuote <> 0 Then
    '''''Quotes Found, Handle Path
    sTemp = Mid$(sTemp, OpenQuote, CloseQuote - OpenQuote)
    sTemp = Replace$(Attr, "\", "/") ' just in case
    LastSlash = InStrRev(sTemp, "/")
    If LastSlash = 0 Then
        'No path info, do nothing.
        RemovePath = sTemp
    Else
        'Remove Path:
        sTemp = Mid$(sTemp, LastSlash + 1, Len(sTemp) - LastSlash - 1)
        RemovePath = Left$(Attr, OpenQuote - 1) & Quote & sTemp & Quote
    End If

Else
    '''''Quotes NOT Found, do nothing
    RemovePath = sTemp
End If

'MsgBox Attr
'MsgBox RemovePath

End Function

Public Function StripBackgroundPathEx(Src As String) As String
' INPUT:  attribute with contents
' Output: attribute with content after removing file path
' Example: (Quotes are part of the string)
' Input   ="file:///cool folder/help.jpg"
' Output  ="help.jpg"

Dim sTemp As String
Dim LastSlash As Long


If Src <> "" Then
    ' 13 = lenngth_of(background=")+1
    sTemp = Mid$(Src, 13, Len(Src) - 13)
    sTemp = Replace$(sTemp, "\", "/") ' just in case
    LastSlash = InStrRev(sTemp, "/")
    If LastSlash = 0 Then
        'No path info, do nothing.
    Else
        'Remove Path:
        sTemp = Mid$(sTemp, LastSlash + 1, Len(sTemp) - LastSlash + 1)
    End If
    
    StripBackgroundPathEx = "BACKGROUND=" & Quote & sTemp & Quote
Else
    StripBackgroundPathEx = ""
End If

MsgBox StripBackgroundPathEx

End Function

Function DoBackgroundAll(Html As String) As String
Dim qpos1 As Integer
Dim qpos2 As Integer
Dim hpos As Integer


Dim xTemp
Dim sTemp As String
Dim StartPos As Long, EndPos As Long
Dim EndChars As String
Dim sTempChars As String
Dim sTempCharsX As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long
Dim sTarget As String


sTemp = Html
EndPos = 1
idx = 1
StartPos = 0
Do
   
   'Find the BACKGROUND attribute
   StartPos = InStr(StartPos + 1, sTemp, "background=", vbTextCompare)
   ' the Opening Quote is at: (UNUSED)
'   StartPos = InStr(StartPos + 1, sTemp, Qout, vbTextCompare)
If StartPos = 0 Then Exit Do
    'Find the Closing Quote '''' 13 = len("BACKGROUND=") + 2
    EndPos = InStr(StartPos + 12, sTemp, Quote, vbTextCompare)
If EndPos = 0 Then Exit Do
    CurrentTag(idx) = Mid$(sTemp, StartPos, EndPos - StartPos + 1)
    sTempCharsX = String$(idx, Chr$(7))
    sTemp = Replace$(sTemp, CurrentTag(idx), sTempCharsX, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx = 1 Then
    'do nothing
Else
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell
End If

For idx = LBound(CurrentTag) To UBound(CurrentTag)
    sTempCharsX = String$(idx, Chr$(7))
    CurrentTag(idx) = RemovePath(CurrentTag(idx))
    sTemp = Replace$(sTemp, sTempCharsX, CurrentTag(idx), 1, 1)
Next idx
 
DoBackgroundAll = sTemp

End Function

Public Function HTML_RemoveImgTags2(ByRef si As String) As String
'<=60  >=62
Dim b() As Byte
Dim c() As Byte
ReDim imgl(0 To 3) As Byte
ReDim imgu(0 To 3) As Byte
Dim InTag  As Boolean, InImg  As Boolean
Dim s100 As String * 100
Dim pos100 As Long
Dim so As String
Dim idx As Long, idxc As Long

ReDim b(0 To Len(si) - 1)
ReDim c(0 To Len(si) - 1)
b = StrConv(si, vbFromUnicode) ' VB Strings are Double-Byte Unicode
imgl = StrConv("<img", vbFromUnicode)
imgu = StrConv("<IMG", vbFromUnicode)

idxc = 0
For idx = 0 To UBound(b)
    If b(idx) = 60 Then
        InTag = True
''''''''''''''''''''''''''''''New Code //Start
'        s100 = Mid$(si, idx + 1, 100) 'read 100 chars ahead
'        pos100 = InStr(s100, ">") 'find first ">"
'        If pos100 > 0 Then idx = idx + pos100      'jump to it
'        ch = Mid$(si, idx, 1)
''''''''''''''''''''''''''''''New Code //End
    End If
    If InTag Then
        If (idx + 1 = InStrB(idx + 1, b, imgl) Or idx + 1 = InStrB(idx + 1, b, imgu)) Then
             InImg = True
             idx = idx + 6  ' 6=length_of[<img ]+1
        End If
    End If
    If b(idx) = 62 Then
        InTag = False
        If InImg Then InImg = False: b(idx) = 0
    End If
    
    If (Not (InImg) And b(idx) <> 0) Then
        c(idxc) = b(idx)
        idxc = idxc + 1
    End If
Next idx


HTML_RemoveImgTags2 = Left$(StrConv(c, vbUnicode), idxc - 1)

End Function

Public Function DoCSS(ByRef sTemp As String) As String

Dim CSS_FilePath  As String, CSS_Contents As String
Dim LinkTag As String
Dim StartPos As Long, EndPos As Long
Dim LinkTagStart As Long, LinkTagEnd As Long
Dim idx As Long


StartPos = InStr(1, sTemp, "<link", vbTextCompare)
If StartPos = 0 Then  ' no link tag
        DoCSS = sTemp
        Exit Function
End If

LinkTagStart = StartPos
LinkTagEnd = InStr(StartPos + 1, sTemp, ">", vbTextCompare)
LinkTag = Mid$(sTemp, LinkTagStart, LinkTagEnd - LinkTagStart + 1)

If StartPos <> 0 Then
    StartPos = InStr(StartPos + 1, sTemp, "href=", vbTextCompare)
End If
If StartPos <> 0 Then
    StartPos = InStr(StartPos + 1, sTemp, Quote, vbTextCompare)
End If
If StartPos <> 0 Then
    EndPos = InStr(StartPos + 1, sTemp, Quote, vbTextCompare)
End If

If ((StartPos = 0) Or (EndPos = 0)) Then
    DoCSS = sTemp
Else
    CSS_FilePath = Mid$(sTemp, StartPos + 1, EndPos - StartPos - 1)
    CSS_FilePath = CurrentDir & "\" & CSS_FilePath
    CSS_Contents = GetTextFileContents(CSS_FilePath)
    CSS_Contents = Make_CSS_Style(CSS_Contents)
    DoCSS = Replace$(sTemp, LinkTag, CSS_Contents, 1, 1, vbTextCompare)
End If

End Function

Public Function Make_CSS_Style(CSS_File_Contents As String) As String
' Convert contents of a CSS file into a
' <STYLE>...</STYLE> tag block

Make_CSS_Style = vbCrLf & "<STYLE type=text/css>" & vbCrLf & _
                 CSS_File_Contents & vbCrLf & "</STYLE>" & vbCrLf

End Function

Public Function StripBackgroundPath(ByVal bg As String) As String
' bg := "cool folder/help.jpg"

Dim sTemp As String
Dim LastSlash As Long

sTemp = bg

sTemp = Replace$(sTemp, "\", "/") ' just in case
LastSlash = InStrRev(sTemp, "/")
If LastSlash = 0 Then
    'do nothing
Else
    sTemp = Mid$(sTemp, LastSlash + 1, Len(sTemp) - LastSlash - 1)
End If

StripBackgroundPath = Quote & sTemp & Quote

End Function

Public Function DoBackground(ByRef sTemp As String) As String


Dim BackGround  As String, Newbackground As String
Dim StartPos As Long, EndPos As Long
Dim idx As Long


StartPos = InStr(1, sTemp, "<body", vbTextCompare)
If StartPos <> 0 Then
    StartPos = InStr(StartPos + 1, sTemp, "background=", vbTextCompare)
End If
If StartPos <> 0 Then
    StartPos = InStr(StartPos + 1, sTemp, Quote, vbTextCompare)
End If
If StartPos <> 0 Then
    EndPos = InStr(StartPos + 1, sTemp, Quote, vbTextCompare)
End If

If ((StartPos = 0) Or (EndPos = 0)) Then
    DoBackground = sTemp
Else
    BackGround = Mid$(sTemp, StartPos, EndPos - StartPos + 1)
    Newbackground = StripBackgroundPath(Mid$(sTemp, StartPos, EndPos - StartPos + 1))
    'MsgBox BackGround
    'MsgBox Newbackground
    DoBackground = Replace$(sTemp, BackGround, Newbackground, 1, 1, vbTextCompare)
End If

End Function
Public Function HTML_RemoveAllTags3(ByRef si As String) As String
' si is passed by Reference to speed things up a little
' Hybrid algorithm:
' Reads char by char, but also uses InStr() to skip some loops

Dim InTag  As Boolean
Dim ch As String * 1
Dim s100 As String * 100
Dim pos100 As Long
Dim so As String
Dim idx As Long, idx2 As Long

so = String$(Len(si), " ") ' Allocate (more than) enough space

For idx = 1 To Len(si)
    ch = Mid$(si, idx, 1)
    If ch = "<" Then
        InTag = True
        ch = ""
''''''''''''''''''''''''''''''New Code //Start
        s100 = Mid$(si, idx + 1, 100) 'read 100 chars ahead
        pos100 = InStr(s100, ">") 'find first ">"
        If pos100 > 0 Then idx = idx + pos100      'jump to it
        ch = Mid$(si, idx, 1)
''''''''''''''''''''''''''''''New Code //End
    End If
    
    
    If ch = ">" Then
        InTag = False
        ch = ""
    End If
    If Not (InTag) Then
        idx2 = idx2 + 1
        Mid$(so, idx2, 1) = ch
    End If
Next idx

HTML_RemoveAllTags3 = Left$(so, idx2)

End Function

Public Function StripHrefPath(Href As String) As String
' href  : href="file:///cool folder/help.html"

Dim sTemp As String
Dim LastSlash As Long

If Href = "" Then
    StripHrefPath = ""
    Exit Function
End If

sTemp = Mid$(Href, 7, Len(Href) - 7)

sTemp = Replace$(sTemp, "\", "/") ' just in case
LastSlash = InStrRev(sTemp, "/")

If LastSlash = 0 Then
    'do nothing
Else
    sTemp = Mid$(sTemp, LastSlash + 1, Len(sTemp) - LastSlash + 1)
End If

StripHrefPath = "HREF=" & Quote & sTemp & Quote

End Function

Function DoHref(Html As String) As String
Dim qpos1 As Integer
Dim qpos2 As Integer
Dim hpos As Integer


Dim xTemp
Dim sTemp As String
Dim StartPos As Long, EndPos As Long
Dim EndChars As String
Dim sTempChars As String
Dim sTempCharsX As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long
Dim sTarget As String


sTemp = Html
EndPos = 1
idx = 1
StartPos = 0
Do
   
   'Find the HREF attribute
   StartPos = InStr(StartPos + 1, sTemp, "href", vbTextCompare)
   ' the Opening Quote is at: (UNUSED)
'   StartPos = InStr(StartPos + 1, sTemp, Qout, vbTextCompare)
If StartPos = 0 Then Exit Do
    'Find the Closing Quote '''' 6 = len("href") + 2
    EndPos = InStr(StartPos + 6, sTemp, Quote, vbTextCompare)
If EndPos = 0 Then Exit Do
    CurrentTag(idx) = Mid$(sTemp, StartPos, EndPos - StartPos + 1)
    'MsgBox CurrentTag(idx)
    sTempCharsX = String$(idx, Chr$(7))
    sTemp = Replace$(sTemp, CurrentTag(idx), sTempCharsX, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx = 1 Then
    'do nothing
Else
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell
End If

For idx = LBound(CurrentTag) To UBound(CurrentTag)
    sTempCharsX = String$(idx, Chr$(7))
    CurrentTag(idx) = StripHrefPath(CurrentTag(idx))
    sTemp = Replace$(sTemp, sTempCharsX, CurrentTag(idx), 1, 1)
Next idx
 
 
DoHref = sTemp

End Function

Function GetImgSrc(sImgTag As String) As String
' FUNCTION: Retrieve the value of the SRC attribute in an IMG tag
' ASSUMPTIONS:
'   - Content of SRC attribute is enclosed in double Quotes
'   - LOWSRC attribute is not present or is after the SRC attribute
' INPUT:   Full IMG tag
' RETURN:  the filename and path in the SRC attribute
'          withoute the Quotes
' Example:
' sImageTag := "<IMG SRC="helpdesk/cool.gif">
' Will Return            helpdesk/cool.gif


Dim SrcPos As Long
Dim FirstQuote As Long
Dim LastQuote As Long

' Find SRC attribute
SrcPos = InStr(1, sImgTag, "src", vbTextCompare)

If SrcPos > 0 Then ' SRC Found, get openinig and closing Quotes
    FirstQuote = InStr(SrcPos + 4, sImgTag, Quote, vbTextCompare) + 1
    LastQuote = InStr(FirstQuote + 1, sImgTag, Quote, vbTextCompare)
End If

If ((FirstQuote > 0) And (LastQuote > 0)) Then
    GetImgSrc = Mid$(sImgTag, FirstQuote, LastQuote - FirstQuote)
Else
    ' No SRC or no Quotes
    GetImgSrc = ""
End If

End Function

Public Function HTML_ValidateImageTags(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long
Dim sImgTag As String
Dim sImgFile As String

iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<img", vbTextCompare)

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, ">", vbTextCompare)

    sImgTag = Mid$(sHTML, iOpeningPos, iClosingPos - iOpeningPos + 1)
    sImgFile = CurrentDir & "\" & GetImgSrc(sImgTag)
    If FileExists(sImgFile) = False Then
        sTempChars = String(iClosingPos - iOpeningPos + 1, Chr$(7))
        Mid$(sHTML, iOpeningPos) = sTempChars
    Else
        ' File exists, so do nothing (keep tag)
    End If

Loop

HTML_ValidateImageTags = Replace(sHTML, Chr$(7), "")

End Function

Function DoSrc(Html As String) As String
Dim qpos1 As Integer
Dim qpos2 As Integer
Dim hpos As Integer


Dim xTemp
Dim sTemp As String
Dim StartPos As Long, EndPos As Long
Dim EndChars As String
Dim sTempChars As String
Dim sTempCharsX As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long
Dim sTarget As String


sTemp = Html
EndPos = 1
idx = 1
StartPos = 0
Do
   
   'Find the SRC attribute
   StartPos = InStr(StartPos + 1, sTemp, "src", vbTextCompare)
   ' the Opening Quote is at: (UNUSED)
'   StartPos = InStr(StartPos + 1, sTemp, Qout, vbTextCompare)
If StartPos = 0 Then Exit Do
    'Find the Closing Quote '''' 5 = len("src") + 2
    EndPos = InStr(StartPos + 5, sTemp, Quote, vbTextCompare)
If EndPos = 0 Then Exit Do
    CurrentTag(idx) = Mid$(sTemp, StartPos, EndPos - StartPos + 1)
    'MsgBox CurrentTag(idx)
    sTempCharsX = String$(idx, Chr$(7))
    sTemp = Replace$(sTemp, CurrentTag(idx), sTempCharsX, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx = 1 Then
    'do nothing
Else
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell
End If

For idx = LBound(CurrentTag) To UBound(CurrentTag)
    sTempCharsX = String$(idx, Chr$(7))
    CurrentTag(idx) = RemovePath(CurrentTag(idx))
    sTemp = Replace$(sTemp, sTempCharsX, CurrentTag(idx), 1, 1)
Next idx
 
 
DoSrc = sTemp

End Function

Function ExtractSRC(Tag As String) As String

Dim qpos1 As Integer
Dim qpos2 As Integer
Dim hpos As Integer

hpos = InStr(LCase(Tag), "src")
qpos1 = InStr(hpos + 1, Tag, Chr(34))
qpos2 = InStr(qpos1 + 1, Tag, Chr(34))
ExtractSRC = LCase(Mid(Tag, qpos1, qpos2 - qpos1 + 1))

End Function

Public Function HTML_RemoveAllTags2(ByRef si As String) As String
' Handle source Char by Char ,
'  A small improvement is achieved by jumping 2 chars when
'  a "<" is found.

Dim InTag  As Boolean
Dim ch As String * 1
Dim so As String
Dim idx As Long, idx2 As Long

so = String$(Len(si), " ")

For idx = 1 To Len(si)
    ch = Mid$(si, idx, 1)
    If ch = "<" Then
        InTag = True
        ch = ""
        idx = idx + 1  'Here we increment the Loop's control variable'
    End If
    
    If ch = ">" Then
        InTag = False
        ch = ""
    End If
    If Not (InTag) Then
        idx2 = idx2 + 1
        Mid$(so, idx2, 1) = ch
    End If
Next idx

HTML_RemoveAllTags2 = Left$(so, idx2)

End Function
Public Function HTML_RemoveAllTags(ByVal sHTML As String) As String
' Depends on InStr(), proved slower than char-by-char proccessig!

Dim sOpenTag As String
Dim sCloseTag As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

sOpenTag = "<"
sCloseTag = ">"

iClosingPos = 1  ' Start of Search

Do

    iOpeningPos = InStr(iClosingPos, sHTML, sOpenTag, vbTextCompare)
    If iOpeningPos = 0 Then Exit Do
    iClosingPos = InStr(iOpeningPos, sHTML, sCloseTag, vbTextCompare)
    If iClosingPos = 0 Then Exit Do
    sTempChars = String(iClosingPos - iOpeningPos + Len(sCloseTag), Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

sHTML = Replace(sHTML, Chr$(7), "")

HTML_RemoveAllTags = sHTML


End Function

Public Function DoLinks(Text As String, KeepHTTP As Boolean, Target As String) As String
'Supported protocols:
' ftp://
' http://
' www.

Dim sTemp As String

sTemp = Text
sTemp = Replace(sTemp, "http://www.", Chr$(7), 1, -1, vbTextCompare)
sTemp = Replace(sTemp, "www.", "http://www.", 1, -1, vbTextCompare)
sTemp = Replace(sTemp, Chr$(7), "http://www.", 1, -1, vbTextCompare)
sTemp = DoHyperLinks(sTemp, "ftp://", True, "") ' no target for "ftp"
sTemp = DoHyperLinks(sTemp, "http://", KeepHTTP, Target)

DoLinks = sTemp

End Function
Public Function BeautifyLink(HyperLink As String, KeepHTTP As Boolean, SmallCase As Boolean) As String
Dim sTemp As String

If LCase(Left(HyperLink, 7)) <> "http://" Then
    sTemp = HyperLink  ' Not HTTP , so do nothing
Else
    If KeepHTTP Then
            sTemp = HyperLink  ' Keep HTTP , so do nothing
    Else
            sTemp = Right(HyperLink, Len(HyperLink) - 7) ' Remove HTTP://
    End If
End If

If SmallCase Then
    sTemp = LCase(sTemp)
End If

BeautifyLink = sTemp

End Function

Public Function DoEMails(Text As String)

Dim sTemp As String
Dim StartPos As Long, EndPos As Long, AtPos As Long
Dim EndChars As String
Dim sTempChars As String
Dim sTempCharsX As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long

'possible delemeters:
EndChars = " ()[],<>" & vbCrLf & vbTab & Quote
sTemp = Text
'++++++++++++++++++++++++++++++++++++++++++++++++++++
EndPos = 1
idx = 1
AtPos = 0
Do
    AtPos = InStr(AtPos + 1, sTemp, "@", vbTextCompare)
    If AtPos = 0 Then Exit Do

    EndPos = MultiInstr(AtPos, sTemp, EndChars, vbTextCompare)
    If EndPos = 0 Then EndPos = Len(sTemp) + 1
    StartPos = MultiInstrRev(AtPos, sTemp, EndChars, vbTextCompare) + 1
    
    CurrentTag(idx) = Mid(sTemp, StartPos, EndPos - StartPos)
   ' MsgBox CurrentTag(idx)
    sTempChars = String(EndPos - StartPos, "X")
    sTempCharsX = String(idx, Chr$(7))
    sTemp = Replace(sTemp, CurrentTag(idx), sTempCharsX, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx = 1 Then
    'do nothing
Else
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell

    For idx = LBound(CurrentTag) To UBound(CurrentTag)
        sTempCharsX = String(idx, Chr$(7))
        CurrentTag(idx) = "<A HREF=" & Quote & "mailto:" & CurrentTag(idx) & Quote & ">" & CurrentTag(idx) & "</A>"
        sTemp = Replace(sTemp, sTempCharsX, CurrentTag(idx), 1, 1)
    Next idx

End If

''++++++++++++++++++++++++++++++++++++++++++++++++++++
'AtPos = InStr(1, sTemp, "@", vbTextCompare)
'EndPos = MultiInstr(AtPos, sTemp, EndChars, vbTextCompare)
'StartPos = MultiInstrRev(AtPos, sTemp, EndChars, vbTextCompare) + 1
'MsgBox Mid(sTemp, StartPos, EndPos - StartPos)

DoEMails = sTemp

End Function



Function xExtractText(sLine As String) As String
Dim idx As Integer
Dim ch As String * 1
Dim sTemp As String
Dim InTag As Boolean

For idx = 1 To Len(sLine)
    ch = Mid(sLine, idx, 1)
    
    Select Case ch
        Case "<"
          InTag = True
        Case ">"
          InTag = False
        Case Else
        'do nothing
    End Select
    
    
    If Not (InTag) Then
        Select Case ch
            Case ">"
                ch = ""
            Case Chr(13), Chr(10), Chr(9)
                ch = " "
            Case Else
            'do nothing
         End Select
    sTemp = sTemp + ch
    End If

Next idx

xExtractText = xDeSpace(sTemp)

End Function


Function xExtractURL(Tag As String) As String
Dim qpos1 As Integer
Dim qpos2 As Integer
Dim hpos As Integer

hpos = InStr(LCase(Tag), "href")
qpos1 = InStr(hpos + 1, Tag, Chr(34))
qpos2 = InStr(qpos1 + 1, Tag, Chr(34))
xExtractURL = LCase(Mid(Tag, qpos1, qpos2 - qpos1 + 1))

End Function

Function xFindTags(SourceText As String, LeftTag As String, RightTag As String, TagArray() As sTag)
Dim pos1 As Long
Dim pos2 As Long
Dim CurrentTag As String
Dim idx As Long
Dim lText As String

lText = LCase(SourceText)
LeftTag = "href="
RightTag = "</a>"

    
pos1 = InStr(lText, LeftTag)
idx = 0
ReDim TagArray(1 To 1)
Do While pos1 <> 0

    pos2 = InStr(pos1 + 1, lText, RightTag)
    CurrentTag = Mid(SourceText, pos1, pos2 - pos1 + 4)
    CurrentTag = "<a " + CurrentTag
    pos1 = InStr(pos2 + 1, lText, LeftTag)
    idx = idx + 1
    ReDim Preserve TagArray(1 To idx)
    'TagArray(idx) = CurrentTag
    TagArray(idx).Href = xExtractURL(CurrentTag)
    TagArray(idx).Text = xExtractText(CurrentTag)
Loop

End Function

Public Function StripSrcPath(Src As String) As String
' INPUT:  SRC attribute with contents
' Output: SRC attribute with content after removing file path
' Example: (Quotes are part of the string)
' Input   src="file:///cool folder/help.jpg"
' Output  SRC="help.jpg"

Dim sTemp As String
Dim LastSlash As Long


If Src <> "" Then
    sTemp = Mid$(Src, 6, Len(Src) - 6)
    sTemp = Replace$(sTemp, "\", "/") ' just in case
    LastSlash = InStrRev(sTemp, "/")
    If LastSlash = 0 Then
        'No path info, do nothing.
    Else
        'Remove Path:
        sTemp = Mid$(sTemp, LastSlash + 1, Len(sTemp) - LastSlash + 1)
    End If
    
    StripSrcPath = "SRC=" & Quote & sTemp & Quote
Else
    StripSrcPath = ""
End If

End Function

Public Function xHTML_RemoveTag(ByVal sHTML As String, ByVal sOpenTag As String, sCloseTag As String) As String
'Removes a HTML Tag with its content

Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, sOpenTag, vbTextCompare)

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, sCloseTag, vbTextCompare)
    sTempChars = String(iClosingPos - iOpeningPos + Len(sCloseTag), Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

sHTML = Replace(sHTML, Chr$(7), "")
If sCloseTag <> ">" Then
    xHTML_RemoveTag = Replace(sHTML, sCloseTag, "")
Else
    xHTML_RemoveTag = sHTML
End If

End Function








Function DoHyperLinks(Text As String, Protocol As String, KeepHTTP As Boolean, Target As String) As String
Dim sTemp As String
Dim StartPos As Long, EndPos As Long
Dim EndChars As String
Dim sTempChars As String
ReDim CurrentTag(1 To 1) As String
Dim idx As Long
Dim sTarget As String

If Target = "" Then
    sTarget = " "
Else
    sTarget = " TARGET=" & Quote & Target & Quote & " "
End If

'Possible Endings:
EndChars = " ,)]<" & vbCrLf & vbTab & Quote

sTemp = Text
EndPos = 1
idx = 1
StartPos = 0
Do
    StartPos = InStr(StartPos + 1, sTemp, Protocol, vbTextCompare)

If StartPos = 0 Then Exit Do

    EndPos = MultiInstr(StartPos, sTemp, EndChars, vbTextCompare)
    CurrentTag(idx) = Mid(sTemp, StartPos, EndPos - StartPos)
    sTempChars = String$(idx, Chr$(7))
    sTemp = Replace$(sTemp, CurrentTag(idx), sTempChars, 1, 1) 'replace only once
    idx = idx + 1
    ReDim Preserve CurrentTag(1 To idx)
Loop

If idx > 1 Then
    ReDim Preserve CurrentTag(1 To idx - 1)  ' Kill the extra cell
End If

For idx = LBound(CurrentTag) To UBound(CurrentTag)
    sTempChars = String$(idx, Chr$(7))
    CurrentTag(idx) = "<A" & sTarget & "HREF=" & Quote & _
                      CurrentTag(idx) & Quote & ">" & _
                      BeautifyLink(CurrentTag(idx), KeepHTTP, False) & _
                      "</A>"
    sTemp = Replace$(sTemp, sTempChars, CurrentTag(idx), 1, 1)
Next idx
 
DoHyperLinks = sTemp

End Function

Public Function AddBR(sText As String, bPre As Boolean) As String
'Adds <BR> tags, and Handles < > & "

Dim sTemp As String
Dim idx As Long
ReDim blines(1 To 1) As String
sTemp = sText
Text2LinesEx sTemp, blines()

sTemp = ""

'REPLACE <  and  >
For idx = LBound(blines) To UBound(blines)
    blines(idx) = Replace(blines(idx), "&", "&amp;")
    blines(idx) = Replace(blines(idx), Chr$(34), "&quot;")
    blines(idx) = Replace(blines(idx), "<", "&lt;")
    blines(idx) = Replace(blines(idx), ">", "&gt;")
Next idx

'ARRAY --> TEXT

If bPre Then
    sTemp = Join(blines(), vbCrLf)
Else
    sTemp = Join(blines(), "<BR>" & vbCrLf)
End If

AddBR = sTemp

End Function


Public Sub Text2LinesEx(Text As String, Lines() As String)
' check if Text is Empty BEFORE calling this sub.

Dim vTemp As Variant
Dim lLBound As Long
Dim lUBound As Long

vTemp = Split(Text, vbCrLf)
lLBound = LBound(vTemp)
lUBound = UBound(vTemp)
ReDim Lines(lLBound To lUBound)

Lines = vTemp

End Sub

Function RevRGB(ByVal VBHexRGB As String) As String
' VB generated Hex RGB must be reversed to be used in HTML

Dim var1 As String
Dim var2 As String
Dim Var3 As String

var1 = Left$(VBHexRGB, 2)
var2 = Mid$(VBHexRGB, 3, 2)
Var3 = Right$(VBHexRGB, 2)

RevRGB = Var3 & var2 & var1

End Function


Public Function HTMLize(Text As String, _
                        PageTitle As String, _
                        PicturePath As String, _
                        PageBackColor As String, _
                        TextFontName As String, _
                        TextColor As String, _
                        TextSize As String, _
                        CopyPicture As Boolean, _
                        BackScroll As Boolean, _
                        TextBold As Boolean, _
                        PreserveSpaces As Boolean, _
                        KeepHTTP As Boolean, _
                        Target As String _
                        ) As String

Dim sHTML As String
Dim sHead As String
Dim sBody As String
Dim sBGPic As String
Dim sBoldOpen As String, sBoldClose As String
Dim sPreOpen As String, sPreClose As String
Dim sFont As String, sBGColor As String, sTextColor As String
Dim sBGScrollable As String
Dim sTarget As String

sHead = "<HEAD>" & vbCrLf
sHead = sHead & "<TITLE>" & PageTitle & "</TITLE>" & vbCrLf & "</HEAD>" & vbCrLf

sFont = "<FONT FACE=" & Chr(34) & TextFontName & Chr(34) & " SIZE=" & TextSize & ">" & vbCrLf

If PicturePath = "" Then
    sBGPic = ""
Else
'    If chkCopy.Value = vbChecked Then
'        sPicFile = ExtractFileName(Trim(txtBGPic.Text))
'        sTgtDir = ExtractDirName(sTgtFile)
'        On Error Resume Next
'        FileCopy Trim(txtBGPic.Text), sTgtDir & sPicFile
'        If Err Then
'            sCopyResult = vbCrLf & "Couldn't copy " & Chr(34) & UCase(sPicFile) & Chr(34)
'        Else
'            sCopyResult = vbCrLf & Chr(34) & UCase(sPicFile) & Chr(34) & " was copied successfully."
'        End If
'        On Error GoTo 0
'    Else
'        sPicFile = Trim(txtBGPic.Text)
'    End If
    sBGPic = " BACKGROUND=" & Chr$(34) & PicturePath & Chr$(34)
End If


If TextBold Then
    sBoldOpen = "<B>" & vbCrLf
    sBoldClose = "</B>" & vbCrLf
Else
    sBoldOpen = ""
    sBoldClose = ""
End If

If PreserveSpaces Then
    sPreOpen = vbCrLf & "<PRE>" & vbCrLf
    sPreClose = vbCrLf & "</PRE>" & vbCrLf
    sHTML = AddBR(Text, True)
Else
    sPreOpen = ""
    sPreClose = ""
    sHTML = AddBR(Text, False)
End If

sHTML = DoLinks(sHTML, KeepHTTP, Target) ' http://  ftp://  www. (will ad http:// to it | IS IT A BUG?)

If BackScroll Then
    sBGScrollable = ""
Else
    sBGScrollable = " BGPROPERTIES = FIXED "
End If

sBody = "<BODY BGCOLOR=" & PageBackColor & " TEXT=" & TextColor & sBGPic & sBGScrollable & ">" & vbCrLf
sBody = sBody & sPreOpen & sFont & sBoldOpen


sHTML = "<HTML>" & vbCrLf & sHead & sBody & sHTML & sBoldClose & "</FONT>" & sPreClose & "</BODY>" & vbCrLf & "</HTML>"

HTMLize = sHTML

End Function

Public Function ColorToHex(ByVal lColor As Long) As String
Dim sTemp As String

sTemp = Hex$(lColor)

If Len(sTemp) < 6 Then sTemp = String(6 - Len(sTemp), "0") + sTemp
sTemp = Chr(34) & "#" & RevRGB(sTemp) & Chr(34)

ColorToHex = sTemp

End Function


Public Function HTML_RemoveTag(ByVal sHTML As String, ByVal sOpenTag As String, sCloseTag As String) As String
'Removes a HTML Tag with its content

Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, sOpenTag, vbTextCompare)

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, sCloseTag, vbTextCompare)
If iClosingPos = 0 Then Exit Do
    sTempChars = String$(iClosingPos - iOpeningPos + Len(sCloseTag), Chr$(7))
    Mid$(sHTML, iOpeningPos) = sTempChars
Loop

sHTML = Replace$(sHTML, Chr$(7), "")
If sCloseTag <> ">" Then
    HTML_RemoveTag = Replace$(sHTML, sCloseTag, "")
Else
    HTML_RemoveTag = sHTML
End If

End Function

Public Function HTML_RemoveScripts(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<script", vbTextCompare)
'MsgBox iOpeningPos

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, "</script>", vbTextCompare)
        
'MsgBox iClosingPos
    sTempChars = String(iClosingPos - iOpeningPos + 9, Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

HTML_RemoveScripts = Replace(sHTML, Chr$(7), "")

End Function



Public Function HTML_RemoveIFrameTags(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<iframe", vbTextCompare)
'MsgBox iOpeningPos

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, ">", vbTextCompare)
        
'MsgBox iClosingPos
    sTempChars = String(iClosingPos - iOpeningPos + 1, Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

sHTML = Replace(sHTML, Chr$(7), "")
HTML_RemoveIFrameTags = Replace(sHTML, "</iframe>", "")
End Function

Public Function HTML_RemoveImageTags(ByVal sHTML As String) As String
Dim sTempChars  As String
Dim iOpeningPos As Long
Dim iClosingPos As Long

'iOpeningPos = -1 ' Non-Zero value
iClosingPos = 1  ' Start of Search

Do
    iOpeningPos = InStr(iClosingPos, sHTML, "<img", vbTextCompare)
'MsgBox iOpeningPos

If iOpeningPos = 0 Then Exit Do
    
    iClosingPos = InStr(iOpeningPos, sHTML, ">", vbTextCompare)
        
'MsgBox iClosingPos
    sTempChars = String(iClosingPos - iOpeningPos + 1, Chr$(7))
    Mid(sHTML, iOpeningPos) = sTempChars
Loop

HTML_RemoveImageTags = Replace(sHTML, Chr$(7), "")

End Function



