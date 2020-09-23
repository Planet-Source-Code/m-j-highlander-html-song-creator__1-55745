Attribute VB_Name = "HTML_Functions"
Option Explicit

Public Function AddBROnly(ByVal Text As String) As String

Text = Replace$(Text, vbCrLf, "<BR>" & vbCrLf)

AddBROnly = Text

End Function
Function Html2Text(ByVal Html As String) As String

Html = Replace$(Html, "<title>", "<!--", 1, -1, vbTextCompare)
Html = Replace$(Html, "</title>", "-->", 1, -1, vbTextCompare)

Html = Replace$(Html, vbCrLf, " ", 1, -1, vbTextCompare)
Html = Replace$(Html, "<br>", vbCrLf, 1, -1, vbTextCompare)
Html = Replace$(Html, "<p>", vbCrLf, 1, -1, vbTextCompare)
Html = Replace$(Html, "</p>", vbCrLf, 1, -1, vbTextCompare)
Html = RemoveAllHTMLTags(Html)

Html = TrimSpaces(Html, True, True, True)
Html = CompactSpaces(Html)
Html = AddBROnly(Html)
Html2Text = Html

End Function

Public Function RemoveAllHTMLTags(ByRef si As String) As String
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

RemoveAllHTMLTags = Left$(so, idx2)

End Function

