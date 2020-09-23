VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmFont 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font Options"
   ClientHeight    =   4935
   ClientLeft      =   3060
   ClientTop       =   825
   ClientWidth     =   5490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Preview "
      Height          =   2085
      Left            =   90
      TabIndex        =   15
      Top             =   2295
      Width           =   5370
      Begin VB.CommandButton Command3 
         Height          =   195
         Left            =   3645
         Picture         =   "Font.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   1590
      End
      Begin SHDocVwCtl.WebBrowser web 
         Height          =   1770
         Left            =   135
         TabIndex        =   17
         Top             =   225
         Width           =   5100
         ExtentX         =   8996
         ExtentY         =   3122
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      Height          =   330
      Left            =   3960
      TabIndex        =   14
      Top             =   315
      Width           =   690
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   330
      Left            =   3060
      TabIndex        =   13
      Top             =   315
      Width           =   690
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   285
      Left            =   4725
      TabIndex        =   12
      Top             =   1755
      Width           =   645
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   3015
      TabIndex        =   11
      Top             =   4455
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4230
      TabIndex        =   10
      Top             =   4455
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   285
      Left            =   4005
      TabIndex        =   9
      Top             =   1305
      Width           =   645
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   765
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1305
      Width           =   3165
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   765
      Width           =   1545
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      Left            =   675
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   315
      Width           =   645
   End
   Begin XSong.ColorButton ClrBtn1 
      Height          =   285
      Left            =   2340
      TabIndex        =   4
      Top             =   765
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   503
   End
   Begin VB.TextBox txtFaces 
      Height          =   285
      Left            =   1125
      TabIndex        =   0
      Top             =   1755
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fonts"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1350
      Width           =   390
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   810
      Width           =   360
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   315
      Width           =   300
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font Face"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   1755
      Width           =   720
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UpdatePreview(Optional BlackBGColor As Boolean)
Dim s As String, f As String

s = Me.Tag
If Me.chkBold Then s = "<B>" & s & "</B>"
If Me.chkItalic Then s = "<I>" & s & "</I>"
f = "<FONT " & "SIZE=" & cboSize.Text & " COLOR=" & txtColor.Text & " FACE=" & txtFaces.Text & ">"
s = f & s & "</FONT>"

If BlackBGColor Then
    s = "<BODY BGCOLOR=BLACK>" & s & "</BODY>"
End If

Me.web.Navigate "about:" & s

End Sub

Private Sub cboSize_Click()
UpdatePreview
End Sub


Private Sub chkBold_Click()
UpdatePreview
End Sub


Private Sub chkItalic_Click()
UpdatePreview
End Sub


Private Sub ClrBtn1_Click()

txtColor.Text = "#" & LongColorToHex(ClrBtn1.Color)

End Sub

Private Sub cmdCancel_Click()

frmFont.txtFaces.Text = ""
frmFont.txtColor.Text = ""
Me.Hide

End Sub

Private Sub cmdOk_Click()

If frmFont.txtFaces.Text = "" Then frmFont.txtFaces.Text = "Verdana,Tahoma,Arial"
If frmFont.cboSize.Text = "" Then frmFont.cboSize.ListIndex = 1
If frmFont.txtColor.Text = "" Then frmFont.txtColor.Text = "#000000"


Me.Hide

End Sub


Private Sub Command1_Click()

If Trim$(txtFaces.Text) = "" Then
    txtFaces.Text = cboFont.List(cboFont.ListIndex)
Else
    txtFaces.Text = Trim$(txtFaces.Text) & "," & cboFont.List(cboFont.ListIndex)
End If

End Sub
Private Sub Label2_Click()

End Sub

Private Sub lbl_cmdCancel_Click()
Me.Hide
End Sub


Private Sub lbl_cmdAddFont_Click()

End Sub


Private Sub Picture1_Click()

End Sub


Private Sub Command2_Click()
txtFaces.Text = ""
End Sub

Private Sub Command3_Click()
Static bState As Boolean
bState = Not bState

UpdatePreview bState

End Sub

Private Sub Form_Activate()
UpdatePreview

End Sub

Private Sub Form_Load()
Dim idx As Integer

For idx = 1 To 7
    cboSize.AddItem CStr(idx)
Next idx
cboSize.ListIndex = 1

For idx = 0 To Screen.FontCount - 1
    cboFont.AddItem Screen.Fonts(idx)
Next idx
cboFont.ListIndex = 0

ClrBtn1.Color = vbBlack

End Sub

Private Sub txtColor_Change()
UpdatePreview
End Sub


Private Sub txtFaces_Change()
UpdatePreview
End Sub


