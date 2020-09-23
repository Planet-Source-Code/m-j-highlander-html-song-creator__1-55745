VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWeb 
   BorderStyle     =   0  'None
   Caption         =   "Preview"
   ClientHeight    =   3525
   ClientLeft      =   1995
   ClientTop       =   1935
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   6660
      TabIndex        =   1
      Top             =   0
      Width           =   6660
      Begin VB.CommandButton Command2 
         Caption         =   "No Way, delete now!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2430
         TabIndex        =   3
         Top             =   0
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         Caption         =   "It's Cool, keep it"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   2
         Top             =   0
         Width           =   2130
      End
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2895
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   6495
      ExtentX         =   11456
      ExtentY         =   5106
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
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mb_Kill As Boolean

Public Property Get Kill() As Boolean
       Kill = mb_Kill
End Property

Public Property Let Kill(ByVal bNewValue As Boolean)
       mb_Kill = bNewValue
End Property

Private Sub Command1_Click()

Kill = False
frmWeb.web.Navigate "about:blank"
Me.Hide

End Sub
Private Sub Command2_Click()

Kill = True
frmWeb.web.Navigate "about:blank"
Me.Hide

End Sub

Private Sub Form_Resize()
web.Top = picTop.Height
web.Left = 0
web.Width = Me.ScaleWidth
web.Height = Me.ScaleHeight - picTop.Height
End Sub


