VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   1995
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   6585
   Begin Project1.ScalablePic RichTextBox1 
      Height          =   1500
      Index           =   0
      Left            =   630
      TabIndex        =   1
      Top             =   855
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   2646
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Command1"
      Height          =   1185
      Left            =   4905
      TabIndex        =   0
      Top             =   1080
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
Dim CDlg As New CdlgEx
  'CDlg.InitDir = ExtractDirName(txtBGPic.Text)
  CDlg.hOwner = Me.hWnd
  CDlg.Left = 100
  CDlg.Top = 100
  CDlg.OKText = "Open"
  CDlg.CancelText = "Cancel"
  CDlg.HelpText = "No Help"
  CDlg.DialogTitle = "Select Image File"
  ' CDlg.CancelError = True
  CDlg.Filter = "Picture Files|*.bmp;*.gif;*.jpg;*.ico;*.wmf|All files|*.*"
  CDlg.flags = &H4 Or &H1000
  Load RichTextBox1(1)

Set rtb = RichTextBox1(1)
CDlg.ShowOpen
If RichTextBox1.Count > 1 Then Unload RichTextBox1(1)

If CDlg.FileName <> "" Then Print CDlg.FileName
Set CDlg = Nothing

End Sub

