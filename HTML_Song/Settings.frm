VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Margins and Background Color"
   ClientHeight    =   2370
   ClientLeft      =   2610
   ClientTop       =   2070
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtColor 
      Height          =   285
      Left            =   1575
      TabIndex        =   11
      Top             =   1305
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Caption         =   " Margins "
      Height          =   915
      Left            =   135
      TabIndex        =   4
      Top             =   135
      Width           =   3975
      Begin VB.TextBox txtLeft 
         Height          =   360
         Left            =   525
         TabIndex        =   7
         Top             =   345
         Width           =   555
      End
      Begin VB.TextBox txtRight 
         Height          =   360
         Left            =   1935
         TabIndex        =   6
         Top             =   345
         Width           =   555
      End
      Begin VB.TextBox txtTop 
         Height          =   360
         Left            =   3195
         TabIndex        =   5
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   375
         Width           =   270
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right"
         Height          =   195
         Left            =   1485
         TabIndex        =   9
         Top             =   375
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top"
         Height          =   195
         Left            =   2790
         TabIndex        =   8
         Top             =   360
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3105
      TabIndex        =   3
      Top             =   1980
      Width           =   1050
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   1890
      TabIndex        =   2
      Top             =   1980
      Width           =   1050
   End
   Begin XSong.ColorButton ClrBtn1 
      Height          =   330
      Left            =   3195
      TabIndex        =   1
      Top             =   1305
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background Color"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   1350
      Width           =   1275
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClrBtn1_Click()

txtColor.Text = "#" & LongColorToHex(ClrBtn1.Color)

End Sub

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdOk_Click()


Me.Hide

End Sub
Private Sub Form_Load()

ClrBtn1.hwndOwner = Me.hwnd
ClrBtn1.Color = vbWhite


End Sub

Private Sub txtLeft_Click()
SelectAllText txtLeft

End Sub

Private Sub txtRight_Click()
SelectAllText txtRight

End Sub


Private Sub txtTop_Click()
SelectAllText txtTop

End Sub


