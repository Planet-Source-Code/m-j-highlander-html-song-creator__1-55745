VERSION 5.00
Begin VB.Form frmSplitPic 
   BackColor       =   &H00404000&
   Caption         =   "NON-VISUAL FORM !!!"
   ClientHeight    =   3525
   ClientLeft      =   1410
   ClientTop       =   1800
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   Begin VB.PictureBox picDn 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   495
      ScaleHeight     =   1035
      ScaleWidth      =   1455
      TabIndex        =   2
      Top             =   705
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   705
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2790
      Left            =   2610
      ScaleHeight     =   2790
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   405
      Visible         =   0   'False
      Width           =   2010
   End
End
Attribute VB_Name = "frmSplitPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function SplitPic(PicFile As String, UpPicFile As String, DnPicFile As String, iUp As Integer)
' sample call
' SplitPic "G:\XSong\Beach.jpg", "C:\up.bmp", "c:\down.bmp", 100

Me.ScaleMode = vbPixels
picOriginal.ScaleMode = vbPixels
picUp.ScaleMode = vbPixels
picDn.ScaleMode = vbPixels

picOriginal.AutoRedraw = True
picUp.AutoRedraw = True
picDn.AutoRedraw = True

picOriginal.Picture = LoadPicture(PicFile)
picUp.Width = picOriginal.Width
picUp.Height = iUp

picDn.Width = picOriginal.Width
picDn.Height = picOriginal.Height - iUp

picUp.PaintPicture picOriginal.Picture, 0, 0, , , , , picOriginal.Width, picUp.Height
picDn.PaintPicture picOriginal.Picture, 0, 0, , , , iUp, picOriginal.Width, picDn.Height

SavePicture picUp.Image, UpPicFile
SavePicture picDn.Image, DnPicFile

End Function



