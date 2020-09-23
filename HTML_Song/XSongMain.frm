VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmXSongMain 
   Caption         =   "XSong"
   ClientHeight    =   6675
   ClientLeft      =   630
   ClientTop       =   795
   ClientWidth     =   9750
   Icon            =   "XSongMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6450
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   11377
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "XSongMain.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cdlg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkAutoPreview"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Media Player && Lyrics"
      TabPicture(1)   =   "XSongMain.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkCenter"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "cmdLyricsSettings"
      Tab(1).Control(3)=   "Command7"
      Tab(1).Control(4)=   "txtLyrics"
      Tab(1).Control(5)=   "Command3"
      Tab(1).Control(6)=   "Label12"
      Tab(1).Control(7)=   "Label5"
      Tab(1).ControlCount=   8
      Begin VB.CheckBox chkCenter 
         Caption         =   "Center"
         Height          =   240
         Left            =   -70185
         TabIndex        =   44
         Top             =   1575
         Width           =   1005
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Create !"
         Height          =   1275
         Left            =   5715
         TabIndex        =   43
         Top             =   2925
         Width           =   1725
      End
      Begin VB.CheckBox chkAutoPreview 
         Caption         =   "Preview"
         Height          =   210
         Left            =   6030
         TabIndex        =   42
         Top             =   4365
         Value           =   1  'Checked
         Width           =   1110
      End
      Begin VB.Frame Frame3 
         Caption         =   " Upper Frame "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -74865
         TabIndex        =   32
         Top             =   405
         Width           =   9195
         Begin VB.TextBox txtUpperHeight 
            Height          =   375
            Left            =   2070
            TabIndex        =   40
            Top             =   315
            Width           =   555
         End
         Begin VB.TextBox txtHeight 
            Height          =   360
            Left            =   7200
            TabIndex        =   37
            Top             =   450
            Width           =   555
         End
         Begin VB.TextBox txtWidth 
            Height          =   360
            Left            =   5760
            TabIndex        =   36
            Top             =   450
            Width           =   555
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Margins and BG Color..."
            Height          =   645
            Left            =   3105
            TabIndex        =   34
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Height of Upper Frame"
            Height          =   195
            Left            =   360
            TabIndex        =   41
            Top             =   405
            Width           =   1605
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            Height          =   195
            Left            =   6615
            TabIndex        =   39
            Top             =   525
            Width           =   465
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            Height          =   195
            Left            =   5265
            TabIndex        =   38
            Top             =   540
            Width           =   420
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Media Player Control:"
            Height          =   195
            Left            =   4860
            TabIndex        =   35
            Top             =   225
            Width           =   1500
         End
      End
      Begin VB.CommandButton cmdLyricsSettings 
         Caption         =   "Margins and BG Color..."
         Height          =   285
         Left            =   -68655
         TabIndex        =   31
         Top             =   1530
         Width           =   2130
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Font..."
         Height          =   285
         Left            =   -66405
         TabIndex        =   30
         Top             =   1530
         Width           =   825
      End
      Begin VB.Frame Frame2 
         Caption         =   " Files "
         Height          =   5865
         Left            =   225
         TabIndex        =   11
         Top             =   450
         Width           =   5010
         Begin VB.CheckBox chkPicPreview 
            Caption         =   "Preview Pic"
            Height          =   210
            Left            =   3600
            TabIndex        =   46
            Top             =   5535
            Value           =   1  'Checked
            Width           =   1245
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Font..."
            Height          =   285
            Left            =   4050
            TabIndex        =   29
            Top             =   2115
            Width           =   825
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Font..."
            Height          =   285
            Left            =   4050
            TabIndex        =   28
            Top             =   1665
            Width           =   825
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Browse"
            Height          =   285
            Left            =   4005
            TabIndex        =   26
            Top             =   1185
            Width           =   915
         End
         Begin VB.TextBox txtFolder 
            Height          =   300
            Left            =   1395
            TabIndex        =   25
            Top             =   1170
            Width           =   2500
         End
         Begin VB.TextBox txtArtist 
            Height          =   300
            Left            =   1395
            TabIndex        =   17
            Top             =   2115
            Width           =   2500
         End
         Begin VB.TextBox txtSongTitle 
            Height          =   300
            Left            =   1395
            TabIndex        =   16
            Top             =   1665
            Width           =   2500
         End
         Begin VB.TextBox txtBackgroundFile 
            Height          =   300
            Left            =   1395
            OLEDropMode     =   1  'Manual
            TabIndex        =   15
            Top             =   750
            Width           =   2500
         End
         Begin VB.TextBox txtSoundFile 
            Height          =   300
            Left            =   1395
            OLEDropMode     =   1  'Manual
            TabIndex        =   14
            Top             =   360
            Width           =   2500
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Browse"
            Height          =   285
            Left            =   4005
            TabIndex        =   13
            Top             =   360
            Width           =   915
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Browse"
            Height          =   285
            Left            =   4005
            TabIndex        =   12
            Top             =   765
            Width           =   915
         End
         Begin XSong.ScalablePic SPic1 
            Height          =   2850
            Left            =   360
            TabIndex        =   18
            Top             =   2610
            Width           =   4470
            _ExtentX        =   7885
            _ExtentY        =   5027
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
         Begin XSong.ScalablePic RichTextBox1 
            Height          =   870
            Index           =   0
            Left            =   675
            TabIndex        =   45
            Top             =   3060
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1535
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
         Begin VB.Label lblDims 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "                                        "
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   360
            TabIndex        =   47
            Top             =   5535
            Width           =   1800
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Output Folder"
            Height          =   195
            Left            =   270
            TabIndex        =   27
            Top             =   1200
            Width           =   960
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Artist"
            Height          =   195
            Left            =   945
            TabIndex        =   22
            Top             =   2115
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Song Title"
            Height          =   195
            Left            =   540
            TabIndex        =   21
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Background File"
            Height          =   195
            Left            =   90
            TabIndex        =   20
            Top             =   780
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sound File"
            Height          =   195
            Left            =   540
            TabIndex        =   19
            Top             =   405
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Layout "
         Height          =   1815
         Left            =   5310
         TabIndex        =   4
         Top             =   540
         Visible         =   0   'False
         Width           =   4110
         Begin VB.TextBox txtLeftWidth 
            Height          =   375
            Left            =   3240
            TabIndex        =   24
            Top             =   495
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.PictureBox picLayout 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1125
            Index           =   2
            Left            =   1845
            Picture         =   "XSongMain.frx":0D02
            ScaleHeight     =   1125
            ScaleWidth      =   1500
            TabIndex        =   7
            Top             =   405
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox picLayout 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1125
            Index           =   1
            Left            =   1440
            Picture         =   "XSongMain.frx":1CC0
            ScaleHeight     =   1125
            ScaleWidth      =   1500
            TabIndex        =   6
            Top             =   450
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.PictureBox picLayout 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1140
            Index           =   0
            Left            =   1215
            Picture         =   "XSongMain.frx":2C7E
            ScaleHeight     =   1140
            ScaleWidth      =   1500
            TabIndex        =   5
            Top             =   495
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label Label10 
            Caption         =   "Width of Left Frame"
            Height          =   555
            Left            =   2250
            TabIndex        =   23
            Top             =   495
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Single Page"
            Height          =   195
            Left            =   270
            TabIndex        =   10
            Top             =   270
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frames Up / Down"
            Height          =   195
            Left            =   765
            TabIndex        =   9
            Top             =   855
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frames Left / Right"
            Height          =   195
            Left            =   1305
            TabIndex        =   8
            Top             =   270
            Visible         =   0   'False
            Width           =   1365
         End
      End
      Begin VB.TextBox txtLyrics 
         Height          =   4470
         Left            =   -74865
         MultiLine       =   -1  'True
         OLEDropMode     =   1  'Manual
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   1890
         Width           =   9255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Browse"
         Height          =   285
         Left            =   -71850
         TabIndex        =   1
         Top             =   1575
         Width           =   915
      End
      Begin MSComDlg.CommonDialog cdlg 
         Left            =   4995
         Top             =   5940
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lyrics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74730
         TabIndex        =   33
         Top             =   1620
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Text or HTML): Type, Paste or"
         Height          =   195
         Left            =   -74145
         TabIndex        =   3
         Top             =   1620
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmXSongMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPicPreview_Click()
If chkPicPreview.Value = vbUnchecked Then
    Set SPic1.Picture = LoadPicture()
    lblDims.Caption = ""
End If

End Sub

Private Sub cmdLyricsSettings_Click()

Load frmSettings
frmSettings.ClrBtn1.Color = HexColorToVBHex(Replace$(LyricsBGColor, "#", ""))
frmSettings.txtColor.Text = "#" & LongColorToHex(frmSettings.ClrBtn1.Color)
With LyricsMargins
    frmSettings.txtLeft.Text = .LeftMargin
    frmSettings.txtRight.Text = .RightMargin
    frmSettings.txtTop.Text = .TopMargin
End With

frmSettings.Show vbModal
If frmSettings.txtColor = "" Then Exit Sub 'Was Canceled

LyricsBGColor = Trim$(frmSettings.txtColor)
With LyricsMargins
    .LeftMargin = Trim$(frmSettings.txtLeft.Text)
    .RightMargin = Trim$(frmSettings.txtRight.Text)
    .TopMargin = Trim$(frmSettings.txtTop.Text)
    If (.LeftMargin = "" Or .RightMargin = "" Or .TopMargin = "") Then
        .LeftMargin = 12
        .RightMargin = 12
        .TopMargin = 12
    End If
End With
If LyricsBGColor = "" Then LyricsBGColor = "#FFFFFF" 'white

Unload frmSettings

End Sub
Private Sub Command1_Click()
On Error Resume Next
Dim vArray As Variant
ReDim sArray(0 To 1) As String
cdlg.FileName = ""
cdlg.Filter = "*.MP?;*.WAV|*.mp?;*.wav|All Files|*.*"
cdlg.flags = &H4&
cdlg.ShowOpen
If Err Then Exit Sub

txtSoundFile.Text = cdlg.FileName
vArray = Read_ID3Tag(cdlg.FileName)
If IsArray(vArray) Then
    txtSongTitle.Text = vArray(1)
    txtArtist.Text = vArray(2)
Else
    'now try to GUESS the info from FileName:
    sArray = HeuristicSongName(cdlg.FileName)
    txtArtist.Text = sArray(0)
    txtSongTitle.Text = sArray(1)
End If

End Sub

Private Sub Command2_Click()
'On Error Resume Next
'Dim vArray As Variant
'cdlg.Filename = ""
'cdlg.Filter = "*.BMP;*.GIF;*.JPG|*.bmp;*.gif;*.jpg|All Files|*.*"
'cdlg.flags = &H4&
'cdlg.ShowOpen
'If Err Then Exit Sub
'
'txtBackgroundFile.Text = cdlg.Filename
'Set SPic1.Picture = LoadPicture(cdlg.Filename)

'///NEW CODE
Dim cdlg As New CdlgEx
  cdlg.InitDir = ExtractDirName(txtBackgroundFile.Text)
  cdlg.hOwner = Me.hWnd
  cdlg.Left = 100
  cdlg.Top = 100
  cdlg.OKText = "Open"
  cdlg.CancelText = "Cancel"
  cdlg.HelpText = "No Help"
  cdlg.DialogTitle = "Select Image File"
  ' CDlg.CancelError = True
  cdlg.Filter = "Picture Files|*.bmp;*.gif;*.jpg;*.ico;*.wmf|All files|*.*"
  cdlg.flags = &H4 Or &H1000
  Load RichTextBox1(1)

Set rtb = RichTextBox1(1)
cdlg.ShowOpen
If RichTextBox1.Count > 1 Then Unload RichTextBox1(1)

If cdlg.FileName <> "" Then
    txtBackgroundFile.Text = cdlg.FileName
    If chkPicPreview.Value = vbChecked Then
        Set SPic1.Picture = LoadPicture(cdlg.FileName)
        lblDims.Caption = CStr(SPic1.ImageWidth) & " x " & CStr(SPic1.ImageHeight)
    End If
End If

Set cdlg = Nothing

End Sub
Private Sub lbl_cmdLyric_Click()
End Sub


Private Sub Command3_Click()
On Error Resume Next
Dim vArray As Variant
cdlg.FileName = ""
cdlg.Filter = "Text and HTML Files|*.txt;*.htm;*.html|All Files|*.*"
cdlg.flags = &H4&
cdlg.ShowOpen
If Err Then Exit Sub

txtLyrics.Text = LoadFile(cdlg.FileName)
If Err Then MsgBox "Couldn't Load File" + vbCrLf + UCase(cdlg.FileName) + vbCrLf + Error + "  " + Str(Err), vbCritical, "Error"

If IsHTML(txtLyrics.Text) Then
    txtLyrics.Text = Html2Text(txtLyrics.Text)
End If

End Sub
Private Sub Command4_Click()
Dim folder As String
folder = BrowseForFolder(Me, "Select Output Folder", CStr(txtFolder.Text))
If Len(folder) <> 0 Then
    txtFolder.Text = folder
End If

End Sub

Private Sub Command5_Click()
Load frmFont
frmFont.txtFaces.Text = SongFont.Face
If SongFont.Size = "" Then SongFont.Size = "2"
frmFont.cboSize.ListIndex = CInt(SongFont.Size) - 1
frmFont.txtColor.Text = SongFont.Color
If SongFont.Color = "" Then SongFont.Color = "0"
frmFont.ClrBtn1.Color = HexColorToVBHex(Replace$(SongFont.Color, "#", ""))
frmFont.chkBold.Value = Abs(CInt(SongFont.Bold))
frmFont.chkItalic.Value = Abs(CInt(SongFont.Italic))
frmFont.Tag = txtSongTitle
frmFont.Show vbModal

If frmFont.txtColor.Text <> "" Then      ' not Canceled
    SongFont.Face = frmFont.txtFaces.Text
    SongFont.Size = frmFont.cboSize.Text
    SongFont.Color = frmFont.txtColor.Text
    SongFont.Bold = IIf(frmFont.chkBold.Value = vbChecked, True, False)
    SongFont.Italic = IIf(frmFont.chkItalic.Value = vbChecked, True, False)
    Unload frmFont
End If

Unload frmFont

End Sub
Private Sub Command6_Click()
Load frmFont
frmFont.txtFaces.Text = ArtistFont.Face
If ArtistFont.Size = "" Then ArtistFont.Size = "2"
frmFont.cboSize.ListIndex = CInt(ArtistFont.Size) - 1
frmFont.txtColor.Text = ArtistFont.Color
If ArtistFont.Color = "" Then ArtistFont.Color = "0"
frmFont.ClrBtn1.Color = HexColorToVBHex(Replace$(ArtistFont.Color, "#", ""))
frmFont.chkBold.Value = Abs(CInt(ArtistFont.Bold))
frmFont.chkItalic.Value = Abs(CInt(ArtistFont.Italic))
frmFont.Tag = txtArtist.Text
frmFont.Show vbModal

If frmFont.txtColor.Text <> "" Then      ' not Canceled
    ArtistFont.Face = frmFont.txtFaces.Text
    ArtistFont.Size = frmFont.cboSize.Text
    ArtistFont.Color = frmFont.txtColor.Text
    ArtistFont.Bold = IIf(frmFont.chkBold.Value = vbChecked, True, False)
    ArtistFont.Italic = IIf(frmFont.chkItalic.Value = vbChecked, True, False)
    Unload frmFont
End If

Unload frmFont

End Sub
Private Sub Command7_Click()
Load frmFont
frmFont.txtFaces.Text = LyricsFont.Face
If LyricsFont.Size = "" Then LyricsFont.Size = "2"
frmFont.cboSize.ListIndex = CInt(LyricsFont.Size) - 1
frmFont.txtColor.Text = LyricsFont.Color
If LyricsFont.Color = "" Then LyricsFont.Color = "0"
frmFont.ClrBtn1.Color = HexColorToVBHex(Replace$(LyricsFont.Color, "#", ""))
frmFont.chkBold.Value = Abs(CInt(LyricsFont.Bold))
frmFont.chkItalic.Value = Abs(CInt(LyricsFont.Italic))
frmFont.Tag = txtLyrics.Text
frmFont.Show vbModal

If frmFont.txtColor.Text <> "" Then      ' not Canceled
    LyricsFont.Face = frmFont.txtFaces.Text
    LyricsFont.Size = frmFont.cboSize.Text
    LyricsFont.Color = frmFont.txtColor.Text
    LyricsFont.Bold = IIf(frmFont.chkBold.Value = vbChecked, True, False)
    LyricsFont.Italic = IIf(frmFont.chkItalic.Value = vbChecked, True, False)
End If

Unload frmFont

End Sub
Private Sub lbl_cmdArtistName_Click()

End Sub

Private Sub Command8_Click()
Dim sTitle As String, sUpHeight As String, sFileName As String
Dim sLyrics As String, NewSongFile As String, ImageFile As String
Dim sTemp As String, sBIOpen As String, sBIClose As String

SongFile = Trim$(txtSoundFile.Text)
OutputFolder = Trim$(txtFolder.Text)
SongTitle = Trim$(txtSongTitle.Text)
ArtistName = Trim$(txtArtist.Text)
ImageFile = Trim$(txtBackgroundFile.Text)
UpperFrameHeight = Trim$(txtUpperHeight.Text)

If (OutputFolder = "" Or SongTitle = "" Or ArtistName = "" Or SongFile = "") Then
    Beep
    MsgBox "Missing Info", vbCritical, "Oops!"
    Exit Sub
End If
sTemp = LoadFile(AppPath() & "templates\frames.htm")
sTitle = ArtistName & " - " & SongTitle
OutputFolder = AddSlash(OutputFolder)
OutputFolder = OutputFolder & RemoveInvalidChars(sTitle)
If (DirExists(OutputFolder) = False) Then
       If CreatePath(OutputFolder) = False Then
            MsgBox "Cannot Create Folder", vbCritical, "Message"
            Exit Sub
       End If
End If
OutputFolder = AddSlash(OutputFolder)
On Error Resume Next ''''''''''ERROR TRAP BEGIN
MkDir OutputFolder & "files"    '<-----------possible error if exists
NewSongFile = OutputFolder & "files\" & ExtractFileName(SongFile)
FileCopy SongFile, NewSongFile  '<-----------possible error if exists
On Error GoTo 0      ''''''''''ERROR TRAP END

SongFile = ExtractFileName(NewSongFile)
''''''''''''''''''''''''''''''''''''''''''''''''''''''FRAMES HTML
sFileName = RemoveInvalidChars(sTitle) & ".html"
sUpHeight = Trim$(txtUpperHeight.Text)
If Val(sUpHeight) = 0 Then sUpHeight = "60"
sTemp = Replace$(sTemp, "%TITLE%", sTitle, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%ROWHEIGHT%", sUpHeight, 1, 1, vbTextCompare)
WriteFile OutputFolder & sFileName, sTemp

''''''''''''''''''''''''''''''''''''''''''''''''''''''LYRICS HTML
sTemp = LoadFile(AppPath() & "templates\lyrics.htm")
sTemp = Replace$(sTemp, "%FACE%", LyricsFont.Face, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%SIZE%", LyricsFont.Size, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%COLOR%", LyricsFont.Color, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%LEFTMARGIN%", LyricsMargins.LeftMargin, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%RIGHTMARGIN%", LyricsMargins.RightMargin, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%TOPMARGIN%", LyricsMargins.TopMargin, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%BGCOLOR%", LyricsBGColor, 1, 2, vbTextCompare)

If IsHTML(txtLyrics.Text) Then
    sLyrics = txtLyrics.Text
Else
    sLyrics = AddBR(txtLyrics.Text)
End If
sBIOpen = "": sBIClose = ""
If LyricsFont.Bold Then sBIOpen = "<B>": sBIClose = "</B>"
If LyricsFont.Italic Then sBIOpen = sBIOpen & "<I>": sBIClose = "</I>" & sBIClose
If Len(sBIOpen) > 0 Then
    sBIOpen = vbCrLf & sBIOpen & vbCrLf
    sBIClose = vbCrLf & sBIClose & vbCrLf
    sLyrics = sBIOpen & sLyrics & sBIClose
End If
If chkCenter.Value = vbChecked Then sLyrics = "<CENTER>" & vbCrLf & sLyrics & vbCrLf & "</CENTER>"
sTemp = Replace$(sTemp, "%LYRICS%", sLyrics, 1, 1, vbTextCompare)
WriteFile OutputFolder & "files\lyrics.html", sTemp
''''''''''''''''''''''''''''''''''''''''''''''''''''''MEDIA PLAYER HTML
sBIOpen = "": sBIClose = ""
If ArtistFont.Bold Then sBIOpen = "<B>": sBIClose = "</B>"
If ArtistFont.Italic Then sBIOpen = sBIOpen & "<I>": sBIClose = "</I>" & sBIClose
If Len(sBIOpen) > 0 Then
    ArtistName = sBIOpen & ArtistName & sBIClose
End If
sBIOpen = "": sBIClose = ""
If SongFont.Bold Then sBIOpen = "<B>": sBIClose = "</B>"
If SongFont.Italic Then sBIOpen = sBIOpen & "<I>": sBIClose = "</I>" & sBIClose
If Len(sBIOpen) > 0 Then
    SongTitle = sBIOpen & SongTitle & sBIClose
End If

sTemp = LoadFile(AppPath() & "templates\up.htm")
sTemp = Replace$(sTemp, "%LEFTMARGIN%", UpMargins.LeftMargin, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%RIGHTMARGIN%", UpMargins.RightMargin, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%TOPMARGIN%", UpMargins.TopMargin, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%BGCOLOR%", UpBGColor, 1, 2, vbTextCompare)
sTemp = Replace$(sTemp, "%SONG%", SongTitle, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%ARTIST%", ArtistName, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%SONGFILE%", SongFile, 1, 1, vbTextCompare)

sTemp = Replace$(sTemp, "%SONGFACE%", SongFont.Face, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%SONGSIZE%", SongFont.Size, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%SONGCOLOR%", SongFont.Color, 1, 1, vbTextCompare)

sTemp = Replace$(sTemp, "%ARTISTFACE%", ArtistFont.Face, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%ARTISTSIZE%", ArtistFont.Size, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%ARTISTCOLOR%", ArtistFont.Color, 1, 1, vbTextCompare)

MediaCtlWidth = Trim$(txtWidth.Text)
MediaCtlHeight = Trim$(txtHeight.Text)
If MediaCtlWidth = "" Then MediaCtlWidth = "150"
If MediaCtlHeight = "" Then MediaCtlHeight = "45"

sTemp = Replace$(sTemp, "%MEDIAWIDTH%", MediaCtlWidth, 1, 1, vbTextCompare)
sTemp = Replace$(sTemp, "%MEDIAHEIGHT%", MediaCtlHeight, 1, 1, vbTextCompare)
WriteFile OutputFolder & "files\up.html", sTemp
If FileExists(ImageFile) Then
    frmSplitPic.SplitPic ImageFile, OutputFolder & "files\up.bmp", OutputFolder & "files\dn.bmp", Val(UpperFrameHeight)
End If
Unload frmSplitPic

If Me.chkAutoPreview.Value = vbChecked Then
'     RunFile OutputFolder & sFileName, OutputFolder, SW_SHOWMAXIMIZED
    frmWeb.web.Navigate OutputFolder & sFileName
    frmWeb.Show vbModal
    If frmWeb.Kill = True Then
        On Error Resume Next
        Kill OutputFolder & sFileName
        Kill OutputFolder & "files\lyrics.html"
        Kill OutputFolder & "files\dn.bmp"
        Kill OutputFolder & "files\up.html"
        Kill OutputFolder & "files\up.bmp"
        Kill NewSongFile
        RmDir OutputFolder & "files"
        On Error GoTo 0
    End If
    Unload frmWeb
Else
    MsgBox "Done!", vbInformation, "Message"
End If

End Sub
Private Sub Label17_Click()

End Sub

Private Sub Command9_Click()

Load frmSettings
frmSettings.ClrBtn1.Color = HexColorToVBHex(Replace$(UpBGColor, "#", ""))
frmSettings.txtColor.Text = "#" & LongColorToHex(frmSettings.ClrBtn1.Color)
With UpMargins
frmSettings.txtLeft.Text = .LeftMargin
    frmSettings.txtRight.Text = .RightMargin
    frmSettings.txtTop.Text = .TopMargin
End With

frmSettings.Show vbModal
If frmSettings.txtColor = "" Then Exit Sub 'Was Canceled

UpBGColor = Trim$(frmSettings.txtColor)
With UpMargins
    .LeftMargin = Trim$(frmSettings.txtLeft.Text)
    .RightMargin = Trim$(frmSettings.txtRight.Text)
    .TopMargin = Trim$(frmSettings.txtTop.Text)
    If (.LeftMargin = "" Or .RightMargin = "" Or .TopMargin = "") Then
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
    End If
End With
If UpBGColor = "" Then UpBGColor = "#FFFFFF" 'white

Unload frmSettings

End Sub

Private Sub Form_Load()

txtWidth.Text = MediaCtlWidth
txtHeight.Text = MediaCtlHeight
txtUpperHeight.Text = UpperFrameHeight

If GetScreenBPP() < 24 Then
    MsgBox "IMPORTANT:" & vbCrLf & "Current Screen Color Depth is less than 24Bit" & vbCrLf & "For better results, switch to 24Bit Mode (TrueColor)", vbExclamation, "Message"
    Command2.Enabled = False
    txtBackgroundFile.Enabled = False
    Label2.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub picLayout_Click(index As Integer)
Dim idx As Integer

picLayout(index).BorderStyle = 1

For idx = 0 To picLayout.Count - 1
    If idx <> index Then
        picLayout(idx).BorderStyle = 0
    End If
Next

End Sub


Private Sub txtArtist_DblClick()
SelectAllText txtArtist

End Sub


Private Sub txtBackgroundFile_DblClick()
SelectAllText txtBackgroundFile

End Sub


Private Sub txtBackgroundFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
txtBackgroundFile = Data.Files(1)

End Sub

Private Sub txtFolder_DblClick()
SelectAllText txtFolder

End Sub


Private Sub txtHeight_Click()
SelectAllText txtHeight

End Sub


Private Sub txtLyrics_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyA Then
    If Shift = vbCtrlMask Then
        txtLyrics.SelStart = 0
        txtLyrics.SelLength = Len(txtLyrics.Text)
    End If
End If

End Sub

Private Sub txtLyrics_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
txtLyrics.Text = LoadFile(Data.Files(1))
''''''''''''''''''''''''''''''''''''''''''

If Err Then MsgBox "Couldn't Load File" + vbCrLf + UCase(cdlg.FileName) + vbCrLf + Error + "  " + Str(Err), vbCritical, "Error"

If IsHTML(txtLyrics.Text) Then
    txtLyrics.Text = Html2Text(txtLyrics.Text)
End If

End Sub

Private Sub txtSongTitle_DblClick()

SelectAllText txtSongTitle

End Sub


Private Sub txtSoundFile_DblClick()
SelectAllText txtSoundFile
End Sub


Private Sub txtSoundFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next
txtSoundFile.Text = Data.Files(1)
'''''''''''''''''''''''''''''''''''''''''''''

Dim vArray As Variant
ReDim sArray(0 To 1) As String

vArray = Read_ID3Tag(txtSoundFile.Text)
If IsArray(vArray) Then
    txtSongTitle.Text = vArray(1)
    txtArtist.Text = vArray(2)
Else
    'now try to GUESS the info from FileName:
    sArray = HeuristicSongName(txtSoundFile.Text)
    txtArtist.Text = sArray(0)
    txtSongTitle.Text = sArray(1)
End If

End Sub

Private Sub txtUpperHeight_Click()
SelectAllText txtUpperHeight

End Sub


Private Sub txtWidth_Click()
SelectAllText txtWidth

End Sub


