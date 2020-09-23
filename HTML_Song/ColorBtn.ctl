VERSION 5.00
Begin VB.UserControl ColorButton 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   PropertyPages   =   "COLORBTN.ctx":0000
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
   ToolboxBitmap   =   "COLORBTN.ctx":0014
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   360
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   0
      Top             =   135
      Width           =   2805
   End
End
Attribute VB_Name = "ColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Add to a form_load:
' ColorButton1.Color = vbGreen
' ColorButton1.hWndOwner = Me.hWnd  // Or 0 for no owner
' it supports only 1 event "Click" where you can use code like:
' Me.BackColor = ColorButton1.Color
' Returns -1 if Canceled

Public Event Click()

Private ml_Color As Long
Private ml_hWndOwner As Long

''''''''''''''''''''''''' API ''''''''''''''''''''''''''''

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As CHOOSE_COLOR_FLAGS
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (lpChoosecolor As ChooseColorStruct) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Enum CHOOSE_COLOR_FLAGS
  CC_RGBINIT = &H1&
  CC_FULLOPEN = &H2&
  CC_PREVENTFULLOPEN = &H4&
  CC_SHOWHELP = &H8&
  CC_ENABLEHOOK = &H10&
  CC_ENABLETEMPLATE = &H20&
  CC_ENABLETEMPLATEHANDLE = &H40&
  CC_SOLIDCOLOR = &H80&
  CC_ANYCOLOR = &H100&
End Enum
  



Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_FLAT = &H4000      ' For flat rather than 3D borders
Private Const BF_LEFT = &H1
Private Const BF_MONO = &H8000      ' For monochrome borders.
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Private Const EDGE_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER
Private Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM


Public Function ShowColorDlg(Owner_hWnd As Long, DefaultColor As Long) As Long
  ' call the choosecolor dialog
  
  Dim lpChoosecolor As ChooseColorStruct
  Dim aColorRef(0 To 15) As Long
  
  With lpChoosecolor
    .lStructSize = Len(lpChoosecolor)
    .hwndOwner = Owner_hWnd
    .rgbResult = DefaultColor
    .lpCustColors = VarPtr(aColorRef(0))
    .flags = CC_SOLIDCOLOR Or CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN
  End With

If ChooseColor(lpChoosecolor) Then
      ShowColorDlg = lpChoosecolor.rgbResult
Else
      ShowColorDlg = -1
End If

End Function


Public Property Get hwndOwner() As Long
       hwndOwner = ml_hWndOwner
End Property

Public Property Let hwndOwner(ByVal lNewValue As Long)
       ml_hWndOwner = lNewValue
End Property

Public Property Get Color() As Long
       Color = ml_Color
End Property

Public Property Let Color(ByVal lNewValue As Long)
       ml_Color = lNewValue
       If lNewValue <> -1 Then pic.BackColor = lNewValue
       Dim r As RECT
r.Left = 0
r.Top = 0
r.Right = pic.ScaleWidth
r.Bottom = pic.ScaleHeight
    

DrawEdge pic.hdc, r, EDGE_RAISED, BF_RECT
pic.Refresh

       
End Property

Private Sub pic_Click()
RaiseEvent Click

End Sub

Private Sub pic_GotFocus()

Dim r As RECT
r.Left = 3
r.Top = 3
r.Right = pic.ScaleWidth - 3
r.Bottom = pic.ScaleHeight - 3

DrawFocusRect pic.hdc, r
pic.Refresh

End Sub


Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
    pic_MouseDown vbLeftButton, 0, 0, 0
End If

End Sub


Private Sub pic_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then

    pic_MouseUp vbLeftButton, 0, 0, 0
End If

End Sub


Private Sub pic_LostFocus()
pic.Cls

Dim r As RECT

r.Left = 0
r.Top = 0
r.Right = pic.Width
r.Bottom = pic.Height


DrawEdge pic.hdc, r, EDGE_RAISED, BF_RECT
pic.Refresh

End Sub


Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim r As RECT

r.Left = 0
r.Top = 0
r.Right = pic.Width
r.Bottom = pic.Height
    

DrawEdge pic.hdc, r, EDGE_SUNKEN, BF_RECT
pic.Refresh

End Sub


Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lColor As Long



Dim r As RECT
r.Left = 0
r.Top = 0
r.Right = pic.Width
r.Bottom = pic.Height

DrawEdge pic.hdc, r, EDGE_RAISED, BF_RECT
pic.Refresh

lColor = ShowColorDlg(hwndOwner, pic.BackColor)
If lColor > -1 Then pic.BackColor = lColor
Color = lColor

DrawEdge pic.hdc, r, EDGE_RAISED, BF_RECT
pic.Refresh

pic_GotFocus

End Sub


Private Sub UserControl_Resize()
pic.Top = 0
pic.Left = 0
pic.Height = ScaleHeight
pic.Width = ScaleWidth

End Sub


