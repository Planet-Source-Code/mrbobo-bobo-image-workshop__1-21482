VERSION 5.00
Begin VB.Form frmAddBord 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Border - Width in Pixels"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000016&
      Height          =   315
      ItemData        =   "frmAddBord.frx":0000
      Left            =   1200
      List            =   "frmAddBord.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H80000016&
      Caption         =   "Apply"
      Height          =   330
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox ChSym 
      Caption         =   "Symmetrical"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.VScrollBar VSTop 
      Height          =   220
      Left            =   1440
      Max             =   400
      TabIndex        =   6
      Top             =   645
      Value           =   2
      Width           =   220
   End
   Begin VB.VScrollBar VSBottom 
      Height          =   220
      Left            =   1440
      Max             =   400
      TabIndex        =   4
      Top             =   1125
      Value           =   2
      Width           =   220
   End
   Begin VB.VScrollBar VSRight 
      Height          =   220
      Left            =   1440
      Max             =   400
      TabIndex        =   2
      Top             =   2085
      Value           =   2
      Width           =   220
   End
   Begin VB.VScrollBar VSLeft 
      Height          =   220
      Left            =   1440
      Max             =   400
      TabIndex        =   0
      Top             =   1605
      Value           =   2
      Width           =   220
   End
   Begin VB.TextBox txtLeft 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "1"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtRight 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "1"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtTop 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Text            =   "1"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtBottom 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Border Color :"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Right"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Left"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1620
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Bottom"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Top"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   660
      Width           =   495
   End
End
Attribute VB_Name = "frmAddBord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dontupdate As Boolean

Private Sub Form_Load()
Combo1.ListIndex = 0
VSTop.Value = 399
End Sub

Private Sub txtTop_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtTop.Text) > 400 Then txtTop.Text = Str(400)
VSTop.Value = VSTop.Max - Val(txtTop.Text)
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSTop.Value
        VSLeft.Value = VSTop.Value
        VSRight.Value = VSTop.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub txtBottom_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtBottom.Text) > 400 Then txtBottom.Text = Str(400)
VSBottom.Value = VSBottom.Max - Val(txtBottom.Text)
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSTop.Value = VSBottom.Value
        VSLeft.Value = VSBottom.Value
        VSRight.Value = VSBottom.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub txtLeft_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtLeft.Text) > 400 Then txtLeft.Text = Str(400)
VSLeft.Value = VSLeft.Max - Val(txtLeft.Text)
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSLeft.Value
        VSTop.Value = VSLeft.Value
        VSRight.Value = VSLeft.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub txtRight_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtRight.Text) > 400 Then txtRight.Text = Str(400)
VSRight.Value = VSRight.Max - Val(txtRight.Text)
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSRight.Value
        VSLeft.Value = VSRight.Value
        VSTop.Value = VSRight.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub txtTop_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtBottom_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtLeft_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtRight_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub cmdApply_Click()
Dim bCol As Long
Select Case Combo1.ListIndex
    Case 0
        bCol = vbWhite
    Case 1
        bCol = vbBlack
    Case 2
        bCol = vbRed
    Case 3
        bCol = vbGreen
    Case 4
        bCol = vbBlue
    Case 5
        bCol = frmMain.TheColor.BackColor
End Select
'Everything else on this form just handles the
'forms controls. this is where the action occurs
LockWindowUpdate frmMain.ActiveForm.Hwnd
frmMain.ActiveForm.SelectLeft = Val(txtLeft.Text)
frmMain.ActiveForm.SelectTop = Val(txtTop.Text)
frmMain.ActiveForm.Selectwidth = frmMain.ActiveForm.PicMerge.Width
frmMain.ActiveForm.Selectheight = frmMain.ActiveForm.PicMerge.Height
frmMain.ActiveForm.picsource.Picture = LoadPicture()
frmMain.ActiveForm.picsource.Width = frmMain.ActiveForm.PicMerge.Width
frmMain.ActiveForm.picsource.Height = frmMain.ActiveForm.PicMerge.Height
frmMain.ActiveForm.picsource.Picture = frmMain.ActiveForm.PicMerge.Image
frmMain.ActiveForm.PicMerge.Picture = LoadPicture()
frmMain.ActiveForm.PicMerge.Width = (frmMain.ActiveForm.PicMerge.ScaleWidth + Val(txtLeft.Text) + Val(txtRight.Text) + 2)
frmMain.ActiveForm.PicMerge.Height = (frmMain.ActiveForm.PicMerge.ScaleHeight + Val(txtTop.Text) + Val(txtBottom.Text) + 2)
frmMain.ActiveForm.Pic1BU.Picture = LoadPicture()
frmMain.ActiveForm.Pic1BU.Width = frmMain.ActiveForm.PicMerge.Width
frmMain.ActiveForm.Pic1BU.Height = frmMain.ActiveForm.PicMerge.Height
frmMain.ActiveForm.Pic1BU.BackColor = bCol
StretchBlt frmMain.ActiveForm.Pic1BU.hdc, frmMain.ActiveForm.SelectLeft, frmMain.ActiveForm.SelectTop, frmMain.ActiveForm.Selectwidth, frmMain.ActiveForm.Selectheight, frmMain.ActiveForm.picsource.hdc, 0, 0, frmMain.ActiveForm.Selectwidth, frmMain.ActiveForm.Selectheight, SRCCOPY
frmMain.ActiveForm.PicMerge.Picture = frmMain.ActiveForm.Pic1BU.Image
frmMain.ActiveForm.PicBG.Width = frmMain.ActiveForm.PicMerge.Width * frmMain.ActiveForm.factor
frmMain.ActiveForm.PicBG.Height = frmMain.ActiveForm.PicMerge.Height * frmMain.ActiveForm.factor
frmMain.ActiveForm.Image1.Width = frmMain.ActiveForm.PicMerge.ScaleWidth * frmMain.ActiveForm.factor
frmMain.ActiveForm.Image1.Height = frmMain.ActiveForm.PicMerge.ScaleHeight * frmMain.ActiveForm.factor
frmMain.ActiveForm.Image1.Picture = frmMain.ActiveForm.PicMerge.Image
frmMain.ActiveForm.Backup
LockWindowUpdate 0
Unload Me
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub VSTop_Scroll()
txtTop.Text = TrimVoid(Str(VSTop.Max - VSTop.Value))
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSTop.Value
        VSLeft.Value = VSTop.Value
        VSRight.Value = VSTop.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub VSTop_Change()
txtTop.Text = TrimVoid(Str(VSTop.Max - VSTop.Value))
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSTop.Value
        VSLeft.Value = VSTop.Value
        VSRight.Value = VSTop.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub VSBottom_Scroll()
txtBottom.Text = TrimVoid(Str(VSBottom.Max - VSBottom.Value))
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSTop.Value = VSBottom.Value
        VSLeft.Value = VSBottom.Value
        VSRight.Value = VSBottom.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub VSBottom_Change()
txtBottom.Text = TrimVoid(Str(VSBottom.Max - VSBottom.Value))
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSTop.Value = VSBottom.Value
        VSLeft.Value = VSBottom.Value
        VSRight.Value = VSBottom.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub VSLeft_Scroll()
txtLeft.Text = TrimVoid(Str(VSLeft.Max - VSLeft.Value))
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSLeft.Value
        VSTop.Value = VSLeft.Value
        VSRight.Value = VSLeft.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub VSLeft_Change()
txtLeft.Text = TrimVoid(Str(VSLeft.Max - VSLeft.Value))
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSLeft.Value
        VSTop.Value = VSLeft.Value
        VSRight.Value = VSLeft.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub VSRight_Scroll()
txtRight.Text = TrimVoid(Str(VSRight.Max - VSRight.Value))
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSRight.Value
        VSLeft.Value = VSRight.Value
        VSTop.Value = VSRight.Value
        dontupdate = False
    End If
End If
End Sub
Private Sub VSRight_Change()
txtRight.Text = TrimVoid(Str(VSRight.Max - VSRight.Value))
If dontupdate = False Then
    If ChSym.Value = 1 Then
        dontupdate = True
        VSBottom.Value = VSRight.Value
        VSLeft.Value = VSRight.Value
        VSTop.Value = VSRight.Value
        dontupdate = False
    End If
End If
End Sub

