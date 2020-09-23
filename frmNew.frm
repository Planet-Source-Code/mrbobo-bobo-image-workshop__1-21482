VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Image"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameImages 
      Caption         =   "FrameImages"
      Height          =   2175
      Left            =   840
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
      Begin VB.PictureBox FillPic 
         Height          =   540
         Left            =   360
         Picture         =   "frmNew.frx":0442
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000016&
      Height          =   315
      ItemData        =   "frmNew.frx":074C
      Left            =   1800
      List            =   "frmNew.frx":0762
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000016&
      Caption         =   "OK"
      Height          =   330
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Size in Pixels"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.VScrollBar VSPXw 
         Height          =   220
         Left            =   1320
         TabIndex        =   2
         Top             =   390
         Width           =   220
      End
      Begin VB.VScrollBar VSPXh 
         Height          =   220
         Left            =   3240
         TabIndex        =   1
         Top             =   390
         Width           =   220
      End
      Begin VB.TextBox txtPXw 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtPXh 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Width"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "x  Height"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Background Color"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'just interface stuff here the action happens
'back on frmMain
Dim OKnew As Boolean
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
If Val(txtPXh.Text) = 0 Or Val(txtPXw.Text) = 0 Then
    MsgBox "Invalid size. Please enter a size between 1 and 32767 pixels."
    Exit Sub
End If
Newcancel = False
OKnew = True
NewHeight = (Val(txtPXh.Text))
NewWidth = (Val(txtPXw.Text))
Select Case Combo1.ListIndex
    Case 0
        NewBGcol = vbWhite
    Case 1
        NewBGcol = vbBlack
    Case 2
        NewBGcol = vbRed
    Case 3
        NewBGcol = vbGreen
    Case 4
        NewBGcol = vbBlue
    Case 5
        NewBGcol = frmMain.TheColor.BackColor
End Select
CurBGindex = Combo1.ListIndex
If NewHeight > 32000 Or NewWidth > 32000 Then
    MsgBox "This is much too big !", vbCritical
    Exit Sub
End If
Unload Me
End Sub
Private Sub Form_Load()
OKnew = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If OKnew = False Then Newcancel = True
End Sub
Private Sub txtPXh_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtPXh_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtPXh.Text) > VSPXh.Max Then txtPXh.Text = TrimVoid(Str(VSPXh.Max))
VSPXh.Value = VSPXh.Max - Val(txtPXh.Text)
End Sub
Private Sub txtPXw_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtPXw_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtPXw.Text) > VSPXw.Max Then txtPXw.Text = TrimVoid(Str(VSPXw.Max))
VSPXw.Value = VSPXw.Max - Val(txtPXw.Text)
End Sub
Private Sub VSPXh_Change()
txtPXh.Text = TrimVoid(Str(VSPXh.Max - VSPXh.Value))
End Sub
Private Sub VSPXw_Change()
txtPXw.Text = TrimVoid(Str(VSPXw.Max - VSPXw.Value))
End Sub
Private Sub VSPXh_scroll()
txtPXh.Text = TrimVoid(Str(VSPXh.Max - VSPXh.Value))
End Sub
Private Sub VSPXw_scroll()
txtPXw.Text = TrimVoid(Str(VSPXw.Max - VSPXw.Value))
End Sub

