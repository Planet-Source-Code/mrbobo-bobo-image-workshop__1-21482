VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmText 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Text"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5205
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4935
      Begin VB.ComboBox ComboFont 
         BackColor       =   &H80000016&
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox ComboSize 
         BackColor       =   &H80000016&
         Height          =   315
         ItemData        =   "frmText.frx":0442
         Left            =   3960
         List            =   "frmText.frx":04B5
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Width           =   255
      End
      Begin VB.PictureBox picFontcol 
         BackColor       =   &H80000007&
         Height          =   255
         Left            =   1200
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox chkStrikeTh 
         Caption         =   "Strikethrough"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkUnderlined 
         Caption         =   "Underline"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox ComboStyle 
         BackColor       =   &H80000016&
         Height          =   315
         ItemData        =   "frmText.frx":0549
         Left            =   2400
         List            =   "frmText.frx":0559
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Text Color"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Font"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Style"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Size"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox rtftext 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   -2147483626
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmText.frx":0583
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000016&
      Caption         =   "OK"
      Height          =   330
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   5280
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Text to insert:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   45
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Settings are applied to a label on the active form
'after the user positions this label the label text
'is applied to the Picture using the print function
Private Sub chkStrikeTh_Click()
lblDisplay.FontStrikethru = chkStrikeTh.Value
End Sub
Private Sub chkUnderlined_Click()
lblDisplay.FontUnderline = chkUnderlined.Value
End Sub
Private Sub ComboFont_Click()
lblDisplay.FontName = Screen.Fonts(ComboFont.ListIndex)
End Sub
Private Sub ComboSize_Click()
If ComboSize.ListIndex = 0 Then lblDisplay.FontSize = 2
If ComboSize.ListIndex > 1 Then lblDisplay.FontSize = (ComboSize.ListIndex + 1) * 2
End Sub
Private Sub ComboStyle_Click()
Select Case ComboStyle.ListIndex
    Case 0
        lblDisplay.FontBold = False
        lblDisplay.FontItalic = False
    Case 1
        lblDisplay.FontBold = True
        lblDisplay.FontItalic = False
    Case 2
        lblDisplay.FontBold = False
        lblDisplay.FontItalic = True
    Case 3
        lblDisplay.FontBold = True
        lblDisplay.FontItalic = True
End Select
lblDisplay.Refresh
End Sub
Private Sub Command1_Click()
On Error GoTo woops
frmMain.CommonDialog1.CancelError = True
frmMain.CommonDialog1.Flags = 5
frmMain.CommonDialog1.action = 3
picFontcol.BackColor = frmMain.CommonDialog1.Color
lblDisplay.ForeColor = frmMain.CommonDialog1.Color
woops:
End Sub
Private Sub Command2_Click()
With frmMain.ActiveForm.lbltext
.FontName = lblDisplay.FontName
.FontStrikethru = lblDisplay.FontStrikethru
.FontUnderline = lblDisplay.FontUnderline
.ForeColor = lblDisplay.ForeColor
.FontSize = lblDisplay.FontSize * frmMain.ActiveForm.factor
.FontBold = lblDisplay.FontBold
.FontItalic = lblDisplay.FontItalic
.Caption = lblDisplay.Caption + " "
If frmMain.ActiveForm.PicBG.Top < 0 Then
    .Top = -frmMain.ActiveForm.PicBG.Top
Else
    .Top = 0
End If
If frmMain.ActiveForm.PicBG.Left < 0 Then
    .Left = -frmMain.ActiveForm.PicBG.Left
Else
    .Left = 0
End If
.Visible = True
End With
Unload Me
End Sub
Private Sub Command3_Click()
Unload Me
End Sub
Private Sub Form_Load()
Dim font As Integer
For font = 0 To Screen.FontCount - 1
    ComboFont.AddItem Screen.Fonts(font)
Next font
ComboFont.ListIndex = 0
ComboStyle.ListIndex = 1
ComboSize.ListIndex = 3
lblDisplay.FontName = Screen.Fonts(ComboFont.ListIndex)
lblDisplay.FontStrikethru = chkStrikeTh.Value
lblDisplay.FontUnderline = chkUnderlined.Value
lblDisplay.ForeColor = picFontcol.BackColor
lblDisplay.FontSize = (ComboSize.ListIndex + 1) * 2
lblDisplay.FontBold = True
lblDisplay.FontItalic = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If finalclose = False Then
    Me.Visible = False
    Cancel = 1
End If
End Sub
Private Sub rtftext_Change()
lblDisplay.Caption = rtftext.Text
End Sub
