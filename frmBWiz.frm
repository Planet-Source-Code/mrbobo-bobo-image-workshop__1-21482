VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBWiz 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Frames and Edges"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check3 
      Caption         =   "Use Current Color"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdColorsel 
      BackColor       =   &H80000016&
      Caption         =   "Select Color"
      Height          =   330
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Add inner line"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Add outline"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   2640
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox BordList 
      BackColor       =   &H80000016&
      Height          =   1620
      ItemData        =   "frmBWiz.frx":0000
      Left            =   240
      List            =   "frmBWiz.frx":0019
      TabIndex        =   11
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H80000016&
      Caption         =   "Apply"
      Height          =   330
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.PictureBox PicDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2535
      Left            =   2160
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox PicSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2775
      Left            =   1200
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "View Sample"
      Height          =   330
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin MSComctlLib.Slider SL1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   5
      TickFrequency   =   10
      Value           =   5
   End
   Begin MSComctlLib.Slider SL3 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   200
      SelStart        =   10
      TickFrequency   =   10
      Value           =   10
   End
   Begin MSComctlLib.Slider SL4 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   200
      SelStart        =   10
      TickFrequency   =   20
      Value           =   10
   End
   Begin MSComctlLib.Slider SL2 
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   5
      TickFrequency   =   10
      Value           =   5
   End
   Begin VB.Label Lb2 
      Alignment       =   2  'Center
      Caption         =   "FRAME  WIDTH"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   360
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Lb4 
      Alignment       =   2  'Center
      Caption         =   "SHADOW  INTENSITY"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   360
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image imgSample 
      Height          =   855
      Left            =   2640
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape ShBG 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   2520
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Style"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Lb3 
      Alignment       =   2  'Center
      Caption         =   "HIGHLIGHT  INTENSITY"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Lb1 
      Alignment       =   2  'Center
      Caption         =   "BORDER  WIDTH"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "frmBWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sfilename As String
Public applied As Boolean
Dim ImgRatio As Double
Private Sub Check1_Click()
If Check1.Value = 1 Then
    outline = True
Else
    outline = False
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then
    inline = True
Else
    inline = False
End If
End Sub
Private Sub Check3_Click()
If Check3.Value = 1 Then
    cmdColorsel.BackColor = frmMain.TheColor.BackColor
    chBGcolor = frmMain.TheColor.BackColor
Else
    cmdColorsel.BackColor = chBGcolor
End If
End Sub
Private Sub cmdApply_Click()
Select Case BordList.ListIndex
    Case 0
        borderwidth = SL1.Value
        curborderlevel2 = SL2.Value
        curborderlevel3 = SL3.Value
    Case 1
        borderwidth = SL1.Value
        curborderlevel2 = SL2.Value
        curborderlevel3 = SL3.Value
    Case 2
        borderwidth = SL1.Value
        framewidth = SL2.Value
    Case 3
        borderwidth = SL1.Value
        framewidth = SL2.Value
        chBGcolor = cmdColorsel.BackColor
    Case 4
        borderwidth = SL1.Value
        chBGcolor = cmdColorsel.BackColor
    Case 5
        borderwidth = SL1.Value
    Case 6
        borderwidth = SL1.Value
End Select
BWcancel = False
applied = True
Unload Me
End Sub
Private Sub cmdCancel_Click()
BWcancel = True
Unload Me
End Sub
Public Sub Startup(picco As PictureBox)
Dim ThumbWidth As Integer
Dim ThumbHeight As Integer
Dim ImgWidth As Integer
Dim ImgHeight As Integer
ImgWidth = picco.ScaleWidth * Screen.TwipsPerPixelX
ImgHeight = picco.ScaleHeight * Screen.TwipsPerPixelY
If ImgWidth > (ShBG.Width - 120) Or ImgHeight > (ShBG.Height - 120) Then
    If ImgWidth > ImgHeight Then
        ImgRatio = ImgWidth / (ShBG.Width - 30)
        ThumbWidth = ShBG.Width - 30
        ThumbHeight = (ImgHeight - 30) / ImgRatio
    ElseIf ImgWidth < ImgHeight Then
        ImgRatio = ImgHeight / (ShBG.Height - 30)
        ThumbWidth = (ShBG.Width - 30) / ImgRatio
        ThumbHeight = (ShBG.Height - 30)
    ElseIf ImgWidth = ImgHeight Then
        ImgRatio = ImgHeight / (ShBG.Height - 30)
        ThumbWidth = ShBG.Width - 30
        ThumbHeight = ShBG.Width - 30
    End If
Else
    If ImgWidth > ImgHeight Then
        ImgRatio = ImgWidth / (ShBG.Width - 30)
    Else
        ImgRatio = ImgHeight / (ShBG.Height - 30)
    End If
    ThumbWidth = ImgWidth / ImgRatio
    ThumbHeight = ImgHeight / ImgRatio
End If
picsource.Height = ThumbHeight
picsource.Width = ThumbWidth
PicDest.Height = ThumbHeight
PicDest.Width = ThumbWidth
imgSample.Height = ThumbHeight
imgSample.Width = ThumbWidth
picsource.BackColor = picco.BackColor
PicDest.BackColor = picco.BackColor
If picco.Picture <> 0 Then
picsource.PaintPicture picco.Picture, 0, 0, picsource.ScaleWidth, picsource.ScaleHeight, 0, 0, picco.ScaleWidth, picco.ScaleHeight, &HCC0020
End If
If GetTempFile("", "BI", 0, sfilename) Then SavePicture picsource.Image, sfilename
picsource.Picture = LoadPicture(sfilename)
PicDest.Picture = LoadPicture(sfilename)
imgSample.Picture = PicDest.Image
imgSample.Left = ShBG.Left + 15
If imgSample.Width < ShBG.Width - 2 Then
    imgSample.Left = ShBG.Left + (ShBG.Width - imgSample.Width - 2) / 2
End If
imgSample.Top = (ShBG.Top + (ShBG.Height / 2)) - (imgSample.Height / 2)
End Sub
Function FileExists(ByVal fileName As String) As Integer
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(fileName)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
                End
            End If
    End Select
End Function
Private Sub cmdColorsel_Click()
Dim NewColor As Long
NewColor = ShowColor
Check3.Value = 0
If NewColor <> -1 Then
    chBGcolor = NewColor
    cmdColorsel.BackColor = NewColor
Else
    cmdColorsel.BackColor = chBGcolor
End If
End Sub
Private Sub Command1_Click()
'After the user has selected the desired settings
'this is where we call the functions in the
'module ImgFiltMod
Dim tempvalue As Integer
LockWindowUpdate Me.Hwnd
frmMain.Pb.Visible = False
For x = 0 To BordList.ListCount - 1
    If BordList.Selected(x) = True Then
        If x = 0 Then
            PicDest.Picture = LoadPicture(sfilename)
            Butt1 picsource, PicDest, SL1.Value / ImgRatio, SL2.Value, SL3.Value, outline, inline, frmMain.Pb
        ElseIf x = 1 Then
            PicDest.Picture = LoadPicture(sfilename)
            Butt2 picsource, PicDest, SL1.Value / ImgRatio, SL2.Value, SL3.Value, outline, inline, frmMain.Pb
        ElseIf x = 2 Then
            PicDest.Picture = LoadPicture(sfilename)
            Frame3D picsource, PicDest, SL1.Value / ImgRatio, SL2.Value / ImgRatio, 12632256, frmMain.Pb
        ElseIf x = 3 Then
            PicDest.Picture = LoadPicture(sfilename)
            Frame3D picsource, PicDest, SL1.Value / ImgRatio, SL2.Value / ImgRatio, chBGcolor, frmMain.Pb
        ElseIf x = 4 Then
            PicDest.Picture = LoadPicture(sfilename)
            FlatBorder picsource, PicDest, SL1.Value / ImgRatio, chBGcolor, outline, inline, frmMain.Pb
        ElseIf x = 5 Then
            PicDest.Picture = LoadPicture(sfilename)
            ButtBW picsource, PicDest, SL1.Value / ImgRatio, SL2.Value, SL3.Value, outline, inline, True, frmMain.Pb
        ElseIf x = 6 Then
            PicDest.Picture = LoadPicture(sfilename)
            ButtBW picsource, PicDest, SL1.Value / ImgRatio, SL2.Value, SL3.Value, outline, inline, False, frmMain.Pb
        End If
        imgSample.Picture = PicDest.Image
        picsource.Picture = LoadPicture(sfilename)
        curborder = x
        Exit For
    End If
Next x
LockWindowUpdate 0
frmMain.Pb.Visible = True
End Sub
Private Sub BordList_Click()
Dim slval(0 To 3) As Integer
slval(0) = SL1.Max / SL1.Value
slval(1) = SL2.Max / SL2.Value
slval(2) = SL3.Max / SL3.Value
slval(3) = SL4.Max / SL4.Value
SL1.Value = SL1.Min
SL2.Value = SL2.Min
SL3.Value = SL3.Min
SL4.Value = SL4.Min
SL1.Visible = False
SL2.Visible = False
SL3.Visible = False
SL4.Visible = False
Lb1.Visible = False
Lb2.Visible = False
Lb3.Visible = False
Lb4.Visible = False
cmdColorsel.Visible = False
Check1.Visible = False
Check2.Visible = False
Check3.Visible = False
For x = 0 To BordList.ListCount - 1
    If BordList.Selected(x) = True Then
        If x = 0 Then
            SL1.Visible = True
            SL2.Visible = True
            SL3.Visible = True
            SL1.Max = 100
            SL2.Max = 100
            SL3.Max = 200
            Lb1.Visible = True
            Lb2.Visible = True
            Lb3.Visible = True
            Lb1.Caption = "BORDER  WIDTH"
            Lb2.Caption = "HIGHLIGHT  INTENSITY"
            Lb3.Caption = "SHADOW  INTENSITY"
            Check1.Visible = True
            Check2.Visible = True
        ElseIf x = 1 Then
            SL1.Visible = True
            SL2.Visible = True
            SL3.Visible = True
            SL1.Max = 100
            SL2.Max = 100
            SL3.Max = 200
            Lb1.Visible = True
            Lb2.Visible = True
            Lb3.Visible = True
            Lb1.Caption = "BORDER  WIDTH"
            Lb2.Caption = "HIGHLIGHT  INTENSITY"
            Lb3.Caption = "SHADOW  INTENSITY"
            Check1.Visible = True
            Check2.Visible = True
        ElseIf x = 2 Then
            SL1.Visible = True
            SL2.Visible = True
            SL1.Max = 100
            SL2.Max = 100
            Lb1.Visible = True
            Lb2.Visible = True
            Lb1.Caption = "BORDER  WIDTH"
            Lb2.Caption = "FRAME  WIDTH"
        ElseIf x = 3 Then
            SL1.Visible = True
            SL2.Visible = True
            SL1.Max = 100
            SL2.Max = 100
            Lb1.Visible = True
            Lb2.Visible = True
            Lb1.Caption = "BORDER  WIDTH"
            Lb2.Caption = "FRAME  WIDTH"
            cmdColorsel.Visible = True
            Check3.Visible = True
        ElseIf x = 4 Then
            SL1.Visible = True
            Lb1.Visible = True
            Lb1.Caption = "BORDER  WIDTH"
            cmdColorsel.Visible = True
            Check1.Visible = True
            Check2.Visible = True
            Check3.Visible = True
        ElseIf x = 5 Then
            SL1.Visible = True
            Lb1.Visible = True
            Lb1.Caption = "BORDER  WIDTH"
            Check1.Visible = True
            Check2.Visible = True
        ElseIf x = 6 Then
            SL1.Visible = True
            Lb1.Visible = True
            Lb1.Caption = "BORDER  WIDTH"
            Check1.Visible = True
            Check2.Visible = True
        End If
        imgSample.Picture = PicDest.Image
        If FileExists(sfilename) Then picsource.Picture = LoadPicture(sfilename)
        curborder = x
        Exit For
    End If
Next x
SL1.Value = SL1.Max / slval(0)
SL2.Value = SL2.Max / slval(1)
SL3.Value = SL3.Max / slval(2)
SL4.Value = SL4.Max / slval(3)
SL1.TickFrequency = SL1.Max / 10
SL2.TickFrequency = SL2.Max / 10
SL3.TickFrequency = SL3.Max / 10
SL4.TickFrequency = SL4.Max / 10
End Sub
Private Sub Form_Load()
If Check2.Value = 1 Then
    inline = True
Else
    inline = False
End If
If Check1.Value = 1 Then
    outline = True
Else
    outline = False
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FileExists(sfilename) Then Kill sfilename
End Sub
Private Sub Form_Unload(Cancel As Integer)
If applied = False Then BWcancel = True
End Sub

