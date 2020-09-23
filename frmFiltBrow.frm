VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFiltBrow 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filters"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   Icon            =   "frmFiltBrow.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "View Sample"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin VB.PictureBox PicSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2775
      Left            =   4320
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox PicDest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.ListBox FiltList 
      BackColor       =   &H80000016&
      Height          =   1620
      ItemData        =   "frmFiltBrow.frx":0442
      Left            =   240
      List            =   "frmFiltBrow.frx":0464
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H80000016&
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblFiltLevel 
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Image imgSample 
      Height          =   855
      Left            =   2640
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Filter"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Shape ShBG 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   2520
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sample"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmFiltBrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sfilename As String
Dim applied As Boolean
Dim ImgRatio As Double
Private Sub cmdApply_Click()
curfilterlevel = Slider1.Value
applied = True
Unload Me
End Sub
Private Sub cmdCancel_Click()
FBcancel = True
Unload Me
End Sub
Public Sub Startup(pic As PictureBox)
Dim ThumbWidth As Integer
Dim ThumbHeight As Integer
Dim ImgWidth As Integer
Dim ImgHeight As Integer
ImgWidth = pic.ScaleWidth * Screen.TwipsPerPixelX
ImgHeight = pic.ScaleHeight * Screen.TwipsPerPixelY
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
picsource.PaintPicture pic.Picture, 0, 0, picsource.ScaleWidth, picsource.ScaleHeight, 0, 0, pic.ScaleWidth, pic.ScaleHeight, &HCC0020
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
Private Sub Command1_Click()
Dim tempvalue As Integer
LockWindowUpdate Me.Hwnd
frmMain.Pb.Visible = False
'After the user has selected the desired settings
'this is where we call the functions in the
'module ImgFiltMod
For x = 0 To FiltList.ListCount - 1
    If FiltList.Selected(x) = True Then
        If x = 0 Then
            MySharpen picsource, PicDest, frmMain.Pb, Slider1.Value
        ElseIf x = 1 Then
            MyBlur picsource, PicDest, frmMain.Pb, Slider1.Value
        ElseIf x = 2 Then
            MyDiffuse picsource, PicDest, frmMain.Pb, Slider1.Value
        ElseIf x = 3 Then
            MyGreyscale picsource, PicDest, frmMain.Pb
        ElseIf x = 4 Then
            MyInvert picsource, PicDest
        ElseIf x = 5 Then
            MyBrightness picsource, PicDest, Slider1.Value, frmMain.Pb
        ElseIf x = 6 Then
            MyBrightness picsource, PicDest, -Slider1.Value, frmMain.Pb
        ElseIf x = 7 Then
            MyOutline picsource, PicDest, frmMain.Pb, Slider1.Value
        ElseIf x = 8 Then
            MyEmboss picsource, PicDest, frmMain.Pb, Slider1.Value
        ElseIf x = 9 Then
            tempvalue = Slider1.Value * (picsource.ScaleWidth / frmMain.ActiveForm.picsource.ScaleWidth)
            If tempvalue < 1 Then tempvalue = 1
            MyPixelate picsource, PicDest, frmMain.Pb, tempvalue
        End If
        imgSample.Picture = PicDest.Image
        picsource.Picture = LoadPicture(sfilename)
        PicDest.Picture = LoadPicture(sfilename)
        curfilter = x
        Exit For
    End If
Next x
LockWindowUpdate 0
frmMain.Pb.Visible = True
End Sub
Private Sub FiltList_Click()
Slider1.Min = 1
Slider1.Value = Slider1.Min
For x = 0 To FiltList.ListCount - 1
    If FiltList.Selected(x) = True Then
        If x <> 3 Or x <> 4 Then
            lblFiltLevel.Visible = True
            Slider1.Visible = True
        End If
        cmdApply.Enabled = True
        Command1.Enabled = True
        If x = 0 Then
            lblFiltLevel = "Sharpness"
            Slider1.Max = 10
            Slider1.TickFrequency = 1
        ElseIf x = 1 Then
            lblFiltLevel = "Blurriness"
            Slider1.Max = 10
            Slider1.TickFrequency = 1
        ElseIf x = 2 Then
            lblFiltLevel = "Diffusion"
            Slider1.Max = 10
            Slider1.TickFrequency = 1
        ElseIf x = 3 Then
            lblFiltLevel.Visible = False
            Slider1.Visible = False
        ElseIf x = 4 Then
            lblFiltLevel.Visible = False
            Slider1.Visible = False
        ElseIf x = 5 Then
            lblFiltLevel = "Brightness"
            Slider1.Max = 100
            Slider1.Min = 10
            Slider1.TickFrequency = 9
        ElseIf x = 6 Then
            lblFiltLevel = "Darkness"
            Slider1.Max = 100
            Slider1.Min = 10
            Slider1.TickFrequency = 9
        ElseIf x = 7 Then
            lblFiltLevel = "Tolerance"
            Slider1.Max = 250
            Slider1.Min = 50
            Slider1.TickFrequency = 20
        ElseIf x = 8 Then
            lblFiltLevel = "Emboss"
            Slider1.Max = 200
            Slider1.Min = 100
            Slider1.TickFrequency = 10
        ElseIf x = 9 Then
            lblFiltLevel = "Pixels/Pixelate"
            Slider1.Min = 5
            Slider1.Max = 50
            Slider1.TickFrequency = 5
        End If
        imgSample.Picture = PicDest.Image
        picsource.Picture = LoadPicture(sfilename)
        curfilter = x
        Exit For
    End If
Next x
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If FileExists(sfilename) Then Kill sfilename
End Sub
Private Sub Form_Unload(Cancel As Integer)
If applied = False Then FBcancel = True
End Sub
