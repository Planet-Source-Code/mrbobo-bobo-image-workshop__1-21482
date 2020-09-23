VERSION 5.00
Begin VB.Form frmImage 
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "frmImage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   6750
   Begin VB.Timer TimerFS 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5400
      Top             =   3120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5520
      Top             =   2640
   End
   Begin VB.ListBox ListBackup 
      Height          =   255
      Left            =   5520
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox ListBUorder 
      Height          =   255
      Left            =   5520
      TabIndex        =   19
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox ListBUtemp 
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picsource1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   6000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   6360
      Top             =   2640
   End
   Begin VB.PictureBox PicLrule 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   5880
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox PicTrule 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox LeftRuler 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   12
      Top             =   375
      Visible         =   0   'False
      Width           =   495
      Begin VB.Line LineY 
         X1              =   0
         X2              =   33
         Y1              =   152
         Y2              =   152
      End
      Begin VB.Image ImgLrule 
         Height          =   1215
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox TopRuler 
      Align           =   1  'Align Top
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   450
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   6750
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00EAFFFF&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   495
         TabIndex        =   9
         Top             =   0
         Width           =   495
         Begin VB.Label lbly 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   140
            Width           =   360
         End
         Begin VB.Label lblx 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   360
         End
      End
      Begin VB.Line LineX 
         X1              =   160
         X2              =   160
         Y1              =   0
         Y2              =   24
      End
      Begin VB.Image ImgTrule 
         Height          =   375
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5880
      Top             =   2640
   End
   Begin VB.PictureBox PicFrame 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   0
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.PictureBox PicMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   2520
         ScaleHeight     =   225
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   409
         TabIndex        =   24
         Top             =   3720
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.PictureBox Pic1BU 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   2775
         Left            =   9000
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   400
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   6060
      End
      Begin VB.PictureBox PicLay 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   4500
         Index           =   0
         Left            =   960
         ScaleHeight     =   300
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.PictureBox picsource 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   4500
         Left            =   1560
         ScaleHeight     =   300
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   316
         TabIndex        =   15
         Top             =   3960
         Visible         =   0   'False
         Width           =   4740
      End
      Begin VB.VScrollBar VS 
         Height          =   615
         Left            =   3720
         TabIndex        =   7
         Top             =   3240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar HS 
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox piccorner 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   5
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox SelHolder 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H008894A0&
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   4320
         ScaleHeight     =   53
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.PictureBox PicMerge 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   4500
         Left            =   240
         ScaleHeight     =   300
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   3
         Top             =   4080
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.PictureBox PicBG 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   0
         ScaleHeight     =   193
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   273
         TabIndex        =   1
         Top             =   0
         Width           =   4095
         Begin VB.PictureBox PicFreeSelect 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H008894A0&
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   1200
            ScaleHeight     =   121
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   177
            TabIndex        =   22
            Top             =   720
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.PictureBox PicSelect 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   2400
            ScaleHeight     =   105
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   97
            TabIndex        =   2
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
            Begin VB.Shape PicSelectShape 
               BorderStyle     =   3  'Dot
               DrawMode        =   6  'Mask Pen Not
               Height          =   495
               Left            =   120
               Top             =   0
               Width           =   255
            End
            Begin VB.Image SelectImage 
               Height          =   1215
               Left            =   360
               Stretch         =   -1  'True
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Label lbltext 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   2280
            MousePointer    =   15  'Size All
            TabIndex        =   21
            Top             =   2400
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Shape ShSquare 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   615
            Left            =   1560
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Shape SelectShape 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Height          =   615
            Left            =   360
            Top             =   1680
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Line Line1 
            BorderStyle     =   3  'Dot
            DrawMode        =   6  'Mask Pen Not
            Visible         =   0   'False
            X1              =   80
            X2              =   168
            Y1              =   136
            Y2              =   136
         End
         Begin VB.Image Image1 
            Height          =   1215
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Image1 is what we show the user. What ever they do
'to image1 we apply to PicMerge. This lets you zoom in
'and out with ease, and makes handling layers possible
'The undo/redo functions use 3 listboxes - this should
'be changed to a series of arrays
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public factorlevel As Integer
Public factor As Double
Public curpoly As Boolean
Public sx As Integer
Public sy As Integer
Public sfilename As String
Public SelX As Long
Public SelY As Long
Public selecting As Boolean
Public SelectLeft As Double
Public SelectTop As Double
Public Selectwidth As Double
Public Selectheight As Double
Public curtranscolor As Long
Public curdrawcolor As Long
Public SelectpicX As Integer
Public SelectpicY As Integer
Public AFCurfile As String
Public alreadysaved As Boolean
Dim pasting As Boolean
Dim curpos As Boolean
Dim slashX As Integer
Dim slashY As Integer
Dim slashSX As Integer
Dim slashSY As Integer
Dim TextX As Integer
Dim TextY As Integer
Dim startX As Integer
Dim startY As Integer
Dim minX As Integer
Dim minY As Integer
Dim maxX As Integer
Dim maxY As Integer
Private Const RGN_DIFF = 4
Dim CurRgn As Long, TempRgn As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2


Private Sub Form_Activate()
SetLevel
End Sub
Private Sub Form_Load()
factorlevel = 15
factor = 1
If NoSizeonStart = False Then
    Me.Width = 4620
    Me.Height = 4980
    PicBG.Picture = LoadPicture()
    PicMerge.Width = 4500
    PicMerge.Height = 4500
    PicBG.Top = 0
Else
    Me.Width = 4620
    Me.Height = 4980
    PicMerge.Picture = LoadPicture()
    PicMerge.BackColor = NewBGcol
    PicMerge.Width = (NewWidth)
    PicMerge.Height = (NewHeight)
    PicBG.Width = (NewWidth)
    PicBG.Height = (NewHeight)
    PicBG.Top = 0
End If
Pic1BU.Width = PicMerge.Width
Pic1BU.Height = PicMerge.Height
Pic1BU.Picture = LoadPicture()
Form_Resize
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If frmPrefs.Check4.Value = 0 Then
    If alreadysaved = True Then
        alreadysaved = False
    Else
        CheckChanged
        If saveCancel = True Then
            saveCancel = False
            Cancel = 1
            Exit Sub
        End If
    End If
End If
For x = 0 To ListBackup.ListCount - 1
    If FileExists(ListBackup.List(x)) Then Kill ListBackup.List(x)
Next x
End Sub
Public Sub CheckChanged()
On Error Resume Next
Dim temp As String, temp1 As String, sfile As String
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer
If Val(ListBUorder.List(ListBUorder.ListIndex)) <> 0 Then
    DialogType = vbYesNoCancel
    DialogTitle = "Bobo Enterprises"
    DialogMsg = "Do you you wish save chamges to " + Me.Caption + " ?"
    Response = MsgBox(DialogMsg, DialogType, DialogTitle)
        Select Case Response
            Case vbYes
                frmMain.MySaveAs
                If saveCancel = True Then Exit Sub
                saveCancel = False
            Case vbNo
                saveCancel = False
            Case vbCancel
                saveCancel = True
                Exit Sub
        End Select
End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
If RulersVis = False Then
    PicFrame.Top = 0
    PicFrame.Left = 0
    PicFrame.Width = Me.Width - 120
    PicFrame.Height = Me.Height - 420
    PicBG.Top = 0
    Image1.Top = 0
Else
    PicFrame.Top = TopRuler.Height
    PicFrame.Left = LeftRuler.Width
    PicFrame.Width = Me.Width - LeftRuler.Width - 120
    PicFrame.Height = Me.Height - TopRuler.Height - 420 '510
    PicBG.Top = 0
    Image1.Top = 0
End If
Picsize
drawrulers
End Sub
Private Sub Form_Unload(Cancel As Integer)
ImageCount = ImageCount - 1
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
x = x / Screen.TwipsPerPixelX
y = y / Screen.TwipsPerPixelY
PicMerge.DrawWidth = frmMain.Slider2.Value
curtranscolor = PicMerge.BackColor
curdrawcolor = frmMain.TheColor.BackColor
If curdrawcolor = curtranscolor Then curdrawcolor = curtranscolor + 4
If Button = 1 Then
    If Curtool = 23 Then
        startVSval = (Image1.Height + 1) / (y + 1)
        startHSval = (Image1.Width + 1) / (x + 1)
        Zoom True, Image1, PicMerge
        drawrulers
        Picsize
        If VS.Visible = True Then VS.Value = Int(VS.Max / startVSval)
        If HS.Visible = True Then HS.Value = Int(HS.Max / startHSval)
        Exit Sub
    End If
If lbltext.Visible = True Then
    PicMerge.FontName = lbltext.FontName
    PicMerge.FontStrikethru = lbltext.FontStrikethru
    PicMerge.FontUnderline = lbltext.FontUnderline
    PicMerge.ForeColor = lbltext.ForeColor
    PicMerge.FontSize = lbltext.FontSize
    PicMerge.FontBold = lbltext.FontBold
    PicMerge.FontItalic = lbltext.FontItalic
    PicMerge.CurrentX = lbltext.Left
    PicMerge.CurrentY = lbltext.Top
    PicMerge.Print lbltext.Caption
    lbltext.Visible = False
    Exit Sub
End If
    If PicSelect.Visible = True Then
        SelectLeft = PicSelect.Left / factor
        SelectTop = PicSelect.Top / factor
        Selectwidth = PicSelect.Width / factor
        Selectheight = PicSelect.Height / factor
        PicSelect.Visible = False
        SelectShape.Visible = False
        StretchBlt PicMerge.hdc, SelectLeft, SelectTop, Selectwidth, Selectheight, SelHolder.hdc, 0, 0, Selectwidth, Selectheight, SRCCOPY
        Image1.Picture = PicMerge.Image
        frmMain.TB2.Buttons(5).Enabled = False
        frmMain.TB2.Buttons(6).Enabled = False
        frmMain.mnuEditCut.Enabled = False
        frmMain.mnuEditCopy.Enabled = False
        Backup
        selecting = False
        Masterpasting = False
    End If
    If PicFreeSelect.Visible = True Then
        TransparentBlt Pic1BU.hdc, Int(PicFreeSelect.Left / factor), Int(PicFreeSelect.Top / factor), Int(PicFreeSelect.ScaleWidth / factor), Int(PicFreeSelect.ScaleHeight / factor), SelHolder.hdc, 0, 0, SelHolder.ScaleWidth, SelHolder.ScaleHeight, 8950944
        PicMerge.Picture = Pic1BU.Image
        PicMerge.Refresh
        Image1.Picture = PicMerge.Image
        PicFreeSelect.Visible = False
        PicFreeSelect.Picture = LoadPicture()
        Backup
        pasting = True
        Masterpasting = False
        Exit Sub
    End If
    If Curtool = 13 Then
        selecting = True
        sx = Int(x / factor) * factor
        sy = Int(y / factor) * factor
        SelectShape.Left = Int(x / factor) * factor
        SelectShape.Top = Int(y / factor) * factor
        SelX = Int(x / factor) * factor
        SelY = Int(y / factor) * factor
        SelectShape.Width = 0
        SelectShape.Height = 0
        SelectShape.Visible = True
    End If
    If Curtool = 14 Then
        PicMask.Width = PicMerge.Width
        PicMask.Height = PicMerge.Height
        Pic1BU.Picture = PicMerge.Image
        minX = Int(x / factor)
        minY = Int(y / factor)
        maxX = Int(x / factor)
        maxY = Int(y / factor)
        PicMerge.DrawMode = 6
        PicMerge.DrawStyle = 2
        PicMerge.DrawWidth = 1
        PicMerge.Line (Int(x / factor), Int(y / factor))-(Int(x / factor), Int(y / factor)), vbBlack
        PicMask.Cls
        PicMask.BackColor = PicMerge.BackColor
        PicMask.DrawMode = 13
        PicMask.DrawStyle = 0
        PicMask.Line (Int(x / factor), Int(y / factor))-(Int(x / factor), Int(y / factor)), vbBlack
        sx = Int(x / factor)
        sy = Int(y / factor)
        startX = Int(x / factor)
        startY = Int(y / factor)
    End If
    If Curtool = 15 Then
         Select Case PenTip
            Case 0
                sx = Int(x / factor)
                sy = Int(y / factor)
                slashSX = 0
                slashSY = 0
                slashX = 0
                slashY = 0
                PicMerge.CurrentX = Int(x / factor)
                PicMerge.CurrentY = Int(y / factor)
                PicMerge.PSet (Int(x / factor), Int(y / factor)), curdrawcolor
                GoTo woops
            Case 1
                slashSX = -frmMain.Slider1.Value
                slashSY = -frmMain.Slider1.Value
                slashX = frmMain.Slider1.Value
                slashY = frmMain.Slider1.Value
            Case 2
                slashSX = -frmMain.Slider1.Value
                slashSY = frmMain.Slider1.Value
                slashX = frmMain.Slider1.Value
                slashY = -frmMain.Slider1.Value
            Case 3
                slashSX = frmMain.Slider1.Value
                slashSY = 0
                slashX = -frmMain.Slider1.Value
                slashY = 0
            Case 4
                slashSX = 0
                slashSY = frmMain.Slider1.Value
                slashX = 0
                slashY = -frmMain.Slider1.Value
        End Select
        PicMerge.Line ((x + slashSX) / factor, (y + slashSY) / factor)-((x + slashX) / factor, (y + slashY) / factor), curdrawcolor
        sx = Int(x / factor)
        sy = Int(y / factor)
    End If
    If Curtool = 16 Then
        sx = Int(x / factor)
        sy = Int(y / factor)
        PicMerge.CurrentX = Int(x / factor)
        PicMerge.CurrentY = Int(y / factor)
        PicMerge.PSet (Int(x / factor), Int(y / factor)), PicMerge.BackColor
    End If
    If Curtool = 17 Then
        Call Filling(PicMerge, PicMerge.Point(Int((x - 14) / factor), Int((y + 16) / factor)), 0, Int((x - 14) / factor), Int((y + 16) / factor))
    End If
    If Curtool = 18 Then
        sx = Int(x / factor)
        sy = Int(y / factor)
        Line1.X1 = x
        Line1.Y1 = y
        Line1.X2 = x
        Line1.Y2 = y
        Line1.Visible = True
    End If
    If Curtool = 19 Then
        If curpoly = True Then
        sx = Int(x / factor)
        sy = Int(y / factor)
            curpoly = False
        End If
            Line1.X1 = sx * factor
            Line1.Y1 = sy * factor
            Line1.X2 = x
            Line1.Y2 = y
        Line1.Visible = True
    End If
    If Curtool = 20 Then
            sx = Int(x / factor)
            sy = Int(y / factor)
        If Shapetype = 0 Or Shapetype = 1 Then
            ShSquare.Shape = 0
            ShSquare.borderwidth = 1
            ShSquare.BorderColor = vbBlack
            ShSquare.Visible = True
            ShSquare.Left = sx * factor
            ShSquare.Top = sy * factor
            ShSquare.Height = 0
            ShSquare.Width = 0
        Else
            Line1.X1 = sx * factor
            Line1.Y1 = sy * factor
            Line1.X2 = x
            Line1.Y2 = y
            Line1.Visible = True
        End If
    End If
    If Curtool = 21 Then
        Image1.Refresh
        picsource.Height = PicMerge.Height
        picsource.Width = PicMerge.Width
        picsource.Picture = PicMerge.Image
        PicMerge.AutoRedraw = False
        CWcancel = False
        Select Case frmMain.Combo2.ListIndex
        Case 0
            DrawLight PicMerge, picsource, picsource.hdc, Int(x * Screen.TwipsPerPixelX / factor), Int(y * Screen.TwipsPerPixelY / factor), frmMain.Slider4.Value, frmMain.Slider4.Value, frmMain.Slider4.Value, frmMain.Slider5.Value, frmMain.Slider6.Value
        Case 1
            DrawLight PicMerge, picsource, picsource.hdc, Int(x * Screen.TwipsPerPixelX / factor), Int(y * Screen.TwipsPerPixelY / factor), -frmMain.Slider4.Value, -frmMain.Slider4.Value, -frmMain.Slider4.Value, frmMain.Slider5.Value, frmMain.Slider6.Value
        Case 2
            DrawLight PicMerge, picsource, picsource.hdc, Int(x * Screen.TwipsPerPixelX / factor), Int(y * Screen.TwipsPerPixelY / factor), frmMain.Slider4.Value, -1, -1, frmMain.Slider5.Value, frmMain.Slider6.Value
        Case 3
            DrawLight PicMerge, picsource, picsource.hdc, Int(x * Screen.TwipsPerPixelX / factor), Int(y * Screen.TwipsPerPixelY / factor), -1, frmMain.Slider4.Value, -1, frmMain.Slider5.Value, frmMain.Slider6.Value
        Case 4
            DrawLight PicMerge, picsource, picsource.hdc, Int(x * Screen.TwipsPerPixelX / factor), Int(y * Screen.TwipsPerPixelY / factor), -1, -1, frmMain.Slider4.Value, frmMain.Slider5.Value, frmMain.Slider6.Value
        Case 5
            Dim rr As Long, gg As Long, bb As Long
            rr = (Val(frmMain.txtRed.Text) / 255) * frmMain.Slider4.Value
            gg = (Val(frmMain.txtGreen.Text) / 255) * frmMain.Slider4.Value
            bb = (Val(frmMain.txtBlue.Text) / 255) * frmMain.Slider4.Value
            DrawLight PicMerge, picsource, picsource.hdc, Int(x * Screen.TwipsPerPixelX / factor), Int(y * Screen.TwipsPerPixelY / factor), rr, gg, bb, frmMain.Slider5.Value, frmMain.Slider6.Value
          End Select
        Image1.Picture = PicMerge.Image
    End If
End If
If Button = 2 Then
    If Curtool = 23 Then
        startVSval = (Image1.Height + 1) / (y + 1)
        startHSval = (Image1.Width + 1) / (x + 1)
        Zoom False, Image1, PicMerge
        drawrulers
        Picsize
        If VS.Visible = True Then VS.Value = Int(VS.Max / startVSval)
        If HS.Visible = True Then HS.Value = Int(HS.Max / startHSval)
        Exit Sub
    End If
End If
Image1.Picture = PicMerge.Image
Image1.Refresh
woops:
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
x = Int(x / Screen.TwipsPerPixelX)
y = Int(y / Screen.TwipsPerPixelY)
LineX.X1 = (x + Image1.Left + LeftRuler.ScaleWidth) + PicBG.Left
LineX.X2 = LineX.X1
LineX.Y1 = 0
LineX.Y2 = 24
LineY.X1 = -LeftRuler.Width
LineY.X2 = 33
LineY.Y1 = (y + Image1.Top) + PicBG.Top
LineY.Y2 = LineY.Y1
If LineX.X1 < 0 Then LineX.X1 = 0
If LineY.Y1 < 0 Then LineY.Y1 = 0
If LineX.X2 < 0 Then LineX.X2 = 0
If LineY.Y2 < 0 Then LineY.Y2 = 0
lblx = Int(x / factor)
lbly = Int(y / factor)
If Curtool = 13 Then
    Image1.MousePointer = vbCrosshair
ElseIf Curtool = 22 Then
    Image1.MousePointer = 99
    Image1.MouseIcon = frmMain.EyePic.Picture
ElseIf Curtool = 17 Then
    Image1.MousePointer = 99
    Image1.MouseIcon = frmMain.FillPic.Picture
ElseIf Curtool = 23 Then
    Image1.MousePointer = 99
    Image1.MouseIcon = frmMain.MagPic.Picture
Else
    Image1.MousePointer = vbDefault
End If
If Curtool = 22 Then
    On Local Error Resume Next
    Dim R As Long
    R = GetPixel(PicMerge.hdc, (x - 16) / factor, (y + 14) / factor)
    frmMain.Picture6.BackColor = R
    If ReadHex = True Then frmMain.lblColor = HexRGB(frmMain.Picture6.BackColor)
    If ReadRgb = True Then frmMain.lblColor = MyRGB(frmMain.Picture6.BackColor)
    If ReadLong = True Then frmMain.lblColor = "LONG:" + Str(frmMain.Picture6.BackColor)
    Exit Sub
End If
If Button = 1 Then
If Curtool = 13 Or Curtool = 14 Then
    If x > HS.Width - 3 - PicBG.Left And x < HS.Width + 3 - PicBG.Left Then
        If HS.Value + 3 < HS.Max Then
            HS.Value = HS.Value + 3
        Else
            HS.Value = HS.Max
        End If
    End If
    If x > HS.Width + 3 - PicBG.Left Then
        If HS.Value + 10 < HS.Max Then
            HS.Value = HS.Value + 10
        Else
            HS.Value = HS.Max
        End If
    End If
    If x < HS.Left + 3 - PicBG.Left And x > HS.Left - 3 - PicBG.Left Then
        If HS.Value - 3 > -1 Then
            HS.Value = HS.Value - 3
        Else
            HS.Value = 0
        End If
    End If
    If x < HS.Left - 3 - PicBG.Left Then
        If HS.Value - 10 > -1 Then
            HS.Value = HS.Value - 10
        Else
            HS.Value = 0
        End If
    End If
    If y > VS.Height - 3 - PicBG.Top And y < VS.Height + 3 - PicBG.Top Then
        If VS.Value + 3 < VS.Max Then
            VS.Value = VS.Value + 3
        Else
            VS.Value = VS.Max
        End If
    End If
    If y > VS.Height + 3 - PicBG.Top Then
        If VS.Value + 10 < VS.Max Then
            VS.Value = VS.Value + 10
        Else
            VS.Value = VS.Max
        End If
    End If
    If y < VS.Top + 3 - PicBG.Top And y > VS.Top - 3 - PicBG.Top Then
        If VS.Value - 3 > -1 Then
            VS.Value = VS.Value - 3
        Else
            VS.Value = 0
        End If
    End If
    If y < VS.Top - 3 - PicBG.Top Then
        If VS.Value - 10 > -1 Then
            VS.Value = VS.Value - 10
        Else
            VS.Value = 0
        End If
    End If
End If
    If Curtool = 13 Then
        x = Int(x / factor) * factor
        y = Int(y / factor) * factor
        If x > SelX Then
            SelectShape.Width = (x - SelX)
            SelectShape.Left = sx
        Else
            SelectShape.Width = (SelX - x)
            SelectShape.Left = (Image1.Left + x)
        End If
        If y > SelY Then
            SelectShape.Height = (y - SelY)
            SelectShape.Top = sy
        Else
            SelectShape.Height = (SelY - y)
            SelectShape.Top = (Image1.Top + y)
        End If
            SelectShape.Visible = True
    End If
    If Curtool = 14 Then
        If pasting = True Then Exit Sub
        If Int(x / factor) > maxX Then maxX = Int(x / factor)
        If Int(x / factor) < minX Then minX = Int(x / factor)
        If Int(y / factor) > maxY Then maxY = Int(y / factor)
        If Int(y / factor) < minY Then minY = Int(y / factor)
        PicMerge.Line (sx, sy)-(Int(x / factor), Int(y / factor)), vbBlack
        PicMask.Line (sx, sy)-(Int(x / factor), Int(y / factor)), vbBlack
        sx = Int(x / factor)
        sy = Int(y / factor)
    End If
    If Curtool = 15 Then
        If PenTip = 0 Then
            PicMerge.CurrentX = Int(x / factor)
            PicMerge.CurrentY = Int(y / factor)
            PicMerge.PSet (Int(x / factor), Int(y / factor)), curdrawcolor
        End If
        PicMerge.Line ((sx + slashSX), (sy + slashSY))-(Int((x / factor + slashX)), Int((y / factor + slashY))), curdrawcolor
        sx = Int(x / factor)
        sy = Int(y / factor)
    End If
    If Curtool = 16 Then
        PicMerge.CurrentX = Int(x / factor)
        PicMerge.CurrentY = Int(y / factor)
        PicMerge.PSet (Int(x / factor), Int(y / factor)), PicMerge.BackColor
        PicMerge.Line (sx, sy)-(Int(x / factor), Int(y / factor)), PicMerge.BackColor
        sx = Int(x / factor)
        sy = Int(y / factor)
    End If
    If Curtool = 18 Or Curtool = 19 Then
        Line1.X2 = x
        Line1.Y2 = y
    End If
    If Curtool = 20 Then
    Dim rulesx As Integer, rulesy As Integer
        If Shapetype = 0 Then
            If x > sx * factor Then
                ShSquare.Left = sx * factor
                ShSquare.Width = (x - sx * factor)
            Else
                ShSquare.Left = x
                ShSquare.Width = (sx * factor - x)
            End If
            ShSquare.Height = ShSquare.Width
            If y > sy * factor Then
                ShSquare.Top = sy * factor
             Else
                ShSquare.Top = (sy * factor - ShSquare.Height)
            End If
        ElseIf Shapetype = 1 Then
            If x > sx * factor Then
                ShSquare.Left = sx * factor
                ShSquare.Width = (x - sx * factor)
            Else
                ShSquare.Left = x
                ShSquare.Width = (sx * factor - x)
            End If
            If y > sy * factor Then
                ShSquare.Top = sy * factor
                ShSquare.Height = (y - sy * factor)
             Else
                ShSquare.Top = y
                ShSquare.Height = (sy * factor - y)
            End If
        Else
            Line1.X2 = x
            Line1.Y2 = y
        End If
    End If
    If Curtool = 21 Then
        Image1.Picture = PicMerge.Image
        Image1.Refresh
        Image1_MouseDown Button, Shift, x, y
    End If
    If Curtool <> 17 Then
        Image1.Picture = PicMerge.Image
        Image1.Refresh
    End If
End If
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
x = Int(x / Screen.TwipsPerPixelX)
y = Int(y / Screen.TwipsPerPixelY)
If Button = 1 Then
    If Curtool = 13 Then
        On Error Resume Next
        If x = SelX And y = SelY Then Exit Sub
        freeselection = False
        dontusePicBU = False
        PicSelect.Left = SelectShape.Left
        PicSelect.Top = SelectShape.Top
        PicSelect.Height = SelectShape.Height
        PicSelect.Width = SelectShape.Width
        SelectLeft = (SelectShape.Left) / factor
        SelectTop = (SelectShape.Top) / factor
        Selectwidth = (SelectShape.Width) / factor
        Selectheight = (SelectShape.Height) / factor
        myinternalcopy PicMerge, SelHolder, PicSelectShape
        SelectImage.Left = 0
        SelectImage.Top = 0
        SelectImage.Width = PicSelect.Width
        SelectImage.Height = PicSelect.Height
        SelectImage.Picture = SelHolder.Image
        PicSelectShape.Left = 0
        PicSelectShape.Top = 0
        PicSelectShape.Width = PicSelect.Width
        PicSelectShape.Height = PicSelect.Height
        PicSelect.Visible = True
        SelectShape.Visible = False
        frmMain.TB2.Buttons(5).Enabled = True
        frmMain.TB2.Buttons(6).Enabled = True
        frmMain.mnuEditCut.Enabled = True
        frmMain.mnuEditCopy.Enabled = True
        selecting = False
        GoTo woops
    End If
    If Curtool = 14 Then
        If pasting = True Then
            pasting = False
            Exit Sub
        End If
        dontusePicBU = False
        PicMerge.Line (sx, sy)-(startX, startY), vbBlack
        PicMask.Line (sx, sy)-(startX, startY), vbBlack
        MakeSelection
        PicMerge.DrawMode = 13
        PicMerge.DrawStyle = 0
        frmMain.TB2.Buttons(5).Enabled = True
        frmMain.TB2.Buttons(6).Enabled = True
        frmMain.mnuEditCut.Enabled = True
        frmMain.mnuEditCopy.Enabled = True
        freeselection = True
        Exit Sub
End If
     If Curtool = 22 Then
        curdrawcolor = frmMain.Picture6.BackColor
        frmMain.TheColor.BackColor = curdrawcolor
        frmMain.LblCurCol.ForeColor = ContrastingColor(frmMain.TheColor.BackColor)
        frmMain.LblCurCol.Refresh
        frmMain.txtRGB frmMain.Picture6.BackColor
        If ReadHex = True Then frmMain.lblColor = HexRGB(frmMain.TheColor.BackColor)
        If ReadRgb = True Then frmMain.lblColor = MyRGB(frmMain.TheColor.BackColor)
        If ReadLong = True Then frmMain.lblColor = "LONG:" + Str(frmMain.TheColor.BackColor)
        If colslocked = False Then frmMain.loadselected
        
        GoTo woops
    End If
    If Curtool = 18 Then
        Line1.X2 = x
        Line1.Y2 = y
        Line1.Visible = False
        PicMerge.Line (sx, sy)-(Int(x / factor), Int(y / factor)), curdrawcolor
    End If
    If Curtool = 19 Then
        Line1.X2 = Int(x / factor)
        Line1.Y2 = Int(y / factor)
        Line1.Visible = False
        PicMerge.Line (sx, sy)-(Int(x / factor), Int(y / factor)), curdrawcolor
        sx = Int(x / factor)
        sy = Int(y / factor)
    End If
    If Curtool = 20 Then
        If frmMain.Option1.Value = True Then
            PicMerge.FillStyle = 1
            PicMerge.DrawWidth = frmMain.Slider3.Value
            If Shapetype = 0 Then
                ShSquare.Visible = False
                PicMerge.Line (ShSquare.Left / factor - Image1.Left / factor, ShSquare.Top / factor - Image1.Top / factor)-(ShSquare.Left / factor - Image1.Left / factor + ShSquare.Width / factor, ShSquare.Top / factor - Image1.Top / factor + ShSquare.Height / factor), curdrawcolor, B
            End If
            If Shapetype = 1 Then
                ShSquare.Visible = False
                PicMerge.Line (sx, sy)-((Int(x / factor)), (Int(y / factor))), curdrawcolor, B
            End If
            If Shapetype = 2 Then
                R = Sqr((sx - (Int(x / factor))) * (sx - (Int(x / factor))) + (sy - (Int(y / factor))) * (sy - (Int(y / factor))))
                PicMerge.Circle (sx, sy), R, curdrawcolor
                Line1.Visible = False
            End If
            If Shapetype = 3 Then
                PicMerge.ForeColor = curdrawcolor
                Call Ellipse(PicMerge.hdc, sx, sy, (Int(x / factor)), (Int(y / factor)))
                Line1.Visible = False
            End If
        Else
            PicMerge.DrawWidth = 1
            If Shapetype = 0 Then
                ShSquare.Visible = False
                PicMerge.Line (ShSquare.Left / factor - Image1.Left / factor, ShSquare.Top / factor - Image1.Top / factor)-(ShSquare.Left / factor - Image1.Left / factor + ShSquare.Width / factor, ShSquare.Top / factor - Image1.Top / factor + ShSquare.Height / factor), curdrawcolor, BF
            End If
            If Shapetype = 1 Then
                ShSquare.Visible = False
                PicMerge.Line (sx, sy)-((Int(x / factor)), (Int(y / factor))), curdrawcolor, BF
            End If
            If Shapetype = 2 Then
                PicMerge.FillStyle = 0
                PicMerge.FillColor = curdrawcolor
                R = Sqr((sx - (Int(x / factor))) * (sx - (Int(x / factor))) + (sy - (Int(y / factor))) * (sy - (Int(y / factor))))
                PicMerge.Circle (sx, sy), R, curdrawcolor
                PicMerge.FillStyle = 1
                Line1.Visible = False
            End If
            If Shapetype = 3 Then
                PicMerge.FillStyle = 0
                PicMerge.FillColor = curdrawcolor
                PicMerge.ForeColor = curdrawcolor
                Call Ellipse(PicMerge.hdc, sx, sy, (Int(x / factor)), (Int(y / factor)))
                PicMerge.FillStyle = 1
                Line1.Visible = False
            End If
        End If
        PicMerge.DrawWidth = frmMain.Slider2.Value
    End If
    If Curtool = 21 Then
        CWcancel = True
        PicMerge.AutoRedraw = True
        Image1.Picture = PicMerge.Image
        Image1.Refresh
        PicMerge.Refresh
    End If
 End If
Image1.Picture = PicMerge.Image
Image1.Refresh
If Curtool <> 23 Then Backup
woops:
End Sub

Private Sub lbltext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
GetCursorPos P
P.x = Int(P.x / factor) * factor
P.y = Int(P.y / factor) * factor
TextX = P.x - lbltext.Left
TextY = P.y - lbltext.Top
Timer3.Enabled = True
End Sub
Private Sub lbltext_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer3.Enabled = False
End Sub
Private Sub PicFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
LineX.X1 = (x + LeftRuler.ScaleWidth)
LineX.X2 = LineX.X1
LineX.Y1 = 0
LineX.Y2 = 24
LineY.X1 = -LeftRuler.Width
LineY.X2 = 33
LineY.Y1 = y
LineY.Y2 = y
If LineX.X1 < 0 Then LineX.X1 = 0
If LineY.Y1 < 0 Then LineY.Y1 = 0
If LineX.X2 < 0 Then LineX.X2 = 0
If LineY.Y2 < 0 Then LineY.Y2 = 0
lblx = Int(x / factor) - PicBG.Left
lbly = Int(y / factor) - PicBG.Top
End Sub
Private Sub PicFreeSelect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
PicFreeSelect.MousePointer = 15
LineX.X1 = PicFreeSelect.Left + LeftRuler.ScaleWidth + x
LineX.X2 = LineX.X1
LineX.Y1 = 0
LineX.Y2 = 24
LineY.X1 = -LeftRuler.Width
LineY.X2 = 33
LineY.Y1 = PicFreeSelect.Top + y
LineY.Y2 = LineY.Y1
If LineX.X1 < 0 Then LineX.X1 = 0
If LineY.Y1 < 0 Then LineY.Y1 = 0
If LineX.X2 < 0 Then LineX.X2 = 0
If LineY.Y2 < 0 Then LineY.Y2 = 0
lblx = Int(x / factor)
lbly = Int(y / factor)
End Sub
Private Sub PicSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Masterpasting = False Then
    If frmMain.Option4.Value = True Then
        SelHolder.Height = Selectheight
        SelHolder.Width = Selectwidth
        PicSelect.Picture = LoadPicture()
        StretchBlt PicMerge.hdc, SelectLeft, SelectTop, Selectwidth, Selectheight, PicSelect.hdc, 0, 0, Selectwidth, Selectheight, SRCCOPY
        PicMerge.Refresh
        Image1.Picture = PicMerge.Image
    End If
End If
PicSelect.ScaleMode = 2
GetCursorPos P
P.x = Int(P.x / factor) * factor
P.y = Int(P.y / factor) * factor
SelectpicX = P.x - PicSelect.Left
SelectpicY = P.y - PicSelect.Top
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
GetCursorPos P
P.x = Int(P.x / factor) * factor
P.y = Int(P.y / factor) * factor
PicSelect.Left = P.x - SelectpicX
PicSelect.Top = P.y - SelectpicY
End Sub
Private Sub PicSelect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer1.Enabled = False
PicSelect.ScaleMode = 3
SelectLeft = PicSelect.Left / factor
SelectTop = PicSelect.Top / factor
Selectwidth = PicSelect.Width / factor
Selectheight = PicSelect.Height / factor
End Sub
Private Sub SelectImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
PicSelect_MouseDown Button, Shift, x, y
End Sub
Private Sub SelectImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
SelectImage.MousePointer = 15
LineX.X1 = PicSelect.Left + LeftRuler.ScaleWidth + x / Screen.TwipsPerPixelX
LineX.X2 = LineX.X1
LineX.Y1 = 0
LineX.Y2 = 24
LineY.X1 = -LeftRuler.Width / Screen.TwipsPerPixelX
LineY.X2 = 33
LineY.Y1 = PicSelect.Top + y / Screen.TwipsPerPixelY
LineY.Y2 = LineY.Y1
If LineX.X1 < 0 Then LineX.X1 = 0
If LineY.Y1 < 0 Then LineY.Y1 = 0
If LineX.X2 < 0 Then LineX.X2 = 0
If LineY.Y2 < 0 Then LineY.Y2 = 0
lblx = Int(x / factor)
lbly = Int(y / factor)
End Sub
Private Sub SelectImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
PicSelect_MouseUp Button, Shift, x, y
End Sub
Public Sub Picsize()
If PicBG.Width > PicFrame.ScaleWidth Then
    HS.Left = 0
    HS.Top = PicFrame.ScaleHeight - HS.Height - 2
    HS.Width = PicFrame.ScaleWidth
    HS.Visible = True
    If PicBG.Left > 0 Then PicBG.Left = 0
Else
    HS.Visible = False
    PicBG.Left = Int((PicFrame.ScaleWidth - PicBG.Width) / 2)
End If
If PicBG.Height > PicFrame.ScaleHeight Then
    VS.Top = 0
    VS.Left = PicFrame.ScaleWidth - VS.Width
    VS.Height = PicFrame.ScaleHeight
    VS.Visible = True
    If PicBG.Top > 0 Then PicBG.Top = 0
Else
    VS.Visible = False
'    PicBG.Top = Int((PicFrame.ScaleHeight - PicBG.Height) / 2)
    If PicBG.Height < PicFrame.ScaleHeight - 13 Then
        PicBG.Top = Int((PicFrame.ScaleHeight - PicBG.Height) / 2)
    End If
End If
If VS.Visible = True Then If HS.Visible = True Then HS.Width = HS.Width - VS.Width
If HS.Visible = True Then If VS.Visible = True Then VS.Height = VS.Height - HS.Height
If HS.Visible = True And VS.Visible = True Then
    piccorner.Top = HS.Top
    piccorner.Left = VS.Left
    piccorner.Visible = True
Else
    piccorner.Visible = False
End If
If VS.Visible = True Then
    VS.Max = PicBG.ScaleHeight - VS.Height
    VS.LargeChange = VS.Max / 10
End If
If HS.Visible = True Then
    HS.Max = PicBG.Width - HS.Width
    HS.LargeChange = HS.Max / 10
End If
End Sub
Private Sub Timer2_Timer()
If curpos <> RulersVis Then
    If RulersVis Then
        LeftRuler.Visible = True
        TopRuler.Visible = True
        If Me.WindowState = 0 Then
            Me.Height = Me.Height + TopRuler.Height
            Me.Width = Me.Width + LeftRuler.Width + 15
        End If
        Form_Resize
        drawrulers
    Else
        LeftRuler.Visible = False
        TopRuler.Visible = False
        If Me.WindowState = 1 Then
            Me.Height = Me.Height - TopRuler.Height
            Me.Width = Me.Width - LeftRuler.Width + 15
        End If
        Form_Resize
        hiderulers
    End If
End If
curpos = RulersVis
End Sub
Public Sub hiderulers()
If Me.WindowState <> 2 Then
    Me.Width = Me.Width - LeftRuler.Width
    Me.Height = Me.Height - TopRuler.Height
End If
End Sub
Public Sub drawrulers()
Dim x As Integer, Z As Integer, zx As Integer, zxx As Integer, zlen As Integer
        PicLrule.Height = (Image1.Height + 10) * Screen.TwipsPerPixelY
        PicTrule.Width = (Image1.Width + 10) * Screen.TwipsPerPixelX
        ImgTrule.Width = PicTrule.ScaleWidth
        ImgTrule.Left = PicLrule.ScaleWidth + PicFrame.ScaleLeft + PicBG.Left
        ImgLrule.Top = PicBG.Top
        ImgLrule.Height = PicLrule.ScaleHeight
        PicTrule.Cls
        PicLrule.Cls
        PicLrule.DrawWidth = 1
        PicTrule.DrawWidth = 1
Z = 5
zx = 50
zxx = 100
PicTrule.Line (0, 7)-(0, 23), vbRed
For x = 1 To PicMerge.ScaleWidth
    If factor > 1 / 4 Then
        If factor > 5 Then
            If x <> Z And x <> zx And x <> zxx Then
                PicTrule.Line (x * factor, 18)-(x * factor, 23), vbBlack
            End If
        End If
        If x = Z Then
                If factor > 5 Then
                    PicTrule.Line (x * factor, 15)-(x * factor, 23), vbRed
                    If Z <> zx Then
                        PicTrule.CurrentX = Z * factor - (Len(Str(Z)) * 2 + 1)
                        PicTrule.CurrentY = 5
                        PicTrule.Print Str(Z)
                    End If
                Else
                    PicTrule.Line (x * factor, 18)-(x * factor, 23), vbBlack
                End If
            Z = Z + 5
        End If
        If x = zx Then
                PicTrule.Line (x * factor, 15)-(x * factor, 23), vbRed
                PicTrule.CurrentX = zx * factor - (Len(Str(zx)) * 2 + 1)
                PicTrule.CurrentY = 5
                PicTrule.Print Str(zx)
           zx = zx + 50
        End If
        If x = zxx Then
                PicTrule.Line (x * factor, 15)-(x * factor, 23), vbRed
            zxx = zxx + 100
        End If
    Else
        If x = zxx Then
                PicTrule.Line (x * factor, 15)-(x * factor, 23), vbRed
            zxx = zxx + 100
        End If
    End If
Next x
Z = 5
zx = 50
zxx = 100
PicLrule.Line (15, 0)-(31, 0), vbRed
For x = 1 To PicMerge.ScaleHeight
    If factor > 1 / 4 Then
        If factor > 5 Then
            If x <> Z And x <> zx And x <> zxx Then
                PicLrule.Line (26, x * factor)-(31, x * factor), vbBlack
            End If
        End If
        If x = Z Then
                If factor > 5 Then
                    PicLrule.Line (26, x * factor)-(31, x * factor), vbRed
                    If Z <> zx Then
                        PicLrule.CurrentY = Z * factor - 6
                        PicLrule.CurrentX = 0
                        PicLrule.Print Str(Z)
                    End If
                Else
                    PicLrule.Line (26, x * factor)-(31, x * factor), vbBlack
                End If
            Z = Z + 5
        End If
        If x = zx Then
                PicLrule.Line (23, x * factor)-(31, x * factor), vbRed
                PicLrule.CurrentY = zx * factor - 6
                PicLrule.CurrentX = 0
                PicLrule.Print Str(zx)
           zx = zx + 50
        End If
        If x = zxx Then
            PicLrule.Line (23, x * factor)-(31, x * factor), vbRed
            zxx = zxx + 100
        End If
    Else
        If x = zxx Then
                PicLrule.Line (23, x * factor)-(31, x * factor), vbRed
            zxx = zxx + 100
        End If
    End If
Next x
ImgTrule.Picture = PicTrule.Image
ImgLrule.Picture = PicLrule.Image
End Sub
Private Sub Timer3_Timer()
GetCursorPos P
P.x = Int(P.x / factor) * factor
P.y = Int(P.y / factor) * factor
lbltext.Left = P.x - TextX
lbltext.Top = P.y - TextY
End Sub
Private Sub VS_Change()
If VS.Visible = False Then Exit Sub
If RulersVis = False Then
    PicBG.Top = -VS.Value
Else
    PicBG.Top = -VS.Value
    ImgLrule.Top = -VS.Value
End If
End Sub
Private Sub VS_Scroll()
If VS.Visible = False Then Exit Sub
If RulersVis = False Then
    PicBG.Top = -VS.Value
Else
    PicBG.Top = -VS.Value
    ImgLrule.Top = -VS.Value
End If
End Sub
Private Sub hs_Change()
If HS.Visible = False Then Exit Sub
If RulersVis = False Then
    PicBG.Left = -HS.Value
Else
    PicBG.Left = -HS.Value
    ImgTrule.Left = PicBG.Left + LeftRuler.ScaleWidth
End If
End Sub
Private Sub hs_Scroll()
If HS.Visible = False Then Exit Sub
If RulersVis = False Then
    PicBG.Left = -HS.Value
Else
    PicBG.Left = -HS.Value
    ImgTrule.Left = PicBG.Left + LeftRuler.ScaleWidth
End If
End Sub
Public Function MyUndo() As String
MyUndo = ListBackup.List(Val(ListBUorder.List(ListBUorder.ListIndex - 1)))
ListBUorder.Selected(ListBUorder.ListIndex - 1) = True
undolevel = undolevel - 1
If ListBUorder.ListIndex = 0 Then
    frmMain.mnuEditUndo.Enabled = False
    frmMain.TB2.Buttons(9).Enabled = False
    If ListBUorder.ListCount > 1 Then
        frmMain.mnuEditRedo.Enabled = True
        frmMain.TB2.Buttons(10).Enabled = True
    Else
        frmMain.mnuEditRedo.Enabled = False
        frmMain.TB2.Buttons(10).Enabled = False
    End If
    If AFCurfile = "" Then
        DisableMenus
    End If
ElseIf ListBUorder.ListIndex > 0 Then
    frmMain.mnuEditUndo.Enabled = True
    frmMain.TB2.Buttons(9).Enabled = True
    If ListBUorder.ListIndex < ListBUorder.ListCount - 1 Then
        frmMain.mnuEditRedo.Enabled = True
        frmMain.TB2.Buttons(10).Enabled = True
    Else
        frmMain.mnuEditRedo.Enabled = False
        frmMain.TB2.Buttons(10).Enabled = False
    End If
End If
End Function
Public Function MyRedo() As String
MyRedo = ListBackup.List(Val(ListBUorder.List(ListBUorder.ListIndex + 1)))
ListBUorder.Selected(ListBUorder.ListIndex + 1) = True
undolevel = undolevel + 1
If ListBUorder.ListIndex = 0 Then
    frmMain.mnuEditUndo.Enabled = False
    frmMain.TB2.Buttons(9).Enabled = False
    If ListBUorder.ListCount > 1 Then
        frmMain.mnuEditRedo.Enabled = True
        frmMain.TB2.Buttons(10).Enabled = True
    Else
        frmMain.mnuEditRedo.Enabled = False
        frmMain.TB2.Buttons(10).Enabled = False
    End If
ElseIf ListBUorder.ListIndex > 0 Then
    frmMain.mnuEditUndo.Enabled = True
    frmMain.TB2.Buttons(9).Enabled = True
    If ListBUorder.ListIndex < ListBUorder.ListCount - 1 Then
        frmMain.mnuEditRedo.Enabled = True
        frmMain.TB2.Buttons(10).Enabled = True
    Else
        frmMain.mnuEditRedo.Enabled = False
        frmMain.TB2.Buttons(10).Enabled = False
    End If
        EnableMenus
End If
End Function
Public Sub Blankme()
PicMerge.Picture = LoadPicture()
    If RulersVis Then
          Me.Width = PicBG.Width + 120 + LeftRuler.Width
          Me.Height = PicBG.Height + 420 + TopRuler.Height
    Else
        Me.Width = PicBG.Width + 120
        Me.Height = PicBG.Height + 420
    End If
Form_Resize
End Sub
Public Sub Backup()
Dim x As Integer
Dim savedAs As String
Dim temp As Double
If undolevel = 0 Then
   If GetTempFile("", "BI", 0, sfilename) Then
        SavePicture PicMerge.Image, sfilename
        ListBackup.AddItem sfilename
        ListBUorder.AddItem Str(ListBackup.ListCount - 1)
        ListBUorder.Selected(ListBUorder.ListCount - 1) = True
    End If
Else
    ListBUtemp.Clear
    For x = ListBUorder.ListCount - 2 To (ListBUorder.ListCount - 1) + undolevel Step -1
        ListBUtemp.AddItem ListBUorder.List(x)
    Next x
    For x = 0 To ListBUtemp.ListCount - 1
        ListBUorder.AddItem ListBUtemp.List(x)
    Next x
   If GetTempFile("", "BI", 0, sfilename) Then
        SavePicture PicMerge.Image, sfilename
        ListBackup.AddItem sfilename
        ListBUorder.AddItem Str(ListBackup.ListCount - 1)
        ListBUorder.Selected(ListBUorder.ListCount - 1) = True
    End If
End If
If ListBUorder.ListIndex = 0 Then
    frmMain.mnuEditUndo.Enabled = False
    frmMain.TB2.Buttons(9).Enabled = False
    If ListBUorder.ListCount > 1 Then
        frmMain.mnuEditRedo.Enabled = True
        frmMain.TB2.Buttons(10).Enabled = True
    Else
        frmMain.mnuEditRedo.Enabled = False
        frmMain.TB2.Buttons(10).Enabled = False
    End If
ElseIf ListBUorder.ListIndex > 0 Then
    frmMain.mnuEditUndo.Enabled = True
    frmMain.TB2.Buttons(9).Enabled = True
    If ListBUorder.ListIndex < ListBUorder.ListCount - 1 Then
        frmMain.mnuEditRedo.Enabled = True
        frmMain.TB2.Buttons(10).Enabled = True
    Else
        frmMain.mnuEditRedo.Enabled = False
        frmMain.TB2.Buttons(10).Enabled = False
    End If
    EnableMenus
End If
If Val(GetSetting(App.Title, "Settings", "Undoredo", "1")) = 0 Then
    frmPrefs.Combo2.ListIndex = Val(GetSetting(App.Title, "Settings", "UndoSteps", "0"))
    If ListBUorder.ListCount > Val(frmPrefs.Combo2.Text) + 1 Then
    If FileExists(ListBackup.List(Val(ListBUorder.List(0)))) Then Kill ListBackup.List(Val(ListBUorder.List(0)))
        ListBUorder.RemoveItem 0
    End If
End If
If FileExists(savedAs) Then
    temp = FileLen(sfilename)
End If
alreadysaved = False
undolevel = 0
End Sub
Public Sub MakeSelection()
Dim R As Long
Dim r2 As Long
Me.MousePointer = 11
Me.Enabled = False
If maxX = minX Or maxY = minY Then
    PicMerge.Picture = Pic1BU.Image
    Image1.Picture = PicMerge.Image
    Me.MousePointer = 0
    Me.Enabled = True
    Exit Sub
End If
PicFreeSelect.Picture = LoadPicture()
PicFreeSelect.Width = (maxX - minX) * factor
PicFreeSelect.Height = (maxY - minY) * factor
SelHolder.Picture = LoadPicture()
SelHolder.Width = (maxX - minX)
SelHolder.Height = (maxY - minY)
PicFreeSelect.Top = minY * factor
PicFreeSelect.Left = minX * factor
If minX <> 0 And minY <> 0 Then
    Call Filling2(PicMask, PicMask.Point(0, 0), 0, 0, 0)
End If
For i = minX To maxX
    For j = minY To maxY
        R = GetPixel(PicMask.hdc, i, j)
        If R <> vbBlack Then
            r2 = GetPixel(PicMerge.hdc, i, j)
            SetPixel SelHolder.hdc, i - minX, j - minY, r2
            If frmMain.Option4.Value = True Then
                SetPixel Pic1BU.hdc, i, j, PicMerge.BackColor
            End If
        End If
    Next j
Next i
TransparentBlt PicFreeSelect.hdc, 0, 0, PicFreeSelect.Width, PicFreeSelect.Height, SelHolder.hdc, 0, 0, SelHolder.Width, SelHolder.Height, 8950944
PicFreeSelect.Refresh
PicMerge.Refresh
Image1.Picture = PicMerge.Image
ShapeMe 8950944, True, , PicFreeSelect
PicFreeSelect.Visible = True
minX = 0
maxX = 0
minY = 0
maxY = 0
Me.MousePointer = 0
Me.Enabled = True
End Sub
Private Sub PicFreeSelect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If dontusePicBU = False Then
    PicMerge.Picture = Pic1BU.Image
    Image1.Picture = PicMerge.Image
End If
PicFreeSelect.ScaleMode = 2
GetCursorPos P
P.x = Int(P.x / factor) * factor
P.y = Int(P.y / factor) * factor
SelectpicX = P.x - PicFreeSelect.Left
SelectpicY = P.y - PicFreeSelect.Top
TimerFS.Enabled = True
End Sub
Private Sub TimerFS_Timer()
GetCursorPos P
P.x = Int(P.x / factor) * factor
P.y = Int(P.y / factor) * factor
PicFreeSelect.Left = P.x - SelectpicX
PicFreeSelect.Top = P.y - SelectpicY
End Sub
Private Sub PicFreeSelect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
TimerFS.Enabled = False
PicFreeSelect.ScaleMode = 3
Pic1BU.Picture = PicMerge.Image
End Sub
Public Sub ShapeMe(Color As Long, HorizontalScan As Boolean, Optional Name1 As Form = Nothing, Optional Name2 As PictureBox = Nothing)
Dim x As Integer, y As Integer
Dim dblHeight As Double, dblWidth As Double
Dim lngHDC As Long
Dim booMiddleOfSet As Boolean
Dim colPoints As Collection
Set colPoints = New Collection
Dim Z As Variant
Dim dblTransY As Double
Dim dblTransStartX As Double
Dim dblTransEndX As Double
Dim Name As Object
Set Name = Name2
With Name
    .AutoRedraw = True
    .ScaleMode = 3
    lngHDC = .hdc
    If HorizontalScan = True Then
        dblHeight = .ScaleHeight
        dblWidth = .ScaleWidth
    Else
        dblHeight = .ScaleWidth
        dblWidth = .ScaleHeight
    End If
End With
booMiddleOfSet = False
For y = 0 To dblHeight
    dblTransY = y
    For x = 0 To dblWidth
        If TypeOf Name Is Form Then
            If GetPixel(lngHDC, x, y) = Color Then
                If booMiddleOfSet = False Then
                    dblTransStartX = x
                    dblTransEndX = x
                    booMiddleOfSet = True
                Else
                    dblTransEndX = x
                End If
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(dblTransY, dblTransStartX, dblTransEndX)
                    booMiddleOfSet = False
                End If
            End If
         ElseIf TypeOf Name Is PictureBox Then
            If Name.Point(x, y) = Color Then
                If booMiddleOfSet = False Then
                    dblTransStartX = x
                    dblTransEndX = x
                    booMiddleOfSet = True
                Else
                    dblTransEndX = x
                End If
            Else
                If booMiddleOfSet Then
                    colPoints.Add Array(dblTransY, dblTransStartX, dblTransEndX)
                    booMiddleOfSet = False
                End If
            End If
        End If
    Next x
Next y
CurRgn = CreateRectRgn(0, 0, dblWidth, dblHeight)
For Each Z In colPoints
    TempRgn = CreateRectRgn(Z(1), Z(0), Z(2) + 1, Z(0) + 1)
    CombineRgn CurRgn, CurRgn, TempRgn, RGN_DIFF
    DeleteObject (TempRgn)
Next
SetWindowRgn Name.Hwnd, CurRgn, True
DeleteObject (CurRgn)
Set colPoints = Nothing
End Sub


Public Sub NeedtoPaste()
'Used to paste an existing selection if the
'user presses Paste
If PicSelect.Visible = True Then
    SelectLeft = PicSelect.Left / factor
    SelectTop = PicSelect.Top / factor
    Selectwidth = PicSelect.Width / factor
    Selectheight = PicSelect.Height / factor
    PicSelect.Visible = False
    SelectShape.Visible = False
    StretchBlt PicMerge.hdc, SelectLeft, SelectTop, Selectwidth, Selectheight, SelHolder.hdc, 0, 0, Selectwidth, Selectheight, SRCCOPY
    Image1.Picture = PicMerge.Image
    frmMain.TB2.Buttons(5).Enabled = False
    frmMain.TB2.Buttons(6).Enabled = False
    frmMain.mnuEditCut.Enabled = False
    frmMain.mnuEditCopy.Enabled = False
    Backup
    selecting = False
    Masterpasting = False
ElseIf PicFreeSelect.Visible = True Then
    TransparentBlt Pic1BU.hdc, Int(PicFreeSelect.Left / factor), Int(PicFreeSelect.Top / factor), Int(PicFreeSelect.ScaleWidth / factor), Int(PicFreeSelect.ScaleHeight / factor), SelHolder.hdc, 0, 0, SelHolder.ScaleWidth, SelHolder.ScaleHeight, 8950944
    PicMerge.Picture = Pic1BU.Image
    PicMerge.Refresh
    Image1.Picture = PicMerge.Image
    PicFreeSelect.Visible = False
    PicFreeSelect.Picture = LoadPicture()
    Backup
    pasting = True
    Masterpasting = False
ElseIf lbltext.Visible = True Then
    PicMerge.FontName = lbltext.FontName
    PicMerge.FontStrikethru = lbltext.FontStrikethru
    PicMerge.FontUnderline = lbltext.FontUnderline
    PicMerge.ForeColor = lbltext.ForeColor
    PicMerge.FontSize = lbltext.FontSize
    PicMerge.FontBold = lbltext.FontBold
    PicMerge.FontItalic = lbltext.FontItalic
    PicMerge.CurrentX = lbltext.Left
    PicMerge.CurrentY = lbltext.Top
    PicMerge.Print lbltext.Caption
    lbltext.Visible = False
End If
Image1.Picture = PicMerge.Image
Image1.Refresh
Backup

End Sub
