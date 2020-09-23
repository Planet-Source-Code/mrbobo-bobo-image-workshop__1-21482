VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Image Workshop"
   ClientHeight    =   5880
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   8895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   5040
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4560
      Top             =   720
   End
   Begin MSComctlLib.ImageList TBImage 
      Left            =   2160
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0554
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0666
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0778
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":088A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":099C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AAE
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BC0
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CD2
            Key             =   "Pointer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FEE
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18CA
            Key             =   "FreeSelect"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BE6
            Key             =   "Pen"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24C2
            Key             =   "Eraser"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27DE
            Key             =   "Flood"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AFA
            Key             =   "Line"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E16
            Key             =   "PolyLine"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3132
            Key             =   "Shapes"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A0E
            Key             =   "ColorPick"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46EA
            Key             =   "ColorWand"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53C6
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":56E2
            Key             =   "Mag"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FBE
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":62DA
            Key             =   "Unlock"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65F6
            Key             =   "ColorPick"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6912
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A2E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B4A
            Key             =   "Clear1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":720E
            Key             =   "Clear"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ToolsTB 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8895
      TabIndex        =   61
      Top             =   0
      Width           =   8895
      Begin MSComctlLib.Toolbar TB2 
         Height          =   330
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "TBImage"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   24
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "Standard pointer - no edit tool selected."
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Save"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Cut"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Paste"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Undo"
               ImageKey        =   "Undo"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "Redo"
               ImageKey        =   "Redo"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Pointer"
               ImageKey        =   "Pointer"
               Style           =   2
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Select"
               Object.ToolTipText     =   "Select an area in the current image."
               ImageKey        =   "Select"
               Style           =   2
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FreeSelect"
               Object.ToolTipText     =   "Free Hand Selection"
               ImageKey        =   "FreeSelect"
               Style           =   2
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Pen"
               Object.ToolTipText     =   "Draw a freehand line."
               ImageKey        =   "Pen"
               Style           =   2
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Eraser"
               Object.ToolTipText     =   "Erase parts of the current image."
               ImageKey        =   "Eraser"
               Style           =   2
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Flood"
               Object.ToolTipText     =   "Flood fill areas of the current image."
               ImageKey        =   "Flood"
               Style           =   2
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line"
               Object.ToolTipText     =   "Draw a straight line."
               ImageKey        =   "Line"
               Style           =   2
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PolyLine"
               Object.ToolTipText     =   "Draw a series of connected straight lines."
               ImageKey        =   "PolyLine"
               Style           =   2
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Shapes"
               Object.ToolTipText     =   "Draw preset shapes."
               ImageKey        =   "Shapes"
               Style           =   2
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ColorWand"
               Object.ToolTipText     =   "Alter colors."
               ImageKey        =   "ColorWand"
               Style           =   2
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ColorPick"
               ImageKey        =   "ColorPick"
               Style           =   2
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Mag"
               ImageKey        =   "Mag"
               Style           =   2
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Text"
               Object.ToolTipText     =   "Insert Text"
               ImageKey        =   "Text"
            EndProperty
         EndProperty
         Begin VB.PictureBox FillPic 
            Height          =   540
            Left            =   480
            Picture         =   "frmMain.frx":78D2
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   65
            Top             =   360
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox MagPic 
            Height          =   540
            Left            =   120
            Picture         =   "frmMain.frx":7BDC
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   64
            Top             =   360
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.PictureBox EyePic 
            Height          =   540
            Left            =   960
            Picture         =   "frmMain.frx":84A6
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   540
         End
      End
   End
   Begin VB.PictureBox ColTB 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5505
      Left            =   7005
      ScaleHeight     =   5505
      ScaleWidth      =   1890
      TabIndex        =   30
      Top             =   375
      Width           =   1890
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H008D5F07&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   30
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   59
         Top             =   0
         Width           =   1875
         Begin VB.CommandButton Command2 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1635
            TabIndex        =   60
            Top             =   30
            Width           =   200
         End
      End
      Begin VB.PictureBox ColorBox 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   120
         ScaleHeight     =   417
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   131
         TabIndex        =   37
         Top             =   1680
         Width           =   1965
         Begin VB.PictureBox picColBlowup 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   2910
            Left            =   600
            ScaleHeight     =   194
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   49
            TabIndex        =   38
            Top             =   3000
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   330
            Left            =   0
            TabIndex        =   44
            Top             =   1920
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            ImageList       =   "ImageList3"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Lock"
                  Object.ToolTipText     =   "Dont change my selected colors."
                  ImageKey        =   "Lock"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "Unlock"
                  Object.ToolTipText     =   "Automatically add to my collection of colors."
                  ImageKey        =   "Unlock"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Clear"
                  Object.ToolTipText     =   "Remove the current collection."
                  ImageKey        =   "Clear"
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ColDlg"
                  Object.ToolTipText     =   "Use Windows Color Selector."
                  ImageKey        =   "ColorPick"
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpenCol"
                  Object.ToolTipText     =   "Open a Collection of Colors previously saved."
                  ImageKey        =   "Open"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SaveCols"
                  Object.ToolTipText     =   "Save My Collection of Colors"
                  ImageKey        =   "Save"
               EndProperty
            EndProperty
         End
         Begin VB.PictureBox TheColor 
            BackColor       =   &H80000007&
            Height          =   255
            Left            =   0
            ScaleHeight     =   195
            ScaleWidth      =   1155
            TabIndex        =   42
            Top             =   1560
            Width           =   1215
            Begin VB.Label LblCurCol 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "CURRENT"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   -120
               TabIndex        =   43
               Top             =   30
               Width           =   1215
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1260
            Left            =   1440
            ScaleHeight     =   1260
            ScaleWidth      =   255
            TabIndex        =   41
            Top             =   240
            Width           =   255
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1440
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   40
            Top             =   1560
            Width           =   255
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1200
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   39
            Top             =   1560
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   1260
            Left            =   0
            MouseIcon       =   "frmMain.frx":9170
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":947A
            Stretch         =   -1  'True
            ToolTipText     =   "Right click for more concise color selection"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblColor 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   1695
         End
      End
      Begin VB.VScrollBar VSblue 
         Height          =   220
         Left            =   705
         Max             =   255
         TabIndex        =   36
         Top             =   1305
         Value           =   2
         Width           =   220
      End
      Begin VB.VScrollBar VSgreen 
         Height          =   220
         Left            =   705
         Max             =   255
         TabIndex        =   35
         Top             =   915
         Value           =   2
         Width           =   220
      End
      Begin VB.VScrollBar VSred 
         Height          =   220
         Left            =   705
         Max             =   255
         TabIndex        =   34
         Top             =   525
         Value           =   2
         Width           =   220
      End
      Begin VB.VScrollBar VShue 
         Height          =   220
         Left            =   1575
         Max             =   240
         TabIndex        =   33
         Top             =   525
         Value           =   2
         Width           =   220
      End
      Begin VB.VScrollBar VSsat 
         Height          =   220
         Left            =   1575
         Max             =   240
         TabIndex        =   32
         Top             =   915
         Value           =   2
         Width           =   220
      End
      Begin VB.VScrollBar VSlum 
         Height          =   220
         Left            =   1575
         Max             =   240
         TabIndex        =   31
         Top             =   1305
         Value           =   2
         Width           =   220
      End
      Begin MSComctlLib.ProgressBar Pb 
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   255
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.TextBox txtRed 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   300
         TabIndex        =   47
         Text            =   "1"
         Top             =   480
         Width           =   665
      End
      Begin VB.TextBox txtGreen 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   300
         TabIndex        =   48
         Text            =   "1"
         Top             =   870
         Width           =   665
      End
      Begin VB.TextBox txtBlue 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   300
         TabIndex        =   49
         Text            =   "1"
         Top             =   1260
         Width           =   665
      End
      Begin VB.TextBox txtHue 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1155
         TabIndex        =   50
         Text            =   "1"
         Top             =   480
         Width           =   665
      End
      Begin VB.TextBox txtSat 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1155
         TabIndex        =   51
         Text            =   "1"
         Top             =   870
         Width           =   665
      End
      Begin VB.TextBox txtLum 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   1155
         TabIndex        =   52
         Text            =   "1"
         Top             =   1260
         Width           =   665
      End
      Begin VB.Label Label12 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   58
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   57
         Top             =   930
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   56
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   55
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   54
         Top             =   930
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1020
         TabIndex        =   53
         Top             =   540
         Width           =   255
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21B6C
            Key             =   "round"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21E90
            Key             =   "left"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":221B4
            Key             =   "right"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":224D8
            Key             =   "hori"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":227FC
            Key             =   "vert"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22E3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23158
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23474
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23790
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":240E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ToolSettings 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5505
      Left            =   0
      ScaleHeight     =   5505
      ScaleWidth      =   1905
      TabIndex        =   0
      Top             =   375
      Width           =   1900
      Begin VB.PictureBox PicTools 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   0
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   1860
         TabIndex        =   3
         Top             =   300
         Width           =   1860
         Begin MSComctlLib.Slider Slider2 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Draw Thickness"
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   50
            SelStart        =   1
            TickFrequency   =   5
            Value           =   1
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Length of Slash or Bar"
            Top             =   1440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            Min             =   1
            Max             =   50
            SelStart        =   1
            TickFrequency   =   5
            Value           =   1
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   390
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            ImageList       =   "ImageList2"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "round"
                  Object.ToolTipText     =   "Round"
                  ImageKey        =   "round"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "left"
                  Object.ToolTipText     =   "Left Slash"
                  ImageKey        =   "left"
                  Style           =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "right"
                  Object.ToolTipText     =   "Right Slash"
                  ImageKey        =   "right"
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "hori"
                  Object.ToolTipText     =   "Horizontal Bar"
                  ImageKey        =   "hori"
                  Style           =   2
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "vert"
                  Object.ToolTipText     =   "Vertical Bar"
                  ImageKey        =   "vert"
                  Style           =   2
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdTools 
            BackColor       =   &H80000016&
            Caption         =   "Pen"
            Height          =   300
            Index           =   0
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "DRAW THICKNESS"
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
            TabIndex        =   14
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "TIP  WIDTH"
            Enabled         =   0   'False
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
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.PictureBox PicTools 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   1
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   1860
         TabIndex        =   4
         Top             =   600
         Width           =   1860
         Begin VB.OptionButton Option2 
            Caption         =   "FILLED"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "OUTLINE"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   17
            Top             =   840
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdTools 
            BackColor       =   &H80000016&
            Caption         =   "Shapes"
            Height          =   300
            Index           =   1
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   0
            Width           =   1815
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   390
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            ImageList       =   "ImageList2"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "square"
                  Object.ToolTipText     =   "Circle"
                  ImageIndex      =   6
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "oval1"
                  Object.ToolTipText     =   "Oval"
                  ImageIndex      =   8
                  Style           =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "circle"
                  Object.ToolTipText     =   "Square"
                  ImageIndex      =   10
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "oval"
                  Object.ToolTipText     =   "Rectangle"
                  ImageIndex      =   12
                  Style           =   2
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Slider Slider3 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "Length of Slash or Bar"
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   50
            SelStart        =   1
            TickFrequency   =   5
            Value           =   1
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "OUTLINE  WIDTH"
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
            Left            =   240
            TabIndex        =   20
            Top             =   1440
            Width           =   1455
         End
      End
      Begin VB.PictureBox PicTools 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   4
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   1860
         TabIndex        =   66
         Top             =   1500
         Width           =   1860
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H80000016&
            Caption         =   "CLEAR"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   330
            Width           =   535
         End
         Begin VB.CheckBox ChClipLock 
            Caption         =   "LOCK  CLIPBOARD"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   80
            Top             =   330
            Width           =   1215
         End
         Begin VB.PictureBox PicClip 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   5
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   79
            Top             =   2760
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox PicClip 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   4
            Left            =   0
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   78
            Top             =   2640
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox PicClip 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   3
            Left            =   360
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   77
            Top             =   2640
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox PicClip 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   2
            Left            =   720
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   76
            Top             =   2760
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox PicClip 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   1
            Left            =   480
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   75
            Top             =   2760
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox PicClipBG 
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   5
            Left            =   1080
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   74
            Top             =   2400
            Visible         =   0   'False
            Width           =   675
            Begin VB.Image ImgClip 
               Height          =   615
               Index           =   5
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   615
            End
         End
         Begin VB.PictureBox PicClipBG 
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   4
            Left            =   210
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   73
            Top             =   2400
            Visible         =   0   'False
            Width           =   675
            Begin VB.Image ImgClip 
               Height          =   615
               Index           =   4
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   615
            End
         End
         Begin VB.PictureBox PicClipBG 
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   3
            Left            =   1080
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   72
            Top             =   1560
            Visible         =   0   'False
            Width           =   675
            Begin VB.Image ImgClip 
               Height          =   615
               Index           =   3
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   615
            End
         End
         Begin VB.PictureBox PicClipBG 
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   2
            Left            =   210
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   71
            Top             =   1560
            Visible         =   0   'False
            Width           =   675
            Begin VB.Image ImgClip 
               Height          =   615
               Index           =   2
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   615
            End
         End
         Begin VB.PictureBox PicClipBG 
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   1
            Left            =   1080
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   70
            Top             =   720
            Visible         =   0   'False
            Width           =   675
            Begin VB.Image ImgClip 
               Height          =   615
               Index           =   1
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   615
            End
         End
         Begin VB.PictureBox PicClipBG 
            BackColor       =   &H80000014&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   0
            Left            =   210
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   69
            Top             =   720
            Visible         =   0   'False
            Width           =   675
            Begin VB.Image ImgClip 
               Height          =   615
               Index           =   0
               Left            =   30
               Stretch         =   -1  'True
               Top             =   30
               Width           =   615
            End
         End
         Begin VB.PictureBox PicClip 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   615
            Index           =   0
            Left            =   240
            ScaleHeight     =   615
            ScaleWidth      =   615
            TabIndex        =   68
            Top             =   2520
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdTools 
            BackColor       =   &H80000016&
            Caption         =   "Clipboard"
            Height          =   300
            Index           =   4
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   0
            Width           =   1815
         End
         Begin VB.Shape ShClip 
            BorderColor     =   &H00800000&
            BorderWidth     =   5
            Height          =   855
            Left            =   120
            Top             =   630
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.PictureBox PicTools 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   2
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   1860
         TabIndex        =   5
         Top             =   900
         Width           =   1860
         Begin VB.OptionButton Option4 
            Caption         =   "CUT AND DRAG"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "COPY AND DRAG"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdTools 
            BackColor       =   &H80000016&
            Caption         =   "Selection"
            Height          =   300
            Index           =   2
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox PicTools 
         BorderStyle     =   0  'None
         Height          =   300
         Index           =   3
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   1860
         TabIndex        =   6
         Top             =   1200
         Width           =   1860
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmMain.frx":24400
            Left            =   240
            List            =   "frmMain.frx":24416
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdTools 
            BackColor       =   &H80000016&
            Caption         =   "Color Wand"
            Height          =   300
            Index           =   3
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   0
            Width           =   1815
         End
         Begin MSComctlLib.Slider Slider4 
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1365
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            SelStart        =   1
            Value           =   1
         End
         Begin MSComctlLib.Slider Slider5 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   885
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   1
            Max             =   100
            SelStart        =   1
            TickFrequency   =   10
            Value           =   1
         End
         Begin MSComctlLib.Slider Slider6 
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Min             =   10
            Max             =   50
            SelStart        =   10
            TickFrequency   =   4
            Value           =   10
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "STEPS"
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
            Left            =   240
            TabIndex        =   27
            Top             =   1695
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "RADIUS"
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
            Left            =   240
            TabIndex        =   26
            Top             =   765
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "INTENSITY"
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
            Left            =   240
            TabIndex        =   25
            Top             =   1245
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H008D5F07&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   30
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   124
         TabIndex        =   1
         Top             =   0
         Width           =   1855
         Begin VB.CommandButton Command1 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1635
            TabIndex        =   2
            Top             =   30
            Width           =   200
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   5
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileTwain 
         Caption         =   "Twain"
         Begin VB.Menu mnuFileAquire 
            Caption         =   "Aquire"
         End
         Begin VB.Menu mnuFileTSelect 
            Caption         =   "Select Source"
         End
      End
      Begin VB.Menu mnuFilePrinter 
         Caption         =   "Printer"
         Begin VB.Menu mnuFilePrinterSetup 
            Caption         =   "Printer Setup"
         End
         Begin VB.Menu mnuFilePrinterPrint 
            Caption         =   "Print"
         End
      End
      Begin VB.Menu mnuFileSaveOptions 
         Caption         =   "Preferences"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteNew 
         Caption         =   "Paste As New Image"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuEditSelectNone 
         Caption         =   "Select None"
      End
      Begin VB.Menu mnuEditSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClip 
         Caption         =   "Clipboard Monitor"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "&Image"
      Begin VB.Menu mnuImageProperties 
         Caption         =   "Image Properties"
      End
      Begin VB.Menu spac1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlipV 
         Caption         =   "Flip Vertical"
      End
      Begin VB.Menu mnuFlipH 
         Caption         =   "Flip Horizontal"
      End
      Begin VB.Menu spac2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRotate90 
         Caption         =   "Rotate 90 degrees"
      End
      Begin VB.Menu mnuRotate180 
         Caption         =   "Rotate 180 degrees"
      End
      Begin VB.Menu mnuRotate270 
         Caption         =   "Rotate 270 degrees"
      End
      Begin VB.Menu spac3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResize 
         Caption         =   "Resize Image"
      End
      Begin VB.Menu mnuAddBorders 
         Caption         =   "Add Borders"
      End
      Begin VB.Menu mnuBorderWiz 
         Caption         =   "Frames and Edges"
      End
      Begin VB.Menu mnuFiltBrow 
         Caption         =   "Filter Browser"
      End
      Begin VB.Menu spac4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIncCol 
         Caption         =   "Increase Color"
         Begin VB.Menu mnuColInc 
            Caption         =   "Red"
            Index           =   0
         End
         Begin VB.Menu mnuColInc 
            Caption         =   "Green"
            Index           =   1
         End
         Begin VB.Menu mnuColInc 
            Caption         =   "Blue"
            Index           =   2
         End
      End
      Begin VB.Menu mnuRedCol 
         Caption         =   "Reduce Color"
         Begin VB.Menu mnuColRed 
            Caption         =   "Red"
            Index           =   0
         End
         Begin VB.Menu mnuColRed 
            Caption         =   "Green"
            Index           =   1
         End
         Begin VB.Menu mnuColRed 
            Caption         =   "Blue"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewColorBox 
         Caption         =   "Color Box"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSettingsBar 
         Caption         =   "Settings &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewRulers 
         Caption         =   "Rulers"
      End
      Begin VB.Menu Zspac 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoomActual 
         Caption         =   "Actual Size"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom In"
         Begin VB.Menu mnuZoomInX 
            Caption         =   "1:1"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "2:1"
            Index           =   1
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "3:1"
            Index           =   2
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "4:1"
            Index           =   3
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "5:1"
            Index           =   4
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "6:1"
            Index           =   5
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "7:1"
            Index           =   6
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "8:1"
            Index           =   7
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "9:1"
            Index           =   8
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "10:1"
            Index           =   9
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "11:1"
            Index           =   10
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "12:1"
            Index           =   11
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "13:1"
            Index           =   12
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "14:1"
            Index           =   13
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "15:1"
            Index           =   14
         End
         Begin VB.Menu mnuZoomInX 
            Caption         =   "16:1"
            Index           =   15
         End
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom Out"
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:1"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:2"
            Index           =   1
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:3"
            Index           =   2
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:4"
            Index           =   3
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:5"
            Index           =   4
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:6"
            Index           =   5
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:7"
            Index           =   6
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:8"
            Index           =   7
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:9"
            Index           =   8
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:10"
            Index           =   9
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:11"
            Index           =   10
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:12"
            Index           =   11
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:13"
            Index           =   12
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:14"
            Index           =   13
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:15"
            Index           =   14
         End
         Begin VB.Menu mnuZoomOutX 
            Caption         =   "1:16"
            Index           =   15
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                        
Dim strInt As Integer
Dim dontupdateVS As Boolean
Dim HSLV As HSLCol
Dim colleft As Integer
Dim coltop As Integer
Dim colloaded As Boolean
Dim colsave(0 To 77) As String
Dim colsaved As Boolean
Dim ActiveClipboard As Integer
Dim curClipboard As Integer
Dim Clipheight As Integer
Dim starttimer As Integer
Dim MRUpath() As String
Public MRUlimit As Integer
Dim myCommand As String
Dim Shelled As Boolean
Private Sub ChClipLock_Click()
If ChClipLock.Value = 1 Then
    DontAdd = True
Else
    DontAdd = False
End If
End Sub
Private Sub cmdClear_Click()
If MsgBox("Are you sure you wish to clear the Clipboard ?", vbYesNo + vbQuestion, "Bobo Enterprises") = vbNo Then
    Exit Sub
Else
    Clipboard.Clear
    For x = 0 To 5
        PicClip(x).Picture = LoadPicture()
        ImgClip(x).Picture = PicClip(x).Image
        PicClipBG(x).Visible = False
    Next x
    ShClip.Visible = False
    Clipheight = 580
    curClipboard = 0
    cmdClear.Enabled = False
End If
End Sub
Private Sub cmdTools_Click(Index As Integer)
'This is a simple 'Outlook' style toolbar
Dim tempH As Integer
If Index = 0 Then tempH = 1845
If Index = 1 Then tempH = 1980
If Index = 2 Then tempH = 900
If Index = 3 Then tempH = 2205
If Index = 4 Then tempH = Clipheight
If PicTools(Index).Height > 300 Then
    PicTools(Index).Height = 300
Else
    PicTools(Index).Height = tempH
    If Index = 0 Then
        TB2.Buttons(15).Value = tbrPressed
        Curtool = 15
    End If
    If Index = 1 Then
        TB2.Buttons(20).Value = tbrPressed
        Curtool = 20
    End If
    If Index = 2 Then
        If Curtool <> 14 Then
            TB2.Buttons(13).Value = tbrPressed
            Curtool = 13
        Else
            TB2.Buttons(14).Value = tbrPressed
        End If
    End If
    If Index = 3 Then
        TB2.Buttons(21).Value = tbrPressed
        Curtool = 21
    End If
End If
For x = 1 To 4
    PicTools(x).Top = PicTools(x - 1).Top + PicTools(x - 1).Height
Next x
If PicTools(4).Top + PicTools(4).Height > ToolSettings.Height Then
    For x = 4 To 1 Step -1
        If x <> Index Then
            PicTools(x).Height = 300
            For y = 1 To 4
                PicTools(y).Top = PicTools(y - 1).Top + PicTools(y - 1).Height
            Next y
            If PicTools(4).Top + PicTools(4).Height < ToolSettings.Height Then Exit For
        End If
    Next x
End If
End Sub
Private Sub Command1_Click()
ToolSettings.Visible = False
mnuViewSettingsBar.Checked = False
End Sub
Private Sub Command2_Click()
ColTB.Visible = False
mnuViewColorBox.Checked = False
End Sub
Private Sub ImgClip_Click(Index As Integer)
    ActiveClipboard = Index
    ShClip.Left = PicClipBG(ActiveClipboard).Left - 90
    ShClip.Top = PicClipBG(ActiveClipboard).Top - 90
    DontAdd = True
    Clipboard.Clear
    Clipboard.SetData PicClip(ActiveClipboard).Image
    If PicClip(ActiveClipboard).Tag = "free" Then
        freeselection = True
    Else
        freeselection = False
    End If
    ShClip.Visible = True
    If ChClipLock.Value = 0 Then DontAdd = False
End Sub

Private Sub MDIForm_Load()
If TWAIN_IsAvailable() = 0 Then mnuFileTwain.Enabled = False Else mnuFileTwain.Enabled = True
Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
mnuViewRulers.Checked = GetSetting(App.Title, "Settings", "Rulers", False)
mnuViewToolbar.Checked = GetSetting(App.Title, "Settings", "Toolbar", True)
mnuViewColorBox.Checked = GetSetting(App.Title, "Settings", "ColorBox", True)
mnuViewSettingsBar.Checked = GetSetting(App.Title, "Settings", "SettingsBar", True)
RulersVis = mnuViewRulers.Checked
ToolsTB.Visible = mnuViewToolbar.Checked
ColTB.Visible = mnuViewColorBox.Checked
ToolSettings.Visible = mnuViewSettingsBar.Checked
frmPrefs.Combo1.ListIndex = Val(GetSetting(App.Title, "Settings", "MRUlist", "1"))
MRUlimit = Val(frmPrefs.Combo1.Text)
ReDim MRUpath(0 To MRUlimit - 1)
LoadMRUs mnuFileMRU(0), mnuFileBar5
myCommand = Command()
'This procedure is used to loaded a shelled file into the
'existing instance without sponing another instance
'It is linked to Timer2
If myCommand <> "" Then
    If App.PrevInstance Then
        Open App.path + "\BBdde.tmp" For Output As #1
           Print #1, myCommand
       Close #1
        End
    End If
    Shelled = True
    myCommand = GetLongFilename(myCommand)
End If
StartupConfig
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
finalclose = True
Unload frmNew
Unload frmResize
Unload frmFiltBrow
Unload frmAddBord
Unload frmBWiz
Unload frmText
Unload frmImgprops
Unload frmSplash
Unload frmPrefs
SaveSetting App.Title, "Settings", "Rulers", RulersVis
SaveSetting App.Title, "Settings", "Toolbar", ToolsTB.Visible
SaveSetting App.Title, "Settings", "ColorBox", ColTB.Visible
SaveSetting App.Title, "Settings", "SettingsBar", ToolSettings.Visible
If Me.WindowState <> vbMinimized Then
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
End If
UnHookForm Me
End Sub
Private Sub mnuAddBorders_Click()
frmAddBord.Show vbModal
If ActiveForm.WindowState <> 2 Then
    If RulersVis = False Then
        ActiveForm.Width = ActiveForm.PicMerge.ScaleWidth * Screen.TwipsPerPixelX + 120
        ActiveForm.Height = ActiveForm.PicMerge.ScaleHeight * Screen.TwipsPerPixelY + 420
    Else
        ActiveForm.Width = ActiveForm.PicMerge.ScaleWidth * Screen.TwipsPerPixelX + ActiveForm.LeftRuler.Width + 120
        ActiveForm.Height = ActiveForm.PicMerge.ScaleHeight * Screen.TwipsPerPixelY + ActiveForm.TopRuler.Height + 420
    End If
End If
End Sub
Private Sub mnuBorderWiz_Click()
Dim chsrcpic As PictureBox
Dim chdestpic As PictureBox
Dim selArea As Boolean
If ActiveForm.PicFreeSelect.Visible = True Then
    MsgBox "If selecting an area for this function the area must be rectangular.", vbInformation, "Bobo Enterprises"
    Exit Sub
End If
If ActiveForm.PicSelect.Visible = True Then
    Set chdestpic = ActiveForm.SelHolder
    ActiveForm.PicSelect.Picture = chdestpic.Image
    Set chsrcpic = ActiveForm.PicSelect
    selArea = True
Else
    Set chdestpic = ActiveForm.PicMerge
    ActiveForm.picsource.Picture = chdestpic.Image
    Set chsrcpic = ActiveForm.picsource
    selArea = False
End If
If curborder > 0 Then
    frmBWiz.BordList.ListIndex = curborder
Else
    frmBWiz.BordList.ListIndex = 0
End If
If borderwidth <> 0 Then
    frmBWiz.SL1.Value = borderwidth
Else
    frmBWiz.SL1.Value = 20
End If
If curborderlevel2 <> 0 Then
    frmBWiz.SL2.Value = curborderlevel2
Else
    frmBWiz.SL2.Value = frmBWiz.SL2.Max / 2
End If
If curborderlevel3 <> 0 Then
    frmBWiz.SL3.Value = curborderlevel3
Else
    frmBWiz.SL3.Value = frmBWiz.SL3.Max / 2
End If
frmBWiz.cmdColorsel.BackColor = chBGcolor
frmBWiz.Startup chsrcpic
frmBWiz.Show vbModal
If BWcancel = True Then
    BWcancel = False
    Exit Sub
Else
    If curborder = 0 Then
        Butt1 chsrcpic, chdestpic, borderwidth, curborderlevel2, curborderlevel3, outline, inline, frmMain.Pb
    ElseIf curborder = 1 Then
        Butt2 chsrcpic, chdestpic, borderwidth, curborderlevel2, curborderlevel3, outline, inline, frmMain.Pb
    ElseIf curborder = 2 Then
            Frame3D chsrcpic, chdestpic, borderwidth, framewidth, 12632256, frmMain.Pb
    ElseIf curborder = 3 Then
            Frame3D chsrcpic, chdestpic, borderwidth, framewidth, chBGcolor, frmMain.Pb
    ElseIf curborder = 4 Then
        FlatBorder chsrcpic, chdestpic, borderwidth, chBGcolor, outline, inline, frmMain.Pb
    ElseIf curborder = 5 Then
        ButtBW chsrcpic, chdestpic, borderwidth, curborderlevel2, curborderlevel3, outline, inline, True, frmMain.Pb
    ElseIf curborder = 6 Then
        ButtBW chsrcpic, chdestpic, borderwidth, curborderlevel2, curborderlevel3, outline, inline, False, frmMain.Pb
    End If
If selArea = True Then
    ActiveForm.SelectShape.Visible = False
    StretchBlt ActiveForm.PicMerge.hdc, ActiveForm.SelectLeft, ActiveForm.SelectTop, ActiveForm.Selectwidth, ActiveForm.Selectheight, ActiveForm.SelHolder.hdc, 0, 0, chsrcpic.Width, chsrcpic.Height, SRCCOPY
    ActiveForm.PicSelect.Visible = False
End If
ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
BWcancel = True
frmBWiz.applied = False
ActiveForm.Backup
End If
End Sub
Private Sub mnuColInc_Click(Index As Integer)
If freeselection = True Then
    ActiveForm.PicSelect.Picture = ActiveForm.SelHolder.Image
    MyColAdjust ActiveForm.SelHolder, ActiveForm.PicSelect, Index, frmMain.Pb
    Dim i As Integer, j As Integer, R As Long, r2 As Long
    Pb.Max = ActiveForm.PicFreeSelect.Width
    For i = 0 To ActiveForm.PicFreeSelect.Width
        For j = 0 To ActiveForm.PicFreeSelect.Height
            R = GetPixel(ActiveForm.PicFreeSelect.hdc, i, j)
            If R <> 8950944 Then
                r2 = GetPixel(ActiveForm.SelHolder.hdc, i, j)
                SetPixel ActiveForm.PicFreeSelect.hdc, i, j, r2
            End If
        Next j
        Pb.Value = i
    Next i
    TransparentBlt ActiveForm.Pic1BU.hdc, Int(ActiveForm.PicFreeSelect.Left / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.Top / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.ScaleWidth / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.ScaleHeight / ActiveForm.factor), ActiveForm.PicFreeSelect.hdc, 0, 0, ActiveForm.PicFreeSelect.ScaleWidth, ActiveForm.SelHolder.ScaleHeight, 8950944
    ActiveForm.PicMerge.Picture = ActiveForm.Pic1BU.Image
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.PicFreeSelect.Visible = False
    Pb.Value = 0
ElseIf ActiveForm.PicSelect.Visible = True Then
    ActiveForm.PicSelect.Picture = ActiveForm.SelHolder.Image
    MyColAdjust ActiveForm.SelHolder, ActiveForm.PicSelect, Index, frmMain.Pb
    ActiveForm.SelectShape.Visible = False
    StretchBlt ActiveForm.PicMerge.hdc, ActiveForm.SelectLeft, ActiveForm.SelectTop, ActiveForm.Selectwidth, ActiveForm.Selectheight, ActiveForm.SelHolder.hdc, 0, 0, ActiveForm.Selectwidth, ActiveForm.Selectheight, SRCCOPY
    ActiveForm.PicSelect.Visible = False
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.Backup
Else
    ActiveForm.picsource.Picture = ActiveForm.PicMerge.Image
    MyColAdjust ActiveForm.picsource, ActiveForm.PicMerge, Index, frmMain.Pb
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.Backup
End If
End Sub
Private Sub mnuColRed_Click(Index As Integer)
If freeselection = True Then
    ActiveForm.PicSelect.Picture = ActiveForm.SelHolder.Image
    MyColAdjust ActiveForm.SelHolder, ActiveForm.PicSelect, Index + 3, frmMain.Pb
    Dim i As Integer, j As Integer, R As Long, r2 As Long
    Pb.Max = ActiveForm.PicFreeSelect.Width
    For i = 0 To ActiveForm.PicFreeSelect.Width
        For j = 0 To ActiveForm.PicFreeSelect.Height
            R = GetPixel(ActiveForm.PicFreeSelect.hdc, i, j)
            If R <> 8950944 Then
                r2 = GetPixel(ActiveForm.SelHolder.hdc, i, j)
                SetPixel ActiveForm.PicFreeSelect.hdc, i, j, r2
            End If
        Next j
        Pb.Value = i
    Next i
    TransparentBlt ActiveForm.Pic1BU.hdc, Int(ActiveForm.PicFreeSelect.Left / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.Top / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.ScaleWidth / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.ScaleHeight / ActiveForm.factor), ActiveForm.PicFreeSelect.hdc, 0, 0, ActiveForm.PicFreeSelect.ScaleWidth, ActiveForm.SelHolder.ScaleHeight, 8950944
    ActiveForm.PicMerge.Picture = ActiveForm.Pic1BU.Image
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.PicFreeSelect.Visible = False
    Pb.Value = 0
ElseIf ActiveForm.PicSelect.Visible = True Then
    ActiveForm.PicSelect.Picture = ActiveForm.SelHolder.Image
    MyColAdjust ActiveForm.SelHolder, ActiveForm.PicSelect, Index + 3, frmMain.Pb
    ActiveForm.SelectShape.Visible = False
    StretchBlt ActiveForm.PicMerge.hdc, ActiveForm.SelectLeft, ActiveForm.SelectTop, ActiveForm.Selectwidth, ActiveForm.Selectheight, ActiveForm.SelHolder.hdc, 0, 0, ActiveForm.Selectwidth, ActiveForm.Selectheight, SRCCOPY
    ActiveForm.PicSelect.Visible = False
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.Backup
Else
    ActiveForm.picsource.Picture = ActiveForm.PicMerge.Image
    MyColAdjust ActiveForm.picsource, ActiveForm.PicMerge, Index + 3, frmMain.Pb
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.Backup
End If
End Sub
Private Sub mnuEditClip_Click()
If mnuEditClip.Checked = True Then
    mnuEditClip.Checked = False
    UnHookForm Me
Else
    mnuEditClip.Checked = True
    HookForm Me
    SetClipboardViewer Me.Hwnd
End If
End Sub
Private Sub mnuEditPasteNew_Click()
LockWindowUpdate Me.Hwnd
LoadNewDoc
ActiveForm.PicMerge.Picture = Clipboard.GetData(vbCFBitmap)
SizeNew
frmMain.ActiveForm.Picsize
LockWindowUpdate 0
End Sub

Private Sub mnuEditSelectAll_Click()
If ActiveForm.PicFreeSelect.Visible = True Then
    ActiveForm.PicFreeSelect.Visible = False
    ActiveForm.PicMerge.Picture = ActiveForm.Pic1BU.Image
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
End If
dontusePicBU = False
ActiveForm.SelectShape.Left = 0
ActiveForm.SelectShape.Top = 0
ActiveForm.SelectShape.Height = ActiveForm.PicMerge.Height
ActiveForm.SelectShape.Width = ActiveForm.PicMerge.Width
ActiveForm.PicSelect.Left = ActiveForm.SelectShape.Left
ActiveForm.PicSelect.Top = ActiveForm.SelectShape.Top
ActiveForm.PicSelect.Height = ActiveForm.SelectShape.Height
ActiveForm.PicSelect.Width = ActiveForm.SelectShape.Width
ActiveForm.SelectLeft = (ActiveForm.SelectShape.Left) / ActiveForm.factor
ActiveForm.SelectTop = (ActiveForm.SelectShape.Top) / ActiveForm.factor
ActiveForm.Selectwidth = (ActiveForm.SelectShape.Width) / ActiveForm.factor
ActiveForm.Selectheight = (ActiveForm.SelectShape.Height) / ActiveForm.factor
myinternalcopy ActiveForm.PicMerge, ActiveForm.SelHolder, ActiveForm.PicSelectShape
ActiveForm.SelectImage.Left = 0
ActiveForm.SelectImage.Top = 0
ActiveForm.SelectImage.Width = ActiveForm.PicSelect.Width
ActiveForm.SelectImage.Height = ActiveForm.PicSelect.Height
ActiveForm.SelectImage.Picture = ActiveForm.SelHolder.Image
ActiveForm.PicSelectShape.Left = 0
ActiveForm.PicSelectShape.Top = 0
ActiveForm.PicSelectShape.Width = ActiveForm.PicSelect.Width
ActiveForm.PicSelectShape.Height = ActiveForm.PicSelect.Height
ActiveForm.PicSelect.Visible = True
ActiveForm.SelectShape.Visible = False
TB2.Buttons(5).Enabled = True
TB2.Buttons(6).Enabled = True
mnuEditCut.Enabled = True
mnuEditCopy.Enabled = True
ActiveForm.selecting = False

End Sub

Private Sub mnuEditSelectNone_Click()
ActiveForm.PicSelect.Visible = False
If ActiveForm.PicFreeSelect.Visible = True Then
    ActiveForm.PicFreeSelect.Visible = False
    ActiveForm.PicMerge.Picture = ActiveForm.Pic1BU.Image
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
End If
ActiveForm.selecting = False
TB2.Buttons(5).Enabled = False
TB2.Buttons(6).Enabled = False
mnuEditCut.Enabled = False
mnuEditCopy.Enabled = False
mnuEditSelectNone.Enabled = False
End Sub

Private Sub mnuFileAquire_Click()
'Thanks to Stu Lishman
On Error GoTo woops
Dim ScanFile As String, Sc As Integer
Screen.MousePointer = 11
ScanFile = SafeSave(App.path + "\Untitled.bmp")
Sc = TWAIN_AcquireToFilename(Me.Hwnd, ScanFile)
If Sc = 0 Then
    LockWindowUpdate Me.Hwnd
    LoadNewDoc
    ActiveForm.AFCurfile = ScanFile
    ActiveForm.PicMerge.Picture = LoadPicture(ScanFile)
    SizeNew
    ActiveForm.Picsize
    ActiveForm.Backup
    ActiveForm.Backup
    ActiveForm.Caption = FileOnly(ScanFile)
    LockWindowUpdate 0
    Kill ScanFile
Else
  GoTo woops
End If
Screen.MousePointer = 0
Exit Sub
woops:
MsgBox "Operation aborted", vbInformation, "Bobo Enterprises"
Screen.MousePointer = 0

End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
Dim temp As String
Dim frmD As frmImage
temp = MRUfile(Index)
If Not FileExists(temp) Then
    MsgBox "File not found. Perhaps it has been moved or deleted."
    For x = 1 To mnuFileMRU.Count - 1
        Unload mnuFileMRU(x)
    Next x
    mnuFileMRU(0).Visible = False
    mnuFileBar5.Visible = False
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + App.Title, "MRU" + Str(Index)
    LoadMRUs mnuFileMRU(0), mnuFileBar5
    Exit Sub
Else
    LockWindowUpdate Me.Hwnd
    LoadNewDoc
    ActiveForm.AFCurfile = temp
    ActiveForm.PicMerge.Picture = LoadPicture(temp)
    SizeNew
    ActiveForm.Picsize
    ActiveForm.Backup
    ActiveForm.Caption = FileOnly(temp)
End If
    LockWindowUpdate 0
End Sub

Private Sub mnuFilePrinterPrint_Click()
'I dont own a color printer so I haven't tested
'this - I assume it works, please let me
'know if it fails
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.PaintPicture ActiveForm.PicMerge.Picture, 0, 0
Printer.EndDoc

End Sub

Private Sub mnuFilePrinterSetup_Click()
CommonDialog1.ShowPrinter

End Sub

Private Sub mnuFileSaveOptions_Click()
frmPrefs.Show vbModal
End Sub

Private Sub mnuFileTSelect_Click()
TWAIN_SelectImageSource (Me.Hwnd)

End Sub

Private Sub mnuFiltBrow_Click()
Dim chsrcpic As PictureBox
Dim chdestpic As PictureBox
Dim selArea As Boolean
Dim selFreeArea As Boolean
If ActiveForm.PicSelect.Visible = True Then
    Set chdestpic = ActiveForm.SelHolder
    ActiveForm.PicSelect.Picture = chdestpic.Image
    Set chsrcpic = ActiveForm.PicSelect
    selArea = True
ElseIf ActiveForm.PicFreeSelect.Visible = True Then
    Set chdestpic = ActiveForm.SelHolder
    ActiveForm.PicSelect.Picture = chdestpic.Image
    Set chsrcpic = ActiveForm.PicSelect
    selFreeArea = True
Else
    Set chdestpic = ActiveForm.PicMerge
    ActiveForm.picsource.Picture = chdestpic.Image
    Set chsrcpic = ActiveForm.picsource
    selArea = False
End If
frmFiltBrow.Startup chsrcpic
frmFiltBrow.FiltList.ListIndex = 0
frmFiltBrow.Show vbModal
If FBcancel = True Then
    FBcancel = False
    Exit Sub
Else
DoEvents
    If curfilter = 0 Then
        MySharpen chsrcpic, chdestpic, frmMain.Pb, curfilterlevel
    ElseIf curfilter = 1 Then
        MyBlur chsrcpic, chdestpic, frmMain.Pb, curfilterlevel
    ElseIf curfilter = 2 Then
        MyDiffuse chsrcpic, chdestpic, frmMain.Pb, curfilterlevel
    ElseIf curfilter = 3 Then
        MyGreyscale chsrcpic, chdestpic, frmMain.Pb
    ElseIf curfilter = 4 Then
        MyInvert chsrcpic, chdestpic
    ElseIf curfilter = 5 Then
        MyBrightness chsrcpic, chdestpic, curfilterlevel, frmMain.Pb
    ElseIf curfilter = 6 Then
        MyBrightness chsrcpic, chdestpic, -curfilterlevel, frmMain.Pb
    ElseIf curfilter = 7 Then
        MyOutline chsrcpic, chdestpic, frmMain.Pb, curfilterlevel
    ElseIf curfilter = 8 Then
        MyEmboss chsrcpic, chdestpic, frmMain.Pb, curfilterlevel
    ElseIf curfilter = 9 Then
        MyPixelate chsrcpic, chdestpic, frmMain.Pb, curfilterlevel
    End If
If selArea = True Then
    ActiveForm.SelectShape.Visible = False
    StretchBlt ActiveForm.PicMerge.hdc, ActiveForm.SelectLeft, ActiveForm.SelectTop, ActiveForm.Selectwidth, ActiveForm.Selectheight, ActiveForm.SelHolder.hdc, 0, 0, chsrcpic.Width, chsrcpic.Height, SRCCOPY
    ActiveForm.PicSelect.Visible = False
End If
If selFreeArea = True Then
    TransparentBlt ActiveForm.Pic1BU.hdc, Int(ActiveForm.PicFreeSelect.Left / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.Top / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.ScaleWidth / ActiveForm.factor), Int(ActiveForm.PicFreeSelect.ScaleHeight / ActiveForm.factor), ActiveForm.SelHolder.hdc, 0, 0, ActiveForm.SelHolder.ScaleWidth, ActiveForm.SelHolder.ScaleHeight, 8950944
    ActiveForm.PicMerge.Picture = ActiveForm.Pic1BU.Image
    ActiveForm.PicFreeSelect.Visible = False
    Pb.Value = 0
End If
ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
ActiveForm.Backup
End If
End Sub
Private Sub mnuFlipH_Click()
ActiveForm.picsource.Picture = ActiveForm.PicMerge.Image
Flip ActiveForm.PicMerge, ActiveForm.picsource, True
ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
ActiveForm.Backup
End Sub
Private Sub mnuFlipV_Click()
ActiveForm.picsource.Picture = ActiveForm.PicMerge.Image
Flip ActiveForm.PicMerge, ActiveForm.picsource, False
ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
ActiveForm.Backup
End Sub

Private Sub mnuImageProperties_Click()
'This is an easy way to make a thumbnail
If ActiveForm.PicMerge.Height >= ActiveForm.PicMerge.Width Then
    frmImgprops.Image1.Height = 1215
    frmImgprops.Image1.Width = 1215 * (ActiveForm.PicMerge.Width / ActiveForm.PicMerge.Height)
Else
    frmImgprops.Image1.Width = 1215
    frmImgprops.Image1.Height = 1215 * (ActiveForm.PicMerge.Height / ActiveForm.PicMerge.Width)
End If
frmImgprops.Image1.Picture = ActiveForm.PicMerge.Image
If ActiveForm.AFCurfile = "" Then
    frmImgprops.lblInfo(0) = "Untitled.bmp"
    frmImgprops.lblInfo(1) = "Not saved"
    frmImgprops.lblInfo(2) = "Not saved"
Else
    frmImgprops.lblInfo(0) = FileOnly(ActiveForm.AFCurfile)
    frmImgprops.lblInfo(1) = PathOnly(ActiveForm.AFCurfile)
    frmImgprops.lblInfo(2) = FixSize(FileLen(ActiveForm.AFCurfile))
End If
SavePicture ActiveForm.PicMerge.Image, App.path + "\bbtemp.bmp"
frmImgprops.lblInfo(5) = FixSize(FileLen(App.path + "\bbtemp.bmp"))
If FileExists(App.path + "\bbtemp.bmp") Then Kill App.path + "\bbtemp.bmp"
frmImgprops.lblInfo(3) = Trim(Str(ActiveForm.PicMerge.Height)) + " pixels"
frmImgprops.lblInfo(4) = Trim(Str(ActiveForm.PicMerge.Width)) + " pixels"
frmImgprops.Show vbModal
End Sub

Private Sub mnuResize_Click()
AspectRatio = ActiveForm.PicMerge.Width / ActiveForm.PicMerge.Height
frmResize.dontupdate = True
frmResize.VSPXh.Value = frmResize.VSPXh.Max - ActiveForm.PicMerge.ScaleHeight
frmResize.VSPXw.Value = frmResize.VSPXw.Max - ActiveForm.PicMerge.ScaleWidth
frmResize.VSPCh.Value = frmResize.VSPCh.Max - 100
frmResize.VSPCw.Value = frmResize.VSPCw.Max - 100
frmResize.txtPXh = Str(ActiveForm.PicMerge.ScaleHeight)
frmResize.txtPXw = Str(ActiveForm.PicMerge.ScaleWidth)
frmResize.dontupdate = False
frmResize.ChAspRatio.Value = 1
frmResize.Show vbModal
If RScancel = True Then Exit Sub
ActiveForm.PicMerge.AutoSize = True
ActiveForm.picsource1.Height = NewScaleHeight
ActiveForm.picsource1.Width = NewScaleWidth
ActiveForm.picsource1.Picture = LoadPicture()
StretchBlt ActiveForm.picsource1.hdc, 0, 0, ActiveForm.picsource1.Width / Screen.TwipsPerPixelX, ActiveForm.picsource1.Height / Screen.TwipsPerPixelY, ActiveForm.PicMerge.hdc, 0, 0, ActiveForm.PicMerge.Width, ActiveForm.PicMerge.Height, SRCCOPY
If GetTempFile("", "BI", 0, sfilename) Then SavePicture ActiveForm.picsource1.Image, sfilename
ActiveForm.PicMerge.AutoSize = True
ActiveForm.PicMerge.Picture = LoadPicture(sfilename)
Kill sfilename
ActiveForm.PicBG.Width = ActiveForm.PicMerge.Width
ActiveForm.PicBG.Height = ActiveForm.PicMerge.Height
ActiveForm.Image1.Width = ActiveForm.PicMerge.Width
ActiveForm.Image1.Height = ActiveForm.PicMerge.Height
ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
If ActiveForm.WindowState <> 2 Then
    If RulersVis = False Then
        ActiveForm.Width = ActiveForm.PicMerge.ScaleWidth * Screen.TwipsPerPixelX + 120
        ActiveForm.Height = ActiveForm.PicMerge.ScaleHeight * Screen.TwipsPerPixelY + 420
    Else
        ActiveForm.Width = ActiveForm.PicMerge.ScaleWidth * Screen.TwipsPerPixelX + ActiveForm.LeftRuler.Width + 120
        ActiveForm.Height = ActiveForm.PicMerge.ScaleHeight * Screen.TwipsPerPixelY + ActiveForm.TopRuler.Height + 420
    End If
End If
ActiveForm.Picsize
ActiveForm.drawrulers
ActiveForm.Backup
End Sub
Private Sub mnuRotate180_Click()
ActiveForm.picsource.Picture = ActiveForm.PicMerge.Image
Flip ActiveForm.PicMerge, ActiveForm.picsource, False
ActiveForm.picsource.Picture = ActiveForm.PicMerge.Image
Flip ActiveForm.PicMerge, ActiveForm.picsource, True
ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
ActiveForm.Backup
End Sub
Private Sub mnuRotate270_Click()
LockWindowUpdate Me.Hwnd
ActiveForm.picsource.Picture = ActiveForm.PicMerge.Image
Rotate Me.Hwnd, ActiveForm.picsource, ActiveForm.PicMerge, 270
ActiveForm.Image1.Height = ActiveForm.PicMerge.ScaleHeight
ActiveForm.Image1.Width = ActiveForm.PicMerge.ScaleWidth
ActiveForm.PicBG.Height = ActiveForm.PicMerge.Height
ActiveForm.PicBG.Width = ActiveForm.PicMerge.Width
If ActiveForm.WindowState <> 2 Then
    If RulersVis = False Then
        ActiveForm.Width = ActiveForm.PicMerge.ScaleWidth * Screen.TwipsPerPixelX + 120
        ActiveForm.Height = ActiveForm.PicMerge.ScaleHeight * Screen.TwipsPerPixelY + 420
    Else
        ActiveForm.Width = ActiveForm.PicMerge.ScaleWidth * Screen.TwipsPerPixelX + ActiveForm.LeftRuler.Width + 120
        ActiveForm.Height = ActiveForm.PicMerge.ScaleHeight * Screen.TwipsPerPixelY + ActiveForm.TopRuler.Height + 420
    End If
End If
ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
ActiveForm.Picsize
ActiveForm.drawrulers
ActiveForm.Backup

LockWindowUpdate 0
End Sub
Private Sub mnuRotate90_Click()
LockWindowUpdate Me.Hwnd
ActiveForm.picsource.Picture = ActiveForm.PicMerge.Image
Rotate Me.Hwnd, ActiveForm.picsource, ActiveForm.PicMerge, 90
ActiveForm.Image1.Height = ActiveForm.PicMerge.Height
ActiveForm.Image1.Width = ActiveForm.PicMerge.Width
ActiveForm.PicBG.Height = ActiveForm.PicMerge.ScaleHeight
ActiveForm.PicBG.Width = ActiveForm.PicMerge.ScaleWidth
If ActiveForm.WindowState <> 2 Then
    If RulersVis = False Then
        ActiveForm.Width = ActiveForm.PicMerge.ScaleWidth * Screen.TwipsPerPixelX + 120
        ActiveForm.Height = ActiveForm.PicMerge.ScaleHeight * Screen.TwipsPerPixelY + 420
    Else
        ActiveForm.Width = ActiveForm.PicMerge.ScaleWidth * Screen.TwipsPerPixelX + ActiveForm.LeftRuler.Width + 120
        ActiveForm.Height = ActiveForm.PicMerge.ScaleHeight * Screen.TwipsPerPixelY + ActiveForm.TopRuler.Height + 420
    End If
End If
ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
ActiveForm.Picsize
ActiveForm.drawrulers
ActiveForm.Backup
LockWindowUpdate 0
End Sub
Private Sub mnuViewColorBox_Click()
    mnuViewColorBox.Checked = Not mnuViewColorBox.Checked
    ColTB.Visible = mnuViewColorBox.Checked
End Sub
Private Sub mnuViewRulers_Click()
mnuViewRulers.Checked = Not mnuViewRulers.Checked
RulersVis = mnuViewRulers.Checked
End Sub
Private Sub mnuZoomActual_Click()
startwidth = ActiveForm.Image1.Width
LockWindowUpdate ActiveForm.PicBG.Hwnd
ActiveForm.factorlevel = 15
SetLevel
ActiveForm.Image1.Width = ActiveForm.PicMerge.ScaleWidth * ActiveForm.factor
ActiveForm.Image1.Height = ActiveForm.PicMerge.ScaleHeight * ActiveForm.factor
singlefactor = ActiveForm.Image1.Width / startwidth
ActiveForm.PicBG.Width = ActiveForm.Image1.Width
ActiveForm.PicBG.Height = ActiveForm.Image1.Height
SizeSelection
frmMain.ActiveForm.Picsize
frmMain.ActiveForm.drawrulers
LockWindowUpdate 0
End Sub
Private Sub mnuZoomInX_Click(Index As Integer)
Dim notseenV As Boolean
Dim notseenH As Boolean
If Curtool = 4 Then
    Curtool = 2
    TB2.Buttons(2).Value = tbrPressed
End If
If ActiveForm.PicFreeSelect.Visible = True Then
    ActiveForm.PicMerge.Picture = ActiveForm.Pic1BU.Image
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.PicFreeSelect.Visible = False
End If
If ActiveForm.VS.Visible = True Then
    notseenV = False
    startVSval = (ActiveForm.VS.Max + 1) / (ActiveForm.VS.Value + 1)
Else
    notseenV = True
End If
If ActiveForm.HS.Visible = True Then
    notseenH = False
    startHSval = (ActiveForm.HS.Max + 1) / (ActiveForm.HS.Value + 1)
Else
    notseenH = True
End If
startwidth = ActiveForm.Image1.Width
LockWindowUpdate ActiveForm.PicBG.Hwnd
ActiveForm.factorlevel = 15 + Index
SetLevel
ActiveForm.Image1.Width = ActiveForm.PicMerge.ScaleWidth * ActiveForm.factor
ActiveForm.Image1.Height = ActiveForm.PicMerge.ScaleHeight * ActiveForm.factor
singlefactor = ActiveForm.Image1.Width / startwidth
ActiveForm.PicBG.Width = ActiveForm.Image1.Width
ActiveForm.PicBG.Height = ActiveForm.Image1.Height
SizeSelection
frmMain.ActiveForm.Picsize
frmMain.ActiveForm.drawrulers
If notseenV = True Then startVSval = ActiveForm.VS.Max
If notseenH = True Then startHSval = ActiveForm.HS.Max
ActiveForm.VS.Value = ActiveForm.VS.Max / startVSval
ActiveForm.HS.Value = ActiveForm.HS.Max / startHSval
LockWindowUpdate 0
End Sub
Private Sub mnuZoomOutX_Click(Index As Integer)
Dim notseenV As Boolean
Dim notseenH As Boolean
If Curtool = 4 Then
    Curtool = 2
    TB2.Buttons(2).Value = tbrPressed
End If
If ActiveForm.PicFreeSelect.Visible = True Then
    ActiveForm.PicMerge.Picture = ActiveForm.Pic1BU.Image
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.PicFreeSelect.Visible = False
End If
If ActiveForm.VS.Visible = True Then
    notseenV = False
    startVSval = (ActiveForm.VS.Max + 1) / (ActiveForm.VS.Value + 1)
Else
    notseenV = True
End If
If ActiveForm.HS.Visible = True Then
    notseenH = False
    startHSval = (ActiveForm.HS.Max + 1) / (ActiveForm.HS.Value + 1)
Else
    notseenH = True
End If
startwidth = ActiveForm.Image1.Width
LockWindowUpdate ActiveForm.PicBG.Hwnd
ActiveForm.factorlevel = 15 - Index
SetLevel
ActiveForm.Image1.Width = ActiveForm.PicMerge.ScaleWidth * ActiveForm.factor
ActiveForm.Image1.Height = ActiveForm.PicMerge.ScaleHeight * ActiveForm.factor
singlefactor = ActiveForm.Image1.Width / startwidth
ActiveForm.PicBG.Width = ActiveForm.Image1.Width
ActiveForm.PicBG.Height = ActiveForm.Image1.Height
SizeSelection
frmMain.ActiveForm.Picsize
frmMain.ActiveForm.drawrulers
If notseenV = True Then startVSval = ActiveForm.VS.Max
If notseenH = True Then startHSval = ActiveForm.HS.Max
ActiveForm.VS.Value = ActiveForm.VS.Max / startVSval
ActiveForm.HS.Value = ActiveForm.HS.Max / startHSval
LockWindowUpdate 0
End Sub
Private Sub Option1_Click()
For x = 1 To 4
    Toolbar2.Buttons(x).Image = (x + 2) * 2
Next x
Label2.Enabled = True
Slider3.Enabled = True
End Sub
Private Sub Option2_Click()
For x = 1 To 4
    Toolbar2.Buttons(x).Image = (x + 2) * 2 + 1
Next x
Label2.Enabled = False
Slider3.Enabled = False
End Sub
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub
Private Sub mnuWindowNewWindow_Click()
mnuFileNew_Click
End Sub
Private Sub mnuViewSettingsBar_Click()
    mnuViewSettingsBar.Checked = Not mnuViewSettingsBar.Checked
    ToolSettings.Visible = mnuViewSettingsBar.Checked
End Sub
Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    ToolsTB.Visible = mnuViewToolbar.Checked
End Sub
Private Sub mnuEditPaste_Click()
ActiveForm.Enabled = False
MousePointer = 11
Masterpasting = True
If PicClip(ActiveClipboard).Tag = "free" Then freeselection = True
If freeselection = True Then
    dontusePicBU = True
    If ActiveForm.factor = 1 Then
        ActiveForm.SelHolder.Picture = LoadPicture()
        ActiveForm.SelHolder.Picture = Clipboard.GetData(vbCFBitmap)
        ActiveForm.PicFreeSelect.Picture = LoadPicture()
        ActiveForm.PicFreeSelect.Picture = Clipboard.GetData(vbCFBitmap)
        ActiveForm.PicFreeSelect.Left = 0
        ActiveForm.PicFreeSelect.Top = 0
        ActiveForm.ShapeMe 8950944, True, , ActiveForm.PicFreeSelect
        If ActiveForm.VS.Visible = True Then
            ActiveForm.PicFreeSelect.Top = ActiveForm.VS.Value
        End If
        If ActiveForm.HS.Visible = True Then
            ActiveForm.PicFreeSelect.Left = ActiveForm.HS.Value
        End If
        ActiveForm.PicFreeSelect.Visible = True
    Else
        ActiveForm.SelHolder.Picture = Clipboard.GetData(vbCFBitmap)
        ActiveForm.PicFreeSelect.Width = ActiveForm.SelHolder.Width * ActiveForm.factor
        ActiveForm.PicFreeSelect.Height = ActiveForm.SelHolder.Height * ActiveForm.factor
        StretchBlt ActiveForm.PicFreeSelect.hdc, 0, 0, ActiveForm.PicFreeSelect.Width, ActiveForm.PicFreeSelect.Height, ActiveForm.SelHolder.hdc, 0, 0, ActiveForm.SelHolder.Width, ActiveForm.SelHolder.Height, SRCCOPY
        ActiveForm.ShapeMe 8950944, True, , ActiveForm.PicFreeSelect
        ActiveForm.PicFreeSelect.Left = 0
        ActiveForm.PicFreeSelect.Top = 0
        If ActiveForm.VS.Visible = True Then
            ActiveForm.PicFreeSelect.Top = ActiveForm.VS.Value
        End If
        If ActiveForm.HS.Visible = True Then
            ActiveForm.PicFreeSelect.Left = ActiveForm.HS.Value
        End If
        ActiveForm.PicFreeSelect.Visible = True
    End If
Else
    ActiveForm.PicSelect.Left = 0
    ActiveForm.PicSelect.Top = 0
    ActiveForm.SelHolder.Picture = Clipboard.GetData(vbCFBitmap)
    ActiveForm.PicSelect.Width = ActiveForm.SelHolder.Width * ActiveForm.factor
    ActiveForm.PicSelect.Height = ActiveForm.SelHolder.Height * ActiveForm.factor
    ActiveForm.SelectImage.Left = 0
    ActiveForm.SelectImage.Top = 0
    ActiveForm.SelectImage.Width = ActiveForm.PicSelect.Width
    ActiveForm.SelectImage.Height = ActiveForm.PicSelect.Height
    ActiveForm.SelectImage.Picture = ActiveForm.SelHolder.Image
    ActiveForm.PicSelectShape.Left = 0
    ActiveForm.PicSelectShape.Top = 0
    ActiveForm.PicSelectShape.Height = ActiveForm.PicSelect.Height
    ActiveForm.PicSelectShape.Width = ActiveForm.PicSelect.Width
    ActiveForm.PicSelect.Visible = True
End If
ActiveForm.Enabled = True
MousePointer = 0
End Sub
Private Sub mnuEditCopy_Click()
If freeselection = True Then
    Clipboard.SetData ActiveForm.SelHolder.Image
    ActiveForm.PicFreeSelect.Visible = False
    ActiveForm.PicMerge.Picture = ActiveForm.Pic1BU.Image
    ActiveForm.PicMerge.Refresh
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
Else
    ActiveForm.SelHolder.Picture = LoadPicture()
    ActiveForm.SelHolder.Height = ActiveForm.Selectheight
    ActiveForm.SelHolder.Width = ActiveForm.Selectwidth
    StretchBlt ActiveForm.SelHolder.hdc, 0, 0, ActiveForm.Selectwidth, ActiveForm.Selectheight, ActiveForm.PicMerge.hdc, ActiveForm.SelectLeft, ActiveForm.SelectTop, ActiveForm.Selectwidth, ActiveForm.Selectheight, SRCCOPY
    Clipboard.SetData ActiveForm.SelHolder.Image
    ActiveForm.PicSelect.Visible = False
    ActiveForm.SelectShape.Visible = False
End If
End Sub
Private Sub mnuEditCut_Click()
If freeselection = True Then
    Clipboard.SetData ActiveForm.PicFreeSelect.Image
    ActiveForm.PicFreeSelect.Visible = False
Else
    ActiveForm.PicSelect.Picture = LoadPicture()
    ActiveForm.SelHolder.Picture = LoadPicture()
    ActiveForm.SelHolder.Height = ActiveForm.Selectheight
    ActiveForm.SelHolder.Width = ActiveForm.Selectwidth
    ActiveForm.PicSelect.Height = ActiveForm.Selectheight * ActiveForm.factor
    ActiveForm.PicSelect.Width = ActiveForm.Selectwidth * ActiveForm.factor
    StretchBlt ActiveForm.SelHolder.hdc, 0, 0, ActiveForm.Selectwidth, ActiveForm.Selectheight, ActiveForm.PicMerge.hdc, ActiveForm.SelectLeft, ActiveForm.SelectTop, ActiveForm.Selectwidth, ActiveForm.Selectheight, SRCCOPY
    Clipboard.SetData ActiveForm.SelHolder.Image
    StretchBlt ActiveForm.PicMerge.hdc, ActiveForm.SelectLeft, ActiveForm.SelectTop, ActiveForm.Selectwidth, ActiveForm.Selectheight, ActiveForm.PicSelect.hdc, 0, 0, ActiveForm.Selectwidth, ActiveForm.Selectheight, SRCCOPY
    ActiveForm.PicMerge.Refresh
    ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    ActiveForm.PicSelect.Visible = False
    ActiveForm.SelectShape.Visible = False
End If
End Sub
Private Sub mnuEditUndo_Click()
On Error Resume Next
Dim temp As String
temp = ActiveForm.MyUndo
LockWindowUpdate ActiveForm.Hwnd
If FileExists(temp) Then
    ActiveForm.PicMerge.AutoSize = True
    ActiveForm.PicMerge.Picture = LoadPicture(temp)
    If ActiveForm.factor = 1 Then
        SizeNew
        ActiveForm.Picsize
    Else
        ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    End If
Else
    ActiveForm.Blankme
End If
LockWindowUpdate 0
End Sub
Private Sub mnuEditRedo_Click()
On Error Resume Next
Dim temp As String
temp = ActiveForm.MyRedo
LockWindowUpdate ActiveForm.Hwnd
If FileExists(temp) Then
    ActiveForm.PicMerge.AutoSize = True
    ActiveForm.PicMerge.Picture = LoadPicture(temp)
    If ActiveForm.factor = 1 Then
        SizeNew
        ActiveForm.Picsize
    Else
        ActiveForm.Image1.Picture = ActiveForm.PicMerge.Image
    End If
End If
LockWindowUpdate 0
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileSaveAs_Click()
On Error GoTo woops
    Dim sfile As String
    If ActiveForm Is Nothing Then Exit Sub
    With CommonDialog1
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = "Bitmap (*.bmp)|*.bmp|Jpeg (*.jpg)|*.jpg"
        .ShowSave
        If Len(.fileName) = 0 Then Exit Sub
        sfile = .fileName
    If .FilterIndex = 2 Then
        sfile = ChangeExt(sfile, "jpg")
        SavePicture ActiveForm.PicMerge.Image, App.path + "\bbTmpJpg.bmp"
        BmpToJpeg App.path + "\bbTmpJpg.bmp", sfile, Val(GetSetting(App.Title, "Settings", "JPGsaveQuality", "50"))
        If FileExists(App.path + "\bbTmpJpg.bmp") Then Kill App.path + "\bbTmpJpg.bmp"
    Else
        sfile = ChangeExt(sfile, "bmp")
        SavePicture ActiveForm.PicMerge.Image, sfile
    End If
    End With
    ActiveForm.AFCurfile = sfile
    ActiveForm.Caption = FileOnly(sfile)
    UpdateMRUs mnuFileMRU(0), mnuFileBar5, sfile
    ActiveForm.alreadysaved = True
woops:
End Sub
Private Sub mnuFileSave_Click()
If FileExists(ActiveForm.AFCurfile) Then
    If LCase(ExtOnly(ActiveForm.AFCurfile)) = "jpg" Then
        SavePicture ActiveForm.PicMerge.Image, App.path + "\bbTmpJpg.bmp"
        BmpToJpeg App.path + "\bbTmpJpg.bmp", ActiveForm.AFCurfile, Val(GetSetting(App.Title, "Settings", "JPGsaveQuality", "50"))
        If FileExists(App.path + "\bbTmpJpg.bmp") Then Kill App.path + "\bbTmpJpg.bmp"
        ActiveForm.alreadysaved = True
    Else
        If LCase(ExtOnly(ActiveForm.AFCurfile)) = "bmp" Then
            SavePicture ActiveForm.PicMerge.Image, ActiveForm.AFCurfile
            ActiveForm.alreadysaved = True
        Else
            SavePicture ActiveForm.PicMerge.Image, SafeSave(ChangeExt(ActiveForm.AFCurfile, "bmp"))
            ActiveForm.alreadysaved = True
        End If
    End If
Else
    mnuFileSaveAs_Click
End If
End Sub
Private Sub mnuFileClose_Click()
    Unload ActiveForm
End Sub
Private Sub mnuFileOpen_Click()
On Error GoTo woops
Dim sfile As String
With CommonDialog1
    .DialogTitle = "Open"
    .CancelError = False
    .Filter = "Picture files (*.bmp;*.jpg;*.gif;*.bif)|*.bmp;*.jpg;*.gif;*.bif"
    .ShowOpen
    If Len(.fileName) = 0 Then Exit Sub
    sfile = .fileName
End With
LockWindowUpdate Me.Hwnd
LoadNewDoc
ActiveForm.PicMerge.Picture = LoadPicture(sfile)
ActiveForm.AFCurfile = sfile
UpdateMRUs mnuFileMRU(0), mnuFileBar5, sfile
SizeNew
ActiveForm.Picsize
ActiveForm.Caption = FileOnly(sfile)
ActiveForm.Backup
LockWindowUpdate 0
woops:
End Sub
Private Sub mnuFileNew_Click()
On Error Resume Next
Newcancel = False
If lImageCount > 0 Then
    frmNew.VSPXh.Value = frmNew.VSPXh.Max - ActiveForm.PicMerge.ScaleHeight
    frmNew.VSPXw.Value = frmNew.VSPXw.Max - ActiveForm.PicMerge.ScaleWidth
Else
    frmNew.VSPXh.Value = frmNew.VSPXh.Max - 300
    frmNew.VSPXw.Value = frmNew.VSPXw.Max - 300
End If
frmNew.Combo1.ListIndex = CurBGindex
frmNew.Show vbModal
If Newcancel = True Then Exit Sub
NoSizeonStart = True
LoadNewDoc
ActiveForm.Picsize
SizeNew
ActiveForm.AFCurfile = ""
ActiveForm.Backup
End Sub
Public Sub StartupConfig()
    If frmPrefs.Check1.Value = 1 Then
        FloatWindow frmSplash.Hwnd, FLOAT
        frmSplash.Show
    End If
    ReadRgb = True
    Shapetype = 0
    PenTip = 0
    Curtool = 15
    PreviousTool = 15
    TheColor.BackColor = Picture5.BackColor
    LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
    LblCurCol.Refresh
    txtRGB Picture5.BackColor
    If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
    If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
    If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
    TheColor.BackColor = Picture5.BackColor
    LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
    LblCurCol.Refresh
    txtRGB Picture5.BackColor
    If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
    If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
    If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
    Picture2.CurrentX = 3
    Picture2.CurrentY = 3
    Picture2.Print "Tool Settings"
    Picture1.CurrentX = 3
    Picture1.CurrentY = 3
    Picture1.Print "Color Box"
    Combo2.ListIndex = 0
    If frmPrefs.Check3.Value = 1 Then
        mnuEditClip.Checked = True
        Timer2.Enabled = True
        Timer1.Enabled = True
    Else
        Timer1.Enabled = True
        mnuEditClip.Checked = False
        Me.Show
        If Shelled = True Then
            LockWindowUpdate Me.Hwnd
            LoadNewDoc
            ActiveForm.PicMerge.Picture = LoadPicture(myCommand)
            ActiveForm.AFCurfile = myCommand
            UpdateMRUs mnuFileMRU(0), mnuFileBar5, myCommand
            SizeNew
            ActiveForm.Picsize
            ActiveForm.Caption = FileOnly(myCommand)
            ActiveForm.Backup
            LockWindowUpdate 0
        Else
            mnuFileNew_Click
        End If
        cmdTools_Click (0)
        Clipheight = 580
        curClipboard = 0
        cmdTools_Click (4)
    End If
End Sub
Private Sub Timer1_Timer()
Dim tempShell As String
If FileExists(App.path + "\BBdde.tmp") Then
    tempShell = ReadText(App.path + "\BBdde.tmp")
    If FileExists(tempShell) Then
        tempShell = GetLongFilename(tempShell)
        LockWindowUpdate Me.Hwnd
        LoadNewDoc
        ActiveForm.PicMerge.Picture = LoadPicture(tempShell)
        ActiveForm.AFCurfile = tempShell
        UpdateMRUs mnuFileMRU(0), mnuFileBar5, tempShell
        SizeNew
        ActiveForm.Picsize
        ActiveForm.Caption = FileOnly(tempShell)
        ActiveForm.Backup
        LockWindowUpdate 0
        Me.SetFocus
    End If
    Kill App.path + "\BBdde.tmp"
End If
If IsClipboardFormatAvailable(vbCFBitmap) <> 0 Then
    TB2.Buttons(7).Enabled = True
    mnuEditPasteNew.Enabled = True
    If ImageCount > 0 Then mnuEditPaste.Enabled = True
Else
    TB2.Buttons(7).Enabled = False
    mnuEditPaste.Enabled = False
    mnuEditPasteNew.Enabled = False
End If
If ImageCount > 0 Then
    ImgMenuEnable True
    If ActiveForm.PicFreeSelect.Visible = True Or ActiveForm.PicSelect.Visible = True Then
        mnuEditSelectNone.Enabled = True
    End If
    mnuEditSelectAll.Enabled = True
    TB2.Buttons(3).Enabled = True
    If ActiveForm.ListBackup.ListCount = 0 Then
        mnuEditUndo.Enabled = False
        mnuEditRedo.Enabled = False
        TB2.Buttons(9).Enabled = False
        TB2.Buttons(10).Enabled = False
    Else
        If ActiveForm.ListBUorder.ListIndex = 0 Then
            mnuEditUndo.Enabled = False
            TB2.Buttons(9).Enabled = False
            If ActiveForm.ListBUorder.ListCount > 1 Then
                mnuEditRedo.Enabled = True
                TB2.Buttons(10).Enabled = True
            Else
                mnuEditRedo.Enabled = False
                TB2.Buttons(10).Enabled = False
            End If
        ElseIf ActiveForm.ListBUorder.ListIndex > 0 Then
            mnuEditUndo.Enabled = True
            TB2.Buttons(9).Enabled = True
            If ActiveForm.ListBUorder.ListIndex < ActiveForm.ListBUorder.ListCount - 1 Then
                mnuEditRedo.Enabled = True
                TB2.Buttons(10).Enabled = True
            Else
                mnuEditRedo.Enabled = False
                TB2.Buttons(10).Enabled = False
            End If
        End If
    End If
Else
    ImgMenuEnable False
    mnuEditSelectNone.Enabled = False
    mnuEditSelectAll.Enabled = False
    TB2.Buttons(3).Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
starttimer = starttimer + 1
If starttimer = 5 Then
    If Shelled = True Then
        LockWindowUpdate Me.Hwnd
        LoadNewDoc
        ActiveForm.PicMerge.Picture = LoadPicture(myCommand)
        ActiveForm.AFCurfile = myCommand
        UpdateMRUs mnuFileMRU(0), mnuFileBar5, myCommand
        SizeNew
        ActiveForm.Picsize
        ActiveForm.Caption = FileOnly(myCommand)
        ActiveForm.Backup
        LockWindowUpdate 0
    Else
        mnuFileNew_Click
    End If
    cmdTools_Click (0)
    Clipheight = 580
    curClipboard = 0
    HookForm Me
    SetClipboardViewer Me.Hwnd
    Timer2.Enabled = False
    cmdTools_Click (4)
End If
End Sub
Private Sub PicClipBG_Click(Index As Integer)
    ActiveClipboard = Index
    ShClip.Left = PicClipBG(ActiveClipboard).Left - 90
    ShClip.Top = PicClipBG(ActiveClipboard).Top - 90
    DontAdd = True
    Clipboard.Clear
    Clipboard.SetData PicClip(ActiveClipboard).Image
    If PicClip(ActiveClipboard).Tag = "free" Then
        freeselection = True
    Else
        freeselection = False
    End If
    ShClip.Visible = True
    If ChClipLock.Value = 0 Then DontAdd = False
End Sub
Private Sub TB2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim PreviousTool As Integer
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            If ActiveForm.PicSelect.Visible = True Or ActiveForm.PicFreeSelect.Visible = True Then ActiveForm.NeedtoPaste
            mnuEditPaste_Click
        Case "Undo"
            mnuEditUndo_Click
        Case "Redo"
            mnuEditRedo_Click
    End Select
'Here's where we set the variable Curtool so the
'active form knows what to do
If Button.Index > 11 Then Curtool = Button.Index
If Curtool = 19 Then ActiveForm.curpoly = True
If Curtool = 13 Or Curtool = 14 Then
    If PicTools(2).Height < 301 Then cmdTools_Click (2)
End If
If Curtool = 13 Then ActiveForm.PicFreeSelect.Visible = False
If Curtool = 15 Then
    If PicTools(0).Height < 301 Then cmdTools_Click (0)
End If
If Curtool = 20 Then
    If PicTools(1).Height < 301 Then cmdTools_Click (1)
End If
If Curtool = 21 Then
    If PicTools(3).Height < 301 Then cmdTools_Click (3)
End If
If Curtool = 24 Then
    If ActiveForm.lbltext.Visible = True Then ActiveForm.NeedtoPaste
    frmText.Show vbModal
    Curtool = PreviousTool
End If
PreviousTool = Curtool
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Slider1.Enabled = False
        Label1.Enabled = False
        PenTip = 0
    Case 2
        Slider1.Enabled = True
        Label1.Enabled = True
        PenTip = 1
    Case 3
        Slider1.Enabled = True
        Label1.Enabled = True
        PenTip = 2
    Case 4
        Slider1.Enabled = True
        Label1.Enabled = True
        PenTip = 3
    Case 5
        Slider1.Enabled = True
        Label1.Enabled = True
        PenTip = 4
End Select
End Sub
Public Sub LoadClipboard()
'Creates a thumbnail of available clips
Dim tempW As Integer
Dim tempH As Integer
Dim fred As Integer
fred = 0
If IsClipboardFormatAvailable(vbCFBitmap) <> 0 Then
    PicClip(curClipboard).Picture = Clipboard.GetData(vbCFBitmap)
    If freeselection = True Then
        PicClip(curClipboard).Tag = "free"
    Else
        PicClip(curClipboard).Tag = "standard"
    End If
    PicClipBG(curClipboard).Visible = True
    tempW = PicClip(curClipboard).Width
    tempH = PicClip(curClipboard).Height
    If tempW > tempH Then
        ImgClip(curClipboard).Height = 615 * (tempH / tempW)
        ImgClip(curClipboard).Width = 615
        ImgClip(curClipboard).Top = (PicClipBG(curClipboard).Height - ImgClip(curClipboard).Height) / 2
    Else
        ImgClip(curClipboard).Width = 615 * (tempW / tempH)
        ImgClip(curClipboard).Height = 615
        ImgClip(curClipboard).Left = (PicClipBG(curClipboard).Width - ImgClip(curClipboard).Width) / 2
    End If
    ImgClip(curClipboard).Picture = PicClip(curClipboard).Image
    ShClip.Left = PicClipBG(curClipboard).Left - 90
    ShClip.Top = PicClipBG(curClipboard).Top - 90
    ShClip.Visible = True
    curClipboard = curClipboard + 1
    If curClipboard > 5 Then curClipboard = 0
    ActiveClipboard = curClipboard
End If
For x = 0 To 5
    If PicClipBG(x).Visible = True Then fred = fred + 1
Next x
If fred < 1 Then
    Clipheight = 580
    cmdClear.Enabled = False
Else
    cmdClear.Enabled = True
End If
If fred > 0 Then Clipheight = 1500
If fred > 2 Then Clipheight = 2400
If fred > 4 Then Clipheight = 3300
If PicTools(4).Height > 300 Then PicTools(4).Height = Clipheight
End Sub
'All the color handling starts here
Private Sub Picture4_Click()
TheColor.BackColor = Picture4.BackColor
LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
txtRGB Picture4.BackColor
If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
End Sub
Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
picColBlowup.Visible = False
Picture6.BackColor = TheColor.BackColor
If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
End Sub
Private Sub Picture5_Click()
TheColor.BackColor = Picture5.BackColor
LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
txtRGB Picture5.BackColor
If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
End Sub
Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
picColBlowup.Visible = False
Picture6.BackColor = TheColor.BackColor
If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
End Sub
Public Sub loadselected()
Dim x As Integer
colloaded = True
srarthere:
maxcolchose = maxcolchose + 1
If maxcolchose < 79 Then
    If maxcolchose < 2 Then
        coltop = 156
        colleft = 0
    Else
        colleft = colleft + 18
        If colleft > 100 Then
            colleft = 0
            coltop = coltop + 18
        End If
    End If
    For x = colleft To colleft + 15
        For y = coltop To coltop + 15
            SetPixel ColorBox.hdc, x, y, TheColor.BackColor
        Next y
    Next x
    colsave(maxcolchose - 1) = Str(TheColor.BackColor)
Else
    maxcolchose = 0
    GoTo srarthere
End If
If GetTempFile("", "BI", 0, sfilename) Then SavePicture ColorBox.Image, sfilename
ColorBox.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub unloadselected()
maxcolchose = 0
ColorBox.Picture = LoadPicture()
colloaded = False
End Sub
Public Sub OpenColors()
Dim sfile As String, sfile1 As String, sfile2 As String, Z As Integer, col As Long
Dim DialogType As Integer
Dim DialogTitle As String
Dim DialogMsg As String
Dim Response As Integer
DialogType = vbYesNoCancel
DialogTitle = "Bobo Enterprises"
On Error GoTo woops
If colsaved = False Then
    If colloaded = True Then
        DialogMsg = "Do you wish to save the currently selected colors first ?"
        Response = MsgBox(DialogMsg, DialogType, DialogTitle)
            Select Case Response
                Case vbYes
                    SaveColors
                    colsaved = True
                    OpenColors
                    Exit Sub
                Case vbCancel
                    colsaved = False
                    Exit Sub
                Case vbNo
                    colsaved = False
            End Select
    End If
End If
    With CommonDialog1
        .DialogTitle = "Open Color Collection"
        .CancelError = True
        .Filter = "Bobo Color Collection (*.bbc)|*.bbc"
        .ShowOpen
        If Len(.fileName) = 0 Then
            Exit Sub
        End If
        sfile2 = .fileName
    End With
        unloadselected
        LockWindowUpdate ColorBox.Hwnd
        For Z = 0 To 77
            sfile1 = ReadINI(sfile2, "Colors", Str(Z))
            If sfile1 <> "" Then
                TheColor.BackColor = Val(sfile1)
                loadselected
            End If
        Next Z
        LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
        LblCurCol.Refresh
        txtRGB TheColor.BackColor
        LockWindowUpdate 0
        colsaved = False
woops:
colsaved = False
Exit Sub
End Sub
Public Sub SaveColors()
Dim sfile As String, Z As Integer
On Error GoTo woops
    With CommonDialog1
        .DialogTitle = "Save Colors"
        .CancelError = False
        .Filter = "Bobo Color Collection (*.bbc)|*.bbc"
        .ShowSave
        If Len(.fileName) = 0 Then
            Exit Sub
        End If
        sfile = .fileName
    End With
            If colloaded = False Then
                MsgBox "No colors in your collection to save."
                Exit Sub
            End If
          If FileExists(sfile) Then deleteme sfile
        For Z = 0 To 77
            WriteINI sfile, "Colors", Str(Z), colsave(Z)
        Next Z
woops: Exit Sub
End Sub
Public Sub ColorBlowup(col As Long)
'This is the color pallet that appears on the
'right-click of the main color pallet to give
'better color selection
    Dim lCol As Long
    Dim iRed, iGreen, iBlue As Integer
    lCol = col
    iRed = lCol Mod &H100
    lCol = lCol \ &H100
    iGreen = lCol Mod &H100
    lCol = lCol \ &H100
    iBlue = lCol Mod &H100
    Dim intNumber0 As Long, intNumber1 As Long, limit As Integer
    If iRed > 25 Then iRed = iRed - 25 Else iRed = 0
    If iRed + 25 > 255 Then iRed = 205
    If iBlue > 25 Then iBlue = iBlue - 25 Else iBlue = 0
    If iBlue + 25 > 255 Then iBlue = 205
    If iGreen > 25 Then iGreen = iGreen - 25 Else iGreen = 0
    If iGreen + 25 > 255 Then iGreen = 205
    Dim fred As Integer
    fred = 1
    For intNumber0 = 0 To picColBlowup.ScaleHeight
        For intNumber1 = 0 To picColBlowup.ScaleWidth
            SetPixel picColBlowup.hdc, intNumber1, intNumber0, RGB(iRed, iGreen, iBlue)
        Next
        If fred = 1 Then fred = 0 Else fred = 1
        iRed = iRed + fred
        iBlue = iBlue + fred
        iGreen = iGreen + fred
    Next
picColBlowup.Visible = True
End Sub
Private Sub picColBlowup_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
TheColor.BackColor = Picture6.BackColor
LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
txtRGB TheColor.BackColor
If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
If colslocked = True Then Exit Sub
loadselected
End Sub
Private Sub picColBlowup_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static lX As Long, lY As Long
On Local Error Resume Next
Dim H As Long, hD As Long, R As Long
    GetCursorPos P
    If P.x = lX And P.y = lY Then Exit Sub
    lX = P.x: lY = P.y
    H = WindowFromPoint(lX, lY)
    hD = GetDC(H)
    ScreenToClient H, P
    R = GetPixel(hD, P.x, P.y)
    If R = -1 Then
        BitBlt TheColor.hdc, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
        R = TheColor.Point(0, 0)
    Else
        TheColor.PSet (0, 0), R
    End If
    ReleaseDC H, hD
    ChangeColor R
End Sub
Public Sub ChangeColor(lColor As Long)
Picture6.BackColor = lColor
If ReadHex = True Then lblColor = HexRGB(Picture6.BackColor)
If ReadRgb = True Then lblColor = MyRGB(Picture6.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(Picture6.BackColor)
 End Sub
Private Sub ColTB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
picColBlowup.Visible = False
Picture6.BackColor = TheColor.BackColor
If ReadHex = True Then lblColor = HexRGB(Picture6.BackColor)
If ReadRgb = True Then lblColor = MyRGB(Picture6.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(Picture6.BackColor)
End Sub
Private Sub ColorBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim R As Long, cb As Long, cr As Long
If Button <> 2 Then
If picColBlowup.Visible = True Then
    picColBlowup.Visible = False
    Picture6.BackColor = TheColor.BackColor
    If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
    If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
    If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
End If
If y > 155 Then
    R = GetPixel(ColorBox.hdc, x, y)
    cb = GetPixel(ColorBox.hdc, 1, 1)
    If R <> cb Then
        Picture6.BackColor = R
        If ReadHex = True Then lblColor = HexRGB(R)
        If ReadRgb = True Then lblColor = MyRGB(R)
        If ReadLong = True Then lblColor = "LONG:" + Str(R)
    End If
End If
End If
End Sub
Private Sub ColorBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim R As Long, cb As String, cr As String
If Button = 1 Then
    If y > 155 Then
        R = GetPixel(ColorBox.hdc, x, y)
        cb = "RGB: 192, 192, 192"
        cr = MyRGB(R)
        If cr <> cb Then
            Picture6.BackColor = R
            If ReadHex = True Then lblColor = HexRGB(R)
            If ReadRgb = True Then lblColor = MyRGB(R)
            If ReadLong = True Then lblColor = "LONG:" + Str(R)
            TheColor.BackColor = Picture6.BackColor
            LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
            LblCurCol.Refresh
            txtRGB Picture6.BackColor
        End If
    End If
ElseIf Button = 2 Then
    If y > 155 Then
        R = GetPixel(ColorBox.hdc, x, y)
        cb = "RGB: 192, 192, 192"
        cr = MyRGB(R)
        If cr <> cb Then
            picColBlowup.Left = x - picColBlowup.Width / 2
            If picColBlowup.Left > ColorBox.Width - picColBlowup.Width Then picColBlowup.Left = ColorBox.Width - picColBlowup.Width
            If picColBlowup.Left < 0 Then picColBlowup.Left = 0
            picColBlowup.Top = y - (picColBlowup.Height - picColBlowup.Height / 10)
            ColorBlowup (R)
        End If
    End If
End If
End Sub
Private Sub lblColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    If Left(lblColor, 3) = "HEX" Then
        ReadHex = False
        ReadRgb = True
        ReadLong = False
        lblColor = MyRGB(TheColor.BackColor)
        Exit Sub
    End If
    If Left(lblColor, 3) = "RGB" Then
        ReadHex = False
        ReadRgb = False
        ReadLong = True
        lblColor = "LONG:" + Str(TheColor.BackColor)
        Exit Sub
    End If
    If Left(lblColor, 3) = "LON" Then
        ReadHex = True
        ReadRgb = False
        ReadLong = False
        lblColor = HexRGB(TheColor.BackColor)
        Exit Sub
    End If
ElseIf Button = vbRightButton Then
End If
End Sub
Private Sub lblColor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
picColBlowup.Visible = False
Picture6.BackColor = TheColor.BackColor
If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    TheColor.BackColor = Picture6.BackColor
    LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
    LblCurCol.Refresh
    txtRGB Picture6.BackColor
    If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
    If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
    If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
    If colslocked = True Then Exit Sub
    loadselected
Else
    Static lX As Long, lY As Long
    On Local Error Resume Next
    Dim H As Long, hD As Long, R As Long
        GetCursorPos P
        If P.x = lX And P.y = lY Then Exit Sub
        lX = P.x: lY = P.y
        H = WindowFromPoint(lX, lY)
        hD = GetDC(H)
        ScreenToClient H, P
        R = GetPixel(hD, P.x, P.y)
        If R = -1 Then
            BitBlt TheColor.hdc, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
            R = TheColor.Point(0, 0)
        Else
            TheColor.PSet (0, 0), R
        End If
        ReleaseDC H, hD
        picColBlowup.Left = Image1.Left + x / Screen.TwipsPerPixelX - picColBlowup.Width / 2
        If picColBlowup.Left > ColorBox.Width - picColBlowup.Width Then picColBlowup.Left = ColorBox.Width - picColBlowup.Width
        If picColBlowup.Left < 0 Then picColBlowup.Left = 0
        picColBlowup.Top = Image1.Top + y / Screen.TwipsPerPixelY - picColBlowup.Height / 10
        ColorBlowup (R)
End If
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Then
    Static lX As Long, lY As Long
    On Local Error Resume Next
    Dim H As Long, hD As Long, R As Long
        GetCursorPos P
        If P.x = lX And P.y = lY Then Exit Sub
        lX = P.x: lY = P.y
        H = WindowFromPoint(lX, lY)
        hD = GetDC(H)
        ScreenToClient H, P
        R = GetPixel(hD, P.x, P.y)
        If R = -1 Then
            BitBlt TheColor.hdc, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
            R = TheColor.Point(0, 0)
        Else
            TheColor.PSet (0, 0), R
        End If
        ReleaseDC H, hD
        ChangeColor R
        picColBlowup.Visible = False
End If
End Sub
Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Shapetype = Button.Index - 1
End Sub
Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo woops
Select Case Button.Index
    Case 1
        If colslocked = True Then
            colslocked = False
            Toolbar3.Buttons(1).Image = 2
        Else
            colslocked = True
            Toolbar3.Buttons(1).Image = 1
        End If
    Case 2
        colslocked = False
    Case 3
        If MsgBox("Are you sure you wish to remove your collection of colors ?", vbYesNo) = vbYes Then unloadselected
    Case 4
        Dim NewColor As Long
        NewColor = ShowColor
         If NewColor <> -1 Then
            TheColor.BackColor = NewColor
            LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
            LblCurCol.Refresh
            txtRGB NewColor
            If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
            If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
            If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
        Else
            GoTo woops
        End If
    Case 5
        OpenColors
    Case 6
        SaveColors
End Select
woops:
Exit Sub
End Sub

Private Sub txtHue_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtHue_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtHue.Text) > 255 Then txtHue.Text = Str(255)
strInt = Val(TrimVoid(txtHue.Text))
VShue.Value = VShue.Max - strInt
End Sub
Private Sub txtLum_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtLum.Text) > 255 Then txtLum.Text = Str(255)
strInt = Val(TrimVoid(txtLum.Text))
VSlum.Value = VSlum.Max - strInt
End Sub
Private Sub txtLum_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtSat_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtSat_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtSat.Text) > 255 Then txtSat.Text = Str(255)
strInt = Val(TrimVoid(txtHue.Text))
VSsat.Value = VSsat.Max - strInt
End Sub
Private Sub txtBlue_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtBlue_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtBlue.Text) > 255 Then txtBlue.Text = Str(255)
strInt = Val(TrimVoid(txtBlue.Text))
VSblue.Value = VSblue.Max - strInt
End Sub
Private Sub txtGreen_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtGreen_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtGreen.Text) > 255 Then txtGreen.Text = Str(255)
strInt = Val(TrimVoid(txtGreen.Text))
VSgreen.Value = VSgreen.Max - strInt
End Sub
Private Sub txtRed_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtRed_KeyUp(KeyCode As Integer, Shift As Integer)
If Val(txtRed.Text) > 255 Then txtRed.Text = Str(255)
strInt = Val(TrimVoid(txtRed.Text))
VSred.Value = VSred.Max - strInt
End Sub
Private Sub VSred_Change()
txtRed.Text = TrimVoid(Str(VSred.Max - VSred.Value))
UpdateColScrolls
End Sub
Private Sub VSred_Scroll()
txtRed.Text = TrimVoid(Str(VSred.Max - VSred.Value))
UpdateColScrolls
End Sub
Private Sub VSblue_Change()
txtBlue.Text = TrimVoid(Str(VSblue.Max - VSblue.Value))
UpdateColScrolls
End Sub
Private Sub VSblue_Scroll()
txtBlue.Text = TrimVoid(Str(VSblue.Max - VSblue.Value))
UpdateColScrolls
End Sub
Private Sub VSgreen_Change()
txtGreen.Text = TrimVoid(Str(VSgreen.Max - VSgreen.Value))
UpdateColScrolls
End Sub
Private Sub VSgreen_Scroll()
txtGreen.Text = TrimVoid(Str(VSgreen.Max - VSgreen.Value))
UpdateColScrolls
End Sub
Private Sub VShue_Change()
txtHue.Text = TrimVoid(Str(VShue.Max - VShue.Value))
UpdateHSLscrolls
End Sub
Private Sub VShue_Scroll()
txtHue.Text = TrimVoid(Str(VShue.Max - VShue.Value))
UpdateHSLscrolls
End Sub
Private Sub VSsat_Change()
txtSat.Text = TrimVoid(Str(VSsat.Max - VSsat.Value))
UpdateHSLscrolls
End Sub
Private Sub VSsat_Scroll()
txtSat.Text = TrimVoid(Str(VSsat.Max - VSsat.Value))
UpdateHSLscrolls
End Sub
Private Sub VSlum_Change()
txtLum.Text = TrimVoid(Str(VSlum.Max - VSlum.Value))
UpdateHSLscrolls
End Sub
Private Sub VSlum_Scroll()
txtLum.Text = TrimVoid(Str(VSlum.Max - VSlum.Value))
UpdateHSLscrolls
End Sub
Public Sub txtRGB(lCdlColor As Long)
    dontupdateVS = True
    Dim lCol As Long
    Dim iRed, iGreen, iBlue As Integer
    lCol = lCdlColor
    iRed = lCol Mod &H100
    lCol = lCol \ &H100
    iGreen = lCol Mod &H100
    lCol = lCol \ &H100
    iBlue = lCol Mod &H100
    txtRed.Text = TrimVoid(Str(iRed))
    txtGreen.Text = TrimVoid(Str(iGreen))
    txtBlue.Text = TrimVoid(Str(iBlue))
    VSred.Value = VSred.Max - iRed
    VSgreen.Value = VSgreen.Max - iGreen
    VSblue.Value = VSblue.Max - iBlue
    VShue.Value = VShue.Max - RGBtoHSL(lCdlColor).Hue
    VSsat.Value = VSsat.Max - RGBtoHSL(lCdlColor).Sat
    VSlum.Value = VSlum.Max - RGBtoHSL(lCdlColor).Lum
    dontupdateVS = False
End Sub
Public Sub txtRGB2(lCdlColor As Long)
    dontupdateVS = True
    Dim lCol As Long
    Dim iRed, iGreen, iBlue As Integer
    lCol = lCdlColor
    iRed = lCol Mod &H100
    lCol = lCol \ &H100
    iGreen = lCol Mod &H100
    lCol = lCol \ &H100
    iBlue = lCol Mod &H100
    txtRed.Text = TrimVoid(Str(iRed))
    txtGreen.Text = TrimVoid(Str(iGreen))
    txtBlue.Text = TrimVoid(Str(iBlue))
    VSred.Value = VSred.Max - iRed
    VSgreen.Value = VSgreen.Max - iGreen
    VSblue.Value = VSblue.Max - iBlue
    dontupdateVS = False
End Sub
Public Sub UpdateColScrolls()
Picture6.BackColor = RGB(Val(txtRed.Text), Val(txtGreen.Text), Val(txtBlue.Text))
TheColor.BackColor = Picture6.BackColor
LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
If dontupdateVS = False Then
    dontupdateVS = True
    VShue.Value = VShue.Max - RGBtoHSL(Picture6.BackColor).Hue
    VSsat.Value = VSsat.Max - RGBtoHSL(Picture6.BackColor).Sat
    VSlum.Value = VSlum.Max - RGBtoHSL(Picture6.BackColor).Lum
    dontupdateVS = False
End If
End Sub
Public Sub UpdateHSLscrolls()
HSLV.Hue = Val(txtHue.Text)
HSLV.Sat = Val(txtSat.Text)
HSLV.Lum = Val(txtLum.Text)
If dontupdateVS = False Then
    TheColor.BackColor = Picture6.BackColor
    LblCurCol.ForeColor = ContrastingColor(TheColor.BackColor)
    LblCurCol.Refresh
    If ReadHex = True Then lblColor = HexRGB(TheColor.BackColor)
    If ReadRgb = True Then lblColor = MyRGB(TheColor.BackColor)
    If ReadLong = True Then lblColor = "LONG:" + Str(TheColor.BackColor)
    Picture6.BackColor = HSLtoRGB(HSLV)
    txtRGB2 Picture6.BackColor
End If
End Sub
Public Sub LoadMRUs(mnuFileMRU As Control, mnufileMRUSpace As Control)
Dim temp As String
Dim fred As Integer
fred = 0
For x = 0 To MRUlimit - 1
    temp = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + App.Title, "MRU" + Str(x), "")
    If FileExists(temp) Then
    If fred = 0 Then
        Me.mnuFileMRU(0).Visible = True
        Me.mnuFileMRU(0).Caption = FileOnly(temp)
        MRUpath(0) = temp
        Me.mnuFileBar5.Visible = True
    Else
        Load Me.mnuFileMRU(fred)
        Me.mnuFileMRU(fred).Visible = True
        Me.mnuFileMRU(fred).Caption = FileOnly(temp)
        MRUpath(fred) = temp
    End If
    fred = fred + 1
    End If
Next x
End Sub
Public Function MRUfile(Index As Integer) As String
If FileExists(MRUpath(Index)) Then
    MRUfile = MRUpath(Index)
Else
    MsgBox FileOnly(MRUpath(Index)) + " has been moved or deleted.", vbExclamation
    MRUfile = ""
End If
End Function
Public Sub UpdateMRUs(mnuFileMRU As Control, mnufileMRUSpace As Control, NewMRU As String)
Dim x As Integer
For x = 0 To MRUlimit - 1
    If MRUpath(x) = NewMRU Then
        Exit Sub
    End If
Next x
For x = 1 To Me.mnuFileMRU.Count - 1
    Unload Me.mnuFileMRU(x)
Next x
For x = MRUlimit - 1 To 1 Step -1
    MRUpath(x) = MRUpath(x - 1)
Next x
MRUpath(0) = NewMRU
Me.mnuFileMRU(0).Visible = True
Me.mnuFileMRU(0).Caption = FileOnly(NewMRU)
Me.mnuFileBar5.Visible = True
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + App.Title, "MRU" + Str(0), MRUpath(0)
For x = 1 To MRUlimit - 1
    If MRUpath(x) <> "" Then
        Load Me.mnuFileMRU(x)
        Me.mnuFileMRU(x).Visible = True
        Me.mnuFileMRU(x).Caption = FileOnly(MRUpath(x))
    Else
        Exit For
    End If
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\" + App.Title, "MRU" + Str(x), MRUpath(x)
Next x
End Sub
Public Sub MySaveAs()
mnuFileSave_Click
End Sub


Public Sub ImgMenuEnable(mEn As Boolean)
If mEn = False Then
    mnuAddBorders.Enabled = False
    mnuImageProperties.Enabled = False
    mnuFlipH.Enabled = False
    mnuFlipV.Enabled = False
    mnuRotate90.Enabled = False
    mnuRotate180.Enabled = False
    mnuRotate270.Enabled = False
    mnuResize.Enabled = False
    mnuBorderWiz.Enabled = False
    mnuFiltBrow.Enabled = False
    mnuIncCol.Enabled = False
    mnuRedCol.Enabled = False
    mnuZoomIn.Enabled = False
    mnuZoomOut.Enabled = False
    mnuWindowCascade.Enabled = False
    mnuWindowTileHorizontal.Enabled = False
    mnuWindowTileVertical.Enabled = False
    mnuFileClose.Enabled = False
    mnuFileSave.Enabled = False
    mnuFileSaveAs.Enabled = False
    TB2.Buttons(22).Enabled = False
    TB2.Buttons(23).Enabled = False
    TB2.Buttons(24).Enabled = False
Else
    mnuAddBorders.Enabled = True
    mnuImageProperties.Enabled = True
    mnuFlipH.Enabled = True
    mnuFlipV.Enabled = True
    mnuRotate90.Enabled = True
    mnuRotate180.Enabled = True
    mnuRotate270.Enabled = True
    mnuResize.Enabled = True
    mnuBorderWiz.Enabled = True
    mnuFiltBrow.Enabled = True
    mnuIncCol.Enabled = True
    mnuRedCol.Enabled = True
    mnuZoomIn.Enabled = True
    mnuZoomOut.Enabled = True
    mnuWindowCascade.Enabled = True
    mnuWindowTileHorizontal.Enabled = True
    mnuWindowTileVertical.Enabled = True
    mnuFileClose.Enabled = True
    mnuFileSave.Enabled = True
    mnuFileSaveAs.Enabled = True
    TB2.Buttons(22).Enabled = True
    TB2.Buttons(23).Enabled = True
    TB2.Buttons(24).Enabled = True
End If
End Sub
