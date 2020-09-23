VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPrefs 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferences"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmPrefs.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Limits"
      TabPicture(1)   =   "frmPrefs.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "JPEG"
      TabPicture(2)   =   "frmPrefs.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   4695
         Begin VB.ComboBox Combo2 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmPrefs.frx":0054
            Left            =   2640
            List            =   "frmPrefs.frx":006A
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Unlimited Undo/Redo buffer"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   840
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmPrefs.frx":0085
            Left            =   2640
            List            =   "frmPrefs.frx":0098
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "steps."
            Enabled         =   0   'False
            Height          =   255
            Left            =   3600
            TabIndex        =   19
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Limit Undo/Redo buffer to"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   17
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "files."
            Height          =   255
            Left            =   3600
            TabIndex        =   15
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Limit recently opened file list to"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   4695
         Begin VB.CheckBox Check4 
            Caption         =   "Dont ask to save changes on exit"
            Height          =   195
            Left            =   360
            TabIndex        =   12
            Top             =   1440
            Width           =   2895
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Clipboard Monitor Active at Startup"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Associate with .bmp,.jpg and .gif files"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   2895
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Show Splash Screen at Startup"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Value           =   1  'Checked
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "JPEG Save Quality"
         Height          =   1215
         Left            =   -74640
         TabIndex        =   3
         Top             =   720
         Width           =   4455
         Begin MSComctlLib.Slider Slider1 
            Height          =   255
            Left            =   960
            TabIndex        =   4
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   393216
            Min             =   10
            Max             =   90
            SelStart        =   50
            TickFrequency   =   5
            Value           =   50
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "SMALLEST"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "BEST QUALITY"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   3480
            TabIndex        =   5
            Top             =   480
            Width           =   720
         End
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "OK"
      Height          =   330
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check5_Click()
If Check5.Value = 1 Then
    Label5.Enabled = False
    Label5.Enabled = False
    Combo2.Enabled = False
Else
    Label5.Enabled = True
    Label5.Enabled = True
    Combo2.Enabled = True
End If
End Sub

Private Sub Command1_Click()
SaveSetting App.Title, "Settings", "JPGsaveQuality", Trim(Str(Slider1.Value))
SaveSetting App.Title, "Settings", "ShowSplash", Trim(Str(Check1.Value))
SaveSetting App.Title, "Settings", "Associate", Trim(Str(Check2.Value))
SaveSetting App.Title, "Settings", "Clipmon", Trim(Str(Check3.Value))
SaveSetting App.Title, "Settings", "SavePrompt", Trim(Str(Check4.Value))
SaveSetting App.Title, "Settings", "Undoredo", Trim(Str(Check5.Value))
SaveSetting App.Title, "Settings", "MRUlist", Trim(Str(Combo1.ListIndex))
SaveSetting App.Title, "Settings", "UndoSteps", Trim(Str(Combo2.ListIndex))
frmMain.MRUlimit = Val(Combo1.Text)
If Check2.Value = 1 Then
    Associate App.path + "\ImageWorkshop.exe", ".bmp"
    Associate App.path + "\ImageWorkshop.exe", ".jpg"
    Associate App.path + "\ImageWorkshop.exe", ".gif"
Else
    temp = GetSettingString(HKEY_CLASSES_ROOT, ".jpg", "")
    If temp = "ImageWorkshop.exe" Then DeleteValue HKEY_CLASSES_ROOT, ".jpg", "IELite.exe"
    temp = GetSettingString(HKEY_CLASSES_ROOT, ".gif", "")
    If temp = "ImageWorkshop.exe" Then DeleteValue HKEY_CLASSES_ROOT, ".gif", "IELite.exe"
    temp = GetSettingString(HKEY_CLASSES_ROOT, ".jpeg", "")
    If temp = "ImageWorkshop.exe" Then DeleteValue HKEY_CLASSES_ROOT, ".bmp", "IELite.exe"
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Slider1.Value = Val(GetSetting(App.Title, "Settings", "JPGsaveQuality", "50"))
Check1.Value = Val(GetSetting(App.Title, "Settings", "ShowSplash", "1"))
Check2.Value = Val(GetSetting(App.Title, "Settings", "Associate", "0"))
Check3.Value = Val(GetSetting(App.Title, "Settings", "Clipmon", "1"))
Check4.Value = Val(GetSetting(App.Title, "Settings", "SavePrompt", "0"))
Check5.Value = Val(GetSetting(App.Title, "Settings", "Undoredo", "1"))
Combo1.ListIndex = Val(GetSetting(App.Title, "Settings", "MRUlist", "1"))
Combo2.ListIndex = Val(GetSetting(App.Title, "Settings", "UndoSteps", "0"))

End Sub
