VERSION 5.00
Begin VB.Form frmImgprops 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Image Properties"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmImgprops.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "OK"
      Height          =   330
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Current File Size :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   12
         Top             =   840
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   11
         Top             =   1440
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   10
         Top             =   1140
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   9
         Top             =   540
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Top             =   2490
         Width           =   45
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   2040
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Image Width :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Image Height :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Original File Size :"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Path :"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   2490
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Title :"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   390
      End
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmImgprops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

