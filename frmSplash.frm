VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4125
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   3895
      Left            =   120
      ScaleHeight     =   3840
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Image Image1 
         Height          =   3855
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   2280
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'My wife and son
Dim gg As Integer

Private Sub Form_Load()
Image1.Picture = Me.Picture
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
gg = gg + 1
If gg = 3 Then
Timer1.Enabled = False
Unload Me
End If
End Sub
