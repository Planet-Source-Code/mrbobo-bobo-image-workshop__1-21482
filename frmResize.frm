VERSION 5.00
Begin VB.Form frmResize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resize Image"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   Icon            =   "frmResize.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      Caption         =   "Percentage"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pixel Size"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000016&
      Caption         =   "OK"
      Height          =   330
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CheckBox ChAspRatio 
      Caption         =   "Maintain aspect ratio"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   2070
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
      Begin VB.VScrollBar VSPCh 
         Height          =   220
         Left            =   3240
         Min             =   1
         TabIndex        =   16
         Top             =   390
         Value           =   10
         Width           =   220
      End
      Begin VB.VScrollBar VSPCw 
         Height          =   220
         Left            =   1320
         Min             =   1
         TabIndex        =   15
         Top             =   390
         Value           =   1
         Width           =   220
      End
      Begin VB.TextBox txtPCh 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtPCw 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "x  Height"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Width"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.VScrollBar VSPXh 
         Height          =   220
         Left            =   3240
         TabIndex        =   14
         Top             =   390
         Width           =   220
      End
      Begin VB.VScrollBar VSPXw 
         Height          =   220
         Left            =   1320
         TabIndex        =   13
         Top             =   390
         Width           =   220
      End
      Begin VB.TextBox txtPXh 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtPXw 
         BackColor       =   &H80000016&
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "x  Height"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Width"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'just interface stuff here the action happens
'back on frmMain
Public dontupdate As Boolean
Dim newpxwidth As Integer
Dim tsize As Long
Dim oldtxtw As String
Dim oldtxth As String
Dim oldtxtPCw As String
Dim oldtxtPCh As String
Private Sub ChAspRatio_Click()
If ChAspRatio.Value = 1 Then
    VSPXh.Value = VSPXh.Max - (Val(txtPXw.Text) / AspectRatio)
   VSPCh.Value = VSPCw.Value
End If
End Sub
Private Sub cmdCancel_Click()
RScancel = True
Unload Me
End Sub
Private Sub cmdOK_Click()
RScancel = False
NewScaleHeight = Val(txtPXh.Text) * Screen.TwipsPerPixelY
NewScaleWidth = Val(txtPXw.Text) * Screen.TwipsPerPixelX
If NewScaleHeight > 32000 Or NewScaleWidth > 32000 Then
    MsgBox "This is much too big !", vbCritical
    Exit Sub
End If
Unload Me
End Sub
Private Sub Option1_Click()
Frame2.Enabled = False
txtPCw.BackColor = &H80000004
txtPCh.BackColor = &H80000004
txtPCw.Enabled = False
txtPCh.Enabled = False
Frame1.Enabled = True
txtPXw.BackColor = &H80000016
txtPXh.BackColor = &H80000016
txtPXw.Enabled = True
txtPXh.Enabled = True
End Sub
Private Sub Option2_Click()
Frame2.Enabled = True
txtPCw.BackColor = &H80000016
txtPCh.BackColor = &H80000016
txtPCw.Enabled = True
txtPCh.Enabled = True
Frame1.Enabled = False
txtPXw.BackColor = &H80000004
txtPXh.BackColor = &H80000004
txtPXw.Enabled = False
txtPXh.Enabled = False
End Sub
Private Sub txtPCh_Change()
If Len(txtPCh) > 3 Then
    txtPCh = oldtxtPCh
End If
End Sub
Private Sub txtPCh_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtPCh_KeyUp(KeyCode As Integer, Shift As Integer)
tsize = ((Val(txtPCh) * frmMain.ActiveForm.PicMerge.ScaleHeight / 100) * AspectRatio)
If tsize * Screen.TwipsPerPixelY > 32000 Then tsize = 32000 / Screen.TwipsPerPixelY
If ChAspRatio.Value = 1 Then
    updateall tsize
Else
    VSPXh.Value = VSPXh.Max - ((Val(txtPCh.Text) * frmMain.ActiveForm.PicMerge.ScaleHeight) / 100)
End If
oldtxtPCh = txtPCh
End Sub
Private Sub txtPCw_Change()
If Len(txtPCw) > 3 Then
    txtPCw = oldtxtPCw
End If
End Sub
Private Sub txtPCw_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtPCw_KeyUp(KeyCode As Integer, Shift As Integer)
tsize = ((Val(txtPCw) * frmMain.ActiveForm.PicMerge.ScaleWidth) / 100)
If tsize * Screen.TwipsPerPixelY > 32000 Then tsize = 32000 / Screen.TwipsPerPixelY
If ChAspRatio.Value = 1 Then
    updateall tsize
Else
    VSPXw.Value = VSPXw.Max - ((Val(txtPCw.Text) * frmMain.ActiveForm.PicMerge.ScaleWidth) / 100)
End If
oldtxtPCw = txtPCw
End Sub
Private Sub txtPXh_Change()
If Len(txtPXh) > 4 Then
    txtPXh = oldtxth
End If
End Sub
Private Sub txtPXh_KeyPress(KeyAscii As Integer)
If Len(txtPXh) > 4 Then
    KeyAscii = 0
    txtPXh = oldtxtw
    Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtPXh_KeyUp(KeyCode As Integer, Shift As Integer)
tsize = (Val(txtPXh) * AspectRatio)
If tsize * Screen.TwipsPerPixelY > 32000 Then tsize = 32000 / Screen.TwipsPerPixelY
If ChAspRatio.Value = 1 Then
    updateall tsize
Else
    VSPCh.Value = VSPCh.Max - ((Val(txtPXh.Text) / frmMain.ActiveForm.PicMerge.ScaleHeight) * 100)
End If
oldtxth = txtPXh
End Sub
Private Sub txtPXw_Change()
If Len(txtPXw) > 4 Then
    txtPXw = oldtxtw
End If
End Sub
Private Sub txtPXw_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub
Private Sub txtPXw_KeyUp(KeyCode As Integer, Shift As Integer)
tsize = (Val(txtPXw))
If tsize * Screen.TwipsPerPixelY > 32000 Then tsize = 32000 / Screen.TwipsPerPixelY
If ChAspRatio.Value = 1 Then
    updateall tsize
Else
    VSPCw.Value = VSPCw.Max - ((Val(txtPXw.Text) / frmMain.ActiveForm.PicMerge.ScaleWidth) * 100)
End If
oldtxtw = txtPXw
End Sub
Private Sub VSPCh_Change()
    If ChAspRatio.Value = 0 Then
        txtPCh.Text = TrimVoid(Str(VSPCh.Max - VSPCh.Value))
        VSPXh.Value = VSPXh.Max - ((Val(txtPCh.Text) * frmMain.ActiveForm.PicMerge.ScaleHeight) / 100)
    Else
        txtPCh.Text = TrimVoid(Str(VSPCh.Max - VSPCh.Value))
        If dontupdate = False Then
            newpxwidth = ((VSPCh.Max - VSPCh.Value) * frmMain.ActiveForm.PicMerge.ScaleHeight / 100) * AspectRatio
            updateall (newpxwidth)
        End If
    End If
End Sub
Private Sub VSPCh_Scroll()
    If ChAspRatio.Value = 0 Then
        txtPCh.Text = TrimVoid(Str(VSPCh.Max - VSPCh.Value))
        VSPXh.Value = VSPXh.Max - ((Val(txtPCh.Text) * frmMain.ActiveForm.PicMerge.ScaleHeight) / 100)
    Else
        txtPCh.Text = TrimVoid(Str(VSPCh.Max - VSPCh.Value))
        If dontupdate = False Then
            newpxwidth = ((VSPCh.Max - VSPCh.Value) * frmMain.ActiveForm.PicMerge.ScaleHeight / 100) * AspectRatio
            updateall (newpxwidth)
        End If
    End If
End Sub
Private Sub VSPCw_Change()
    If ChAspRatio.Value = 0 Then
        txtPCw.Text = TrimVoid(Str(VSPCw.Max - VSPCw.Value))
        VSPXw.Value = VSPXw.Max - ((Val(txtPCw.Text) * frmMain.ActiveForm.PicMerge.ScaleWidth) / 100)
    Else
        txtPCw.Text = TrimVoid(Str(VSPCw.Max - VSPCw.Value))
        If dontupdate = False Then
            newpxwidth = ((VSPCw.Max - VSPCw.Value) * frmMain.ActiveForm.PicMerge.ScaleWidth / 100)
            updateall (newpxwidth)
        End If
    End If
End Sub
Private Sub VSPCw_Scroll()
    If ChAspRatio.Value = 0 Then
        txtPCw.Text = TrimVoid(Str(VSPCw.Max - VSPCw.Value))
        VSPXw.Value = VSPXw.Max - ((Val(txtPCw.Text) * frmMain.ActiveForm.PicMerge.ScaleWidth) / 100)
    Else
        txtPCw.Text = TrimVoid(Str(VSPCw.Max - VSPCw.Value))
        If dontupdate = False Then
            newpxwidth = ((VSPCw.Max - VSPCw.Value) * frmMain.ActiveForm.PicMerge.ScaleWidth / 100)
            updateall (newpxwidth)
        End If
    End If
End Sub
Private Sub VSPXh_Change()
    If ChAspRatio.Value = 0 Then
        txtPXh.Text = TrimVoid(Str(VSPXh.Max - VSPXh.Value))
        VSPCh.Value = VSPCh.Max - ((Val(txtPXh.Text) / frmMain.ActiveForm.PicMerge.ScaleHeight) * 100)
    Else
        txtPXh.Text = TrimVoid(Str(VSPXh.Max - VSPXh.Value))
        If dontupdate = False Then
            newpxwidth = (VSPXh.Max - VSPXh.Value) * AspectRatio
            updateall (newpxwidth)
        End If
    End If
End Sub
Private Sub VSPXh_scroll()
    If ChAspRatio.Value = 0 Then
        txtPXh.Text = TrimVoid(Str(VSPXh.Max - VSPXh.Value))
        VSPCh.Value = VSPCh.Max - ((Val(txtPXh.Text) / frmMain.ActiveForm.PicMerge.ScaleHeight) * 100)
    Else
        txtPXh.Text = TrimVoid(Str(VSPXh.Max - VSPXh.Value))
        If dontupdate = False Then
            newpxwidth = (VSPXh.Max - VSPXh.Value) * AspectRatio
            updateall (newpxwidth)
        End If
    End If
End Sub
Private Sub VSPXw_Change()
    If ChAspRatio.Value = 0 Then
        txtPXw.Text = TrimVoid(Str(VSPXw.Max - VSPXw.Value))
        VSPCw.Value = VSPCw.Max - ((Val(txtPXw.Text) / frmMain.ActiveForm.PicMerge.ScaleWidth) * 100)
    Else
        txtPXw.Text = TrimVoid(Str(VSPXw.Max - VSPXw.Value))
        If dontupdate = False Then
            newpxwidth = VSPXw.Max - VSPXw.Value
            updateall (newpxwidth)
        End If
    End If
End Sub
Private Sub VSPXw_scroll()
    If ChAspRatio.Value = 0 Then
        txtPXw.Text = TrimVoid(Str(VSPXw.Max - VSPXw.Value))
        VSPCw.Value = VSPCw.Max - ((Val(txtPXw.Text) / frmMain.ActiveForm.PicMerge.ScaleWidth) * 100)
    Else
        txtPXw.Text = TrimVoid(Str(VSPXw.Max - VSPXw.Value))
        If dontupdate = False Then
            newpxwidth = VSPXw.Max - VSPXw.Value
            updateall (newpxwidth)
        End If
    End If
End Sub
Public Sub updateall(myval As Long)
dontupdate = True
    VSPXw.Value = VSPXw.Max - myval
    VSPCw.Value = VSPCw.Max - (myval / frmMain.ActiveForm.PicMerge.ScaleWidth * 100)
    VSPXh.Value = VSPXw.Max - (myval / AspectRatio)
    VSPCh.Value = VSPCw.Max - (myval / frmMain.ActiveForm.PicMerge.ScaleWidth * 100)
dontupdate = False
End Sub
