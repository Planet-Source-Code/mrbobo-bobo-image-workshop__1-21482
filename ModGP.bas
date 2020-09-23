Attribute VB_Name = "ModGP"
Global Curtool As Integer
Global FBcancel As Boolean
Global BWcancel As Boolean
Global RScancel As Boolean
Global CWcancel As Boolean
Global saveCancel As Boolean
Global RulersVis As Boolean
Global finalclose As Boolean
Global Newcancel As Boolean
Global NewHeight As Long
Global NewWidth As Long
Global NewBGcol As Long
Global CurBGindex As Integer
Global curfilter As Integer
Global curfilterlevel As Integer
Global curborder As Integer
Global curborderlevel2 As Integer
Global curborderlevel3 As Integer
Global borderwidth As Integer
Global framewidth As Integer
Global chBGcolor As Long
Global outline As Boolean
Global inline As Boolean
Global AspectRatio As Double
Global NewScaleHeight As Long
Global NewScaleWidth As Long
Global Savepath As String
Global curfile As String
Global ReadLong As Boolean
Global ReadHex As Boolean
Global ReadRgb As Boolean
Global maxcolchose As Integer
Global colslocked As Boolean
Global pastingasnew As Boolean
Global startVSval As Double
Global startHSval As Double
Global freeselection As Boolean
Global dontusePicBU As Boolean
Global Masterpasting As Boolean
Global P As POINTAPI
Global sfilename As String
Global Cachesize As Long
Public safesavename As String
Private Type POINTAPI
    x As Long
    y As Long
End Type
Public Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal Nwidth As Long, ByVal Nheight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function BmpToJpeg Lib "BBJPeg.dll" (ByVal bmpFileName As String, ByVal JpegFilename As String, ByVal Quality As Integer) As Integer
Declare Function TWAIN_AcquireToFilename Lib "EZTW32.DLL" (ByVal hwndApp%, ByVal bmpFileName$) As Integer
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long
Declare Function TWAIN_IsAvailable Lib "EZTW32.DLL" () As Long
Declare Function TWAIN_EasyVersion Lib "EZTW32.DLL" () As Long
Public ret As String
Public Retlen As String
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long
Private Type SHFILEOPSTRUCT
    Hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
    End Type
    Private Const FO_MOVE = &H1
    Private Const FO_COPY = &H2
    Private Const FOF_SILENT = &H4
    Private Const FOF_RENAMEONCOLLISION = &H8
    Private Const FOF_NOCONFIRMATION = &H10
    Private Const FOF_SIMPLEPROGRESS = &H100
    Private Const FOF_ALLOWUNDO = &H40
    Private Const FO_DELETE = &H3
   Private Const FO_RENAME = &H4&
Dim flag As Integer
Dim fred As Integer
Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim FO_FUNC As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const FLOAT = 1, SINK = 0
Public Declare Sub SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Public Const GWL_HWNDPARENT = (-8)
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public ImageCount As Integer
Public PenTip As Integer
Public Shapetype As Integer
Public PenDrawWidth As Integer
Public PenTipWidth As Integer
Global startwidth As Integer
Global singlefactor As Double
Global NoSizeonStart As Boolean
Public Sub LoadNewDoc()
    Dim frmD As Form
    ImageCount = ImageCount + 1
    Set frmD = New frmImage
    frmD.Caption = "Image " & Str(ImageCount)
    frmD.Show
End Sub
Sub FloatWindow(x As Long, action As Integer)
Dim wFlags As Integer, result As Integer
wFlags = SWP_NOMOVE Or SWP_NOSIZE
If action <> 0 Then
    Call SetWindowPos(x, HWND_TOPMOST, 0, 0, 0, 0, wFlags)
Else
    Call SetWindowPos(x, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags)
End If
End Sub
Public Function temppath() As String
    Dim sBuffer As String
    Dim lRet As Long
    sBuffer = String$(255, vbNullChar)
    lRet = GetTempPath(255, sBuffer)
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    temppath = sBuffer
    If Right(temppath, 1) = "\" Then temppath = Left(temppath, Len(temppath) - 1)
End Function
'Used to ensure a unique filename and thus
'avoid overwriting
Public Function SafeSave(path As String) As String
Dim mPath As String, mTemp As String, mFile As String, mExt As String, m As Integer
On Error Resume Next
mPath = Mid$(path, 1, InStrRev(path, "\"))
mname = Mid$(path, InStrRev(path, "\") + 1)
mFile = Left(Mid$(mname, 1, InStrRev(mname, ".")), Len(Mid$(mname, 1, InStrRev(mname, "."))) - 1) 'File only - no extension
If mFile = "" Then mFile = mname
mExt = Mid$(mname, InStrRev(mname, "."))
mTemp = ""
Do
    If Not FileExists(mPath + mFile + mTemp + mExt) Then
        SafeSave = mPath + mFile + mTemp + mExt
        safesavename = mFile + mTemp + mExt
        Exit Do
    End If
    m = m + 1
    mTemp = Right(Str(m), Len(Str(m)) - 1)
Loop
End Function
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
Private Function PerformShellAction(sSource As String, sDestination As String) As Long
      On Error Resume Next
    sSource = sSource & Chr$(0) & Chr$(0)
     FOF_FLAGS = BuildBrowseFlags()
     With SHFileOp
        .wFunc = FO_FUNC
        .pFrom = sSource
        .pTo = sDestination
        .fFlags = FOF_FLAGS
    End With
    PerformShellAction = SHFileOperation(SHFileOp)
End Function
Public Sub RenameFile(fileName As String, Target As String)
    On Error Resume Next
    Dim FileStruct As SHFILEOPSTRUCT
    Dim P As Boolean
    Dim x As Long
    Dim strNoConfirm As Integer, strNoConfirmMakeDir As Integer, strRenameOnCollision As Integer
    Dim strSilent As Integer, strSimpleProgress As Integer
         FileStruct.pFrom = fileName
        FileStruct.pTo = Target
        FileStruct.wFunc = FO_RENAME
        x = SHFileOperation(FileStruct)
  End Sub
Private Function BuildBrowseFlags() As Long
     On Error Resume Next
     flag = flag Or FOF_SILENT
        flag = flag Or FOF_NOCONFIRMATION
   If fred = 1 Then flag = flag Or FOF_RENAMEONCOLLISION
 BuildBrowseFlags = flag
End Function
Private Sub ShellDeleteOne(sfile As String)
     On Error Resume Next
   Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
Dim R As Long
    FOF_FLAGS = BuildBrowseFlags()
sfile = sfile & Chr$(0)
With SHFileOp
  .wFunc = FO_DELETE
  .pFrom = sfile
  .fFlags = FOF_FLAGS
End With
R = SHFileOperation(SHFileOp)
End Sub
Public Sub moveme(Source As String, dest As String)
FO_FUNC = 1
Call PerformShellAction(Source, dest)
End Sub
Public Sub CopyMe(Source As String, dest As String)
FO_FUNC = 2
Call PerformShellAction(Source, dest)
End Sub
Public Sub deleteme(path As String)
ShellDeleteOne (path)
End Sub
Public Function FileOnly(ByVal FilePath As String) As String
    FileOnly = Mid$(FilePath, InStrRev(FilePath, "\") + 1)
End Function
Public Function ExtOnly(ByVal FilePath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(FilePath, InStrRev(FilePath, ".") + 1)
If dot = True Then ExtOnly = "." + ExtOnly
End Function
Public Function ChangeExt(ByVal FilePath As String, Optional newext As String) As String
Dim temp As String
temp = Mid$(FilePath, 1, InStrRev(FilePath, "."))
temp = Left(temp, Len(temp) - 1)
If newext <> "" Then newext = "." + newext
ChangeExt = temp + newext
End Function
Public Function PathOnly(ByVal FilePath As String) As String
Dim temp As String
    temp = Mid$(FilePath, 1, InStrRev(FilePath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Public Function labeledit(Destination As String, Length As Integer) As String
Dim y As Integer, m As Integer, temp As String, temp1 As String, temp2 As String, temp3 As String
If Len(Destination) > Length Then
    m = 0
    For y = Len(Destination) To 1 Step -1
        m = m + 1
        If Mid(Destination, y, 1) = "\" Then
            temp2 = Right(Destination, m)
            Exit For
        End If
    Next y
    m = 0
    For y = 4 To Len(Destination)
        m = m + 1
        If Mid(Destination, y, 1) = "\" Then
            temp1 = Left(Destination, m + 3)
            Exit For
        End If
    Next y
    If Len(temp1 + temp2) < Length Then
doagain:
    m = Len(temp1) + 1
    For y = Len(temp1) + 2 To Len(Destination)
        m = m + 1
        If Mid(Destination, y, 1) = "\" Then
            temp = Left(Destination, m)
            Exit For
        End If
    Next y
     If Len(temp + temp2) < Length Then
     temp1 = temp
     GoTo doagain
     Else
     GoTo OKdone
     End If
     Else
     temp1 = Left(Destination, 3)
     End If
OKdone:
        m = Length - Len(temp1 + temp2)
        temp3 = "."
        For y = 1 To m
            temp3 = temp3 + "."
        Next y
    labeledit = temp1 + temp3 + temp2
Else
    labeledit = Destination
End If
End Function
Public Function TrimVoid(Expre)
  On Error Resume Next
  Dim i As Integer
  Dim beg As String
  Dim expr As String
  For i = 1 To Len(Expre)
        beg = Mid(Expre, i, 1)
        If beg Like "[a-zA-Z0-9]" Then expr = expr & beg
    Next
    TrimVoid = expr
End Function
Public Sub WriteINI(fileName As String, Section As String, Key As String, Text As String)
WritePrivateProfileString Section, Key, Text, fileName
End Sub
Public Function ReadINI(fileName As String, Section As String, Key As String)
ret = Space$(255)
Retlen = GetPrivateProfileString(Section, Key, "", ret, Len(ret), fileName)
ret = Left$(ret, Retlen)
ReadINI = ret
End Function
Public Function HexRGB(lCdlColor As Long)
    Dim lCol As Long
    Dim iRed, iGreen, iBlue As Integer
    Dim vHexR, vHexG, vHexB As Variant
    lCol = lCdlColor
    iRed = lCol Mod &H100
    lCol = lCol \ &H100
    iGreen = lCol Mod &H100
    lCol = lCol \ &H100
    iBlue = lCol Mod &H100
    vHexR = Hex(iRed)
    If Len(vHexR) < 2 Then
        vHexR = "0" & vHexR
    End If
    vHexG = Hex(iGreen)
    If Len(vHexG) < 2 Then
        vHexG = "0" & iGreen
    End If
    vHexB = Hex(iBlue)
    If Len(vHexB) < 2 Then
        vHexB = "0" & vHexB
    End If
    HexRGB = "HEX:" + "#" & vHexR & vHexG & vHexB
End Function
Public Function MyRGB(lCdlColor As Long)
    Dim lCol As Long
    Dim iRed, iGreen, iBlue As Integer
    lCol = lCdlColor
    iRed = lCol Mod &H100
    lCol = lCol \ &H100
    iGreen = lCol Mod &H100
    lCol = lCol \ &H100
    iBlue = lCol Mod &H100
    MyRGB = "RGB:" + Str(iRed) + "," + Str(iGreen) + "," + Str(iBlue)
End Function
Public Sub myinternalcopy(Spic As PictureBox, Dpic As PictureBox, TempselShape As Shape)
Dpic.Width = frmMain.ActiveForm.Selectwidth
Dpic.Height = frmMain.ActiveForm.Selectheight
Dpic.Picture = LoadPicture()
StretchBlt Dpic.hdc, 0, 0, frmMain.ActiveForm.Selectwidth, frmMain.ActiveForm.Selectheight, Spic.hdc, frmMain.ActiveForm.SelectLeft, frmMain.ActiveForm.SelectTop, frmMain.ActiveForm.Selectwidth, frmMain.ActiveForm.Selectheight, SRCCOPY
Dpic.Picture = Dpic.Image
Dpic.Refresh
End Sub
Public Sub PasteInSitu(Spic As PictureBox, Dpic As PictureBox)
StretchBlt Dpic.hdc, SelectLeft, SelectTop, Selectwidth, Selectheight, Spic.hdc, 0, 0, Selectwidth, Selectheight, SRCCOPY
Spic.Visible = False
End Sub
Public Sub ClearPostDrag(Spic As PictureBox, Dpic As PictureBox)
Spic.Width = Selectwidth
Spic.Height = Selectheight
Spic.BackColor = vbWhite
StretchBlt Dpic.hdc, SelectLeft, SelectTop, Selectwidth, Selectheight, Spic.hdc, 0, 0, Selectwidth, Selectheight, SRCCOPY
End Sub
Public Sub mycopy(Spic As PictureBox, Dpic As PictureBox, TempselShape As Shape)
Dpic.Width = Selectwidth
Dpic.Height = Selectheight
Dpic.Top = SelectTop
Dpic.Left = SelectLeft
Dpic.Visible = True
TempselShape.Left = 1
TempselShape.Top = 1
TempselShape.Width = Dpic.Width - 2
TempselShape.Height = Dpic.Height - 2
TempselShape.Visible = True
Dpic.Picture = LoadPicture()
StretchBlt Dpic.hdc, 0, 0, Selectwidth, Selectheight, Spic.hdc, Dpic.Left, Dpic.Top, Selectwidth, Selectheight, SRCCOPY
Dpic.Refresh
Clipboard.SetData Dpic.Image
End Sub
Public Sub mypaste(Spic As PictureBox, Dpic As PictureBox)
frmMain.ActiveForm.SelectPic.Picture = Clipboard.GetData(vbCFBitmap)
Spic.Left = 0
Spic.Top = 0
Spic.Visible = True
Spic.ZOrder 0
End Sub
Public Function ShowColor() As Long
            On Error GoTo woops
        frmMain.CommonDialog1.CancelError = True
        frmMain.CommonDialog1.Flags = 0
        frmMain.CommonDialog1.action = 3
        ShowColor = frmMain.CommonDialog1.Color
woops:
End Function
Public Function GetTempFile2(lpTempFilename As String, path As String) As Boolean
    lpTempFilename = String(255, vbNullChar)
    GetTempFile2 = GetTempFilename(path, "bb", 0, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function
Public Sub SizeNew()
frmMain.ActiveForm.PicBG.Width = frmMain.ActiveForm.PicMerge.Width
frmMain.ActiveForm.PicBG.Height = frmMain.ActiveForm.PicMerge.Height
frmMain.ActiveForm.Image1.Width = frmMain.ActiveForm.PicMerge.Width
frmMain.ActiveForm.Image1.Height = frmMain.ActiveForm.PicMerge.Height
frmMain.ActiveForm.Image1.Picture = frmMain.ActiveForm.PicMerge.Image
If frmMain.ActiveForm.WindowState <> 2 Then
    If frmMain.ActiveForm.Width > frmMain.ActiveForm.PicBG.Width * Screen.TwipsPerPixelX + 120 Then frmMain.ActiveForm.Width = frmMain.ActiveForm.PicBG.Width * Screen.TwipsPerPixelX + 120
    If frmMain.ActiveForm.Height > frmMain.ActiveForm.PicBG.Height * Screen.TwipsPerPixelY + 520 Then frmMain.ActiveForm.Height = frmMain.ActiveForm.PicBG.Height * Screen.TwipsPerPixelY + 520
End If
End Sub
Public Sub SetLevel()
For x = 0 To 15
    frmMain.mnuZoomInX(x).Checked = False
    frmMain.mnuZoomOutX(x).Checked = False
Next x
If frmMain.ActiveForm.factorlevel = 0 Then frmMain.ActiveForm.factor = 1 / 16
If frmMain.ActiveForm.factorlevel = 1 Then frmMain.ActiveForm.factor = 1 / 15
If frmMain.ActiveForm.factorlevel = 2 Then frmMain.ActiveForm.factor = 1 / 14
If frmMain.ActiveForm.factorlevel = 3 Then frmMain.ActiveForm.factor = 1 / 13
If frmMain.ActiveForm.factorlevel = 4 Then frmMain.ActiveForm.factor = 1 / 12
If frmMain.ActiveForm.factorlevel = 5 Then frmMain.ActiveForm.factor = 1 / 11
If frmMain.ActiveForm.factorlevel = 6 Then frmMain.ActiveForm.factor = 1 / 10
If frmMain.ActiveForm.factorlevel = 7 Then frmMain.ActiveForm.factor = 1 / 9
If frmMain.ActiveForm.factorlevel = 8 Then frmMain.ActiveForm.factor = 1 / 8
If frmMain.ActiveForm.factorlevel = 9 Then frmMain.ActiveForm.factor = 1 / 7
If frmMain.ActiveForm.factorlevel = 10 Then frmMain.ActiveForm.factor = 1 / 6
If frmMain.ActiveForm.factorlevel = 11 Then frmMain.ActiveForm.factor = 1 / 5
If frmMain.ActiveForm.factorlevel = 12 Then frmMain.ActiveForm.factor = 1 / 4
If frmMain.ActiveForm.factorlevel = 13 Then frmMain.ActiveForm.factor = 1 / 3
If frmMain.ActiveForm.factorlevel = 14 Then frmMain.ActiveForm.factor = 1 / 2
If frmMain.ActiveForm.factorlevel = 15 Then frmMain.ActiveForm.factor = 1
If frmMain.ActiveForm.factorlevel = 16 Then frmMain.ActiveForm.factor = 2 / 1
If frmMain.ActiveForm.factorlevel = 17 Then frmMain.ActiveForm.factor = 3 / 1
If frmMain.ActiveForm.factorlevel = 18 Then frmMain.ActiveForm.factor = 4 / 1
If frmMain.ActiveForm.factorlevel = 19 Then frmMain.ActiveForm.factor = 5 / 1
If frmMain.ActiveForm.factorlevel = 20 Then frmMain.ActiveForm.factor = 6 / 1
If frmMain.ActiveForm.factorlevel = 21 Then frmMain.ActiveForm.factor = 7 / 1
If frmMain.ActiveForm.factorlevel = 22 Then frmMain.ActiveForm.factor = 8 / 1
If frmMain.ActiveForm.factorlevel = 23 Then frmMain.ActiveForm.factor = 9 / 1
If frmMain.ActiveForm.factorlevel = 24 Then frmMain.ActiveForm.factor = 10 / 1
If frmMain.ActiveForm.factorlevel = 25 Then frmMain.ActiveForm.factor = 11 / 1
If frmMain.ActiveForm.factorlevel = 26 Then frmMain.ActiveForm.factor = 12 / 1
If frmMain.ActiveForm.factorlevel = 27 Then frmMain.ActiveForm.factor = 13 / 1
If frmMain.ActiveForm.factorlevel = 28 Then frmMain.ActiveForm.factor = 14 / 1
If frmMain.ActiveForm.factorlevel = 29 Then frmMain.ActiveForm.factor = 15 / 1
If frmMain.ActiveForm.factorlevel = 30 Then frmMain.ActiveForm.factor = 16 / 1
If frmMain.ActiveForm.factorlevel < 15 Then
    frmMain.mnuZoomOutX(15 - frmMain.ActiveForm.factorlevel).Checked = True
    frmMain.mnuZoomActual.Enabled = True
ElseIf frmMain.ActiveForm.factorlevel = 15 Then
    frmMain.mnuZoomInX(0).Checked = True
    frmMain.mnuZoomOutX(0).Checked = True
    frmMain.mnuZoomActual.Enabled = False
ElseIf frmMain.ActiveForm.factorlevel > 15 Then
    frmMain.mnuZoomActual.Enabled = True
    frmMain.mnuZoomInX(frmMain.ActiveForm.factorlevel - 15).Checked = True
End If
End Sub
Public Sub Zoom(Direction As Boolean, Img As Image, Srcpic As PictureBox)
Dim startwidth As Integer
Dim singlefactor As Double
startwidth = Img.Width
LockWindowUpdate frmMain.ActiveForm.PicFrame.Hwnd
If Direction = False Then
    If frmMain.ActiveForm.factorlevel < 1 Then Exit Sub
    frmMain.ActiveForm.factorlevel = frmMain.ActiveForm.factorlevel - 1
Else
    If frmMain.ActiveForm.factorlevel > 29 Then Exit Sub
    frmMain.ActiveForm.factorlevel = frmMain.ActiveForm.factorlevel + 1
End If
SetLevel
Img.Width = Srcpic.ScaleWidth * frmMain.ActiveForm.factor
Img.Height = Srcpic.ScaleHeight * frmMain.ActiveForm.factor
singlefactor = Img.Width / startwidth
frmMain.ActiveForm.PicBG.Width = Img.Width
frmMain.ActiveForm.PicBG.Height = Img.Height
frmMain.ActiveForm.PicSelect.Top = frmMain.ActiveForm.PicSelect.Top * singlefactor
frmMain.ActiveForm.PicSelect.Left = frmMain.ActiveForm.PicSelect.Left * singlefactor
frmMain.ActiveForm.PicSelect.Height = frmMain.ActiveForm.PicSelect.Height * singlefactor
frmMain.ActiveForm.PicSelect.Width = frmMain.ActiveForm.PicSelect.Width * singlefactor
frmMain.ActiveForm.PicSelectShape.Left = 0
frmMain.ActiveForm.PicSelectShape.Top = 0
frmMain.ActiveForm.PicSelectShape.Width = frmMain.ActiveForm.PicSelect.Width
frmMain.ActiveForm.PicSelectShape.Height = frmMain.ActiveForm.PicSelect.Height
frmMain.ActiveForm.Picsize
LockWindowUpdate 0
End Sub
Public Sub SizeSelection()
    frmMain.ActiveForm.PicSelect.Top = frmMain.ActiveForm.PicSelect.Top * singlefactor
    frmMain.ActiveForm.PicSelect.Left = frmMain.ActiveForm.PicSelect.Left * singlefactor
    frmMain.ActiveForm.PicSelect.Height = frmMain.ActiveForm.PicSelect.Height * singlefactor
    frmMain.ActiveForm.PicSelect.Width = frmMain.ActiveForm.PicSelect.Width * singlefactor
    frmMain.ActiveForm.PicSelectShape.Left = 0
    frmMain.ActiveForm.PicSelectShape.Top = 0
    frmMain.ActiveForm.PicSelectShape.Width = frmMain.ActiveForm.PicSelect.Width
    frmMain.ActiveForm.PicSelectShape.Height = frmMain.ActiveForm.PicSelect.Height
    frmMain.ActiveForm.SelectImage.Left = 0
    frmMain.ActiveForm.SelectImage.Top = 0
    frmMain.ActiveForm.SelectImage.Width = frmMain.ActiveForm.PicSelect.Width
    frmMain.ActiveForm.SelectImage.Height = frmMain.ActiveForm.PicSelect.Height
    If frmMain.ActiveForm.PicBG.Left + frmMain.ActiveForm.PicBG.Width < frmMain.ActiveForm.PicFrame.ScaleWidth Then
        frmMain.ActiveForm.PicBG.Left = frmMain.ActiveForm.PicFrame.ScaleWidth - frmMain.ActiveForm.PicBG.Width
        If frmMain.ActiveForm.PicBG.Left < 0 Then frmMain.ActiveForm.HS.Value = -frmMain.ActiveForm.PicBG.Left
        If frmMain.ActiveForm.PicBG.Width < frmMain.ActiveForm.PicFrame.ScaleWidth Then frmMain.ActiveForm.PicBG.Left = 0
    End If
    If frmMain.ActiveForm.PicBG.Top + frmMain.ActiveForm.PicBG.Height < frmMain.ActiveForm.PicFrame.ScaleHeight Then
        frmMain.ActiveForm.PicBG.Top = frmMain.ActiveForm.PicFrame.ScaleHeight - frmMain.ActiveForm.PicBG.Height
        If frmMain.ActiveForm.PicBG.Top < 0 Then frmMain.ActiveForm.VS.Value = -frmMain.ActiveForm.PicBG.Top
        If frmMain.ActiveForm.PicBG.Height < frmMain.ActiveForm.PicFrame.ScaleHeight Then frmMain.ActiveForm.PicBG.Top = 0
    End If
End Sub
Public Sub EnableMenus()
    frmMain.mnuFlipV.Enabled = True
    frmMain.mnuFlipH.Enabled = True
    frmMain.mnuRotate90.Enabled = True
    frmMain.mnuRotate180.Enabled = True
    frmMain.mnuRotate270.Enabled = True
    frmMain.mnuFiltBrow.Enabled = True
    frmMain.mnuIncCol.Enabled = True
    frmMain.mnuRedCol.Enabled = True
    frmMain.mnuFileSave.Enabled = True
    frmMain.mnuBorderWiz.Enabled = True
    frmMain.mnuAddBorders.Enabled = True
    frmMain.mnuResize.Enabled = True
    frmMain.mnuFileSaveAs.Enabled = True
    frmMain.mnuFileClose.Enabled = True
    frmMain.mnuImageProperties.Enabled = True
    frmMain.mnuWindowCascade.Enabled = True
    frmMain.mnuWindowTileHorizontal.Enabled = True
    frmMain.mnuWindowTileVertical.Enabled = True
End Sub
Public Sub DisableMenus(Optional full As Boolean)
    frmMain.mnuFlipV.Enabled = False
    frmMain.mnuFlipH.Enabled = False
    frmMain.mnuRotate90.Enabled = False
    frmMain.mnuRotate180.Enabled = False
    frmMain.mnuRotate270.Enabled = False
    frmMain.mnuFiltBrow.Enabled = False
    frmMain.mnuIncCol.Enabled = False
    frmMain.mnuRedCol.Enabled = False
    frmMain.mnuFileSave.Enabled = False
    frmMain.mnuFileSaveAs.Enabled = False
    frmMain.mnuFileClose.Enabled = False
    frmMain.mnuImageProperties.Enabled = False
    If full Then
        frmMain.mnuBorderWiz.Enabled = False
        frmMain.mnuAddBorders.Enabled = False
        frmMain.mnuResize.Enabled = False
        frmMain.mnuEditUndo.Enabled = False
        frmMain.mnuEditRedo.Enabled = False
        frmMain.mnuWindowCascade.Enabled = False
        frmMain.mnuWindowTileHorizontal.Enabled = False
        frmMain.mnuWindowTileVertical.Enabled = False
        frmMain.TB2.Buttons(3).Enabled = False
        frmMain.TB2.Buttons(9).Enabled = False
        frmMain.TB2.Buttons(10).Enabled = False
        full = False
    End If
End Sub

Public Function FixSize(mSize As Long) As String
Dim totsize As Double
Dim strsize As String
Dim m As Integer, y As Integer
Dim temp As String, temp1 As String, temp2 As String
If mSize > 1024 Then
    totsize = mSize / 1024
    strsize = Str(totsize)
    temp2 = " Kb"
    If totsize > 1024 Then
        totsize = totsize / 1024
        strsize = Str(totsize)
        temp2 = " Mb"
    End If
      m = 0
     For y = Len(strsize) To 1 Step -1
     m = m + 1
     If Mid(strsize, y, 1) = "." Then
      temp1 = Right(strsize, m)
      temp1 = Left(temp1, 3)
      temp = Left(strsize, Len(strsize) - m)
      Exit For
     End If
    Next y
    strsize = temp + temp1 + temp2
Else
    strsize = Str(mSize) + " Bytes"
    
End If
FixSize = strsize
End Function

Public Function ReadText(path As String) As String
    Dim Line
    Dim temptxt As String
    temptxt = ""
    Open path For Input As #1
    Do While Not EOF(1)
        Input #1, Line
        temptxt = temptxt + Line
    Loop
    Close #1
    ReadText = temptxt
End Function


