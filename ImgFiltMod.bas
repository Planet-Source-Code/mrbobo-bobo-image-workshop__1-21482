Attribute VB_Name = "ImgFiltMod"
'Each fuctions name describes its function
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Dim Red As Long
Dim Green As Long
Dim Blue As Long
Dim larrCol() As Long
Dim R As Long, G As Long, b As Long, tempval As Long
Dim i As Integer, j As Integer
Dim pixvalue, hex_pixval, red_val, green_val, blue_val, newred_val, newgreen_val, newblue_val
Private Declare Function GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFilename As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Standard MS procedure
Public Function GetTempFile(ByVal strDestPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, lpTempFilename As String) As Boolean
   If strDestPath = "" Then
        strDestPath = String(255, vbNullChar)
        If GetTempPath(255, strDestPath) = 0 Then
            GetTempFile = False
            Exit Function
        End If
    End If
    lpTempFilename = String(255, vbNullChar)
    GetTempFile = GetTempFilename(strDestPath, lpPrefixString, wUnique, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
Public Sub Flip(dest As PictureBox, src As PictureBox, H As Boolean)
If H = True Then
    dest.PaintPicture src.Picture, 0, 0, src.ScaleWidth, src.ScaleHeight, src.ScaleWidth, 0, -src.ScaleWidth, src.ScaleHeight, &HCC0020
Else
    dest.PaintPicture src.Picture, 0, 0, src.ScaleWidth, src.ScaleHeight, 0, src.ScaleHeight, src.ScaleWidth, -src.ScaleHeight, &HCC0020
End If
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.AutoSize = True
dest.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub mypastemerge(Spic As PictureBox, Spic2 As PictureBox, Mergepic As PictureBox, Dpic As PictureBox)
Spic2.Width = Selectwidth
Spic2.Height = Selectheight
Spic2.Top = SelectTop
Spic2.Left = SelectLeft
Dpic.Visible = True
Spic2.Picture = LoadPicture()
StretchBlt Spic2.hdc, 0, 0, Selectwidth, Selectheight, Spic.hdc, Dpic.Left, Dpic.Top, Selectwidth, Selectheight, SRCCOPY
Dpic.Refresh
MERGE Dpic, Spic2, Mergepic
StretchBlt Spic.hdc, SelectLeft, SelectTop, Selectwidth, Selectheight, Mergepic.hdc, 0, 0, Selectwidth, Selectheight, SRCCOPY
Spic.Refresh
MergePasting = False
End Sub
Public Sub MERGE(img1 As PictureBox, img2 As PictureBox, dest As PictureBox)
    On Error GoTo fin:
    Dim x As Long
    Dim y As Long
    Dim R1 As Integer
    Dim G1 As Integer
    Dim B1 As Integer
    Dim r2 As Integer
    Dim G2 As Integer
    Dim B2 As Integer
    dest.Cls
    dest.Height = img1.Height
    dest.Width = img1.Width
    For x = 0 To img1.ScaleWidth
        DoEvents
        For y = 0 To img1.ScaleHeight
                GET_COLORS img1.Point(x, y), R1, G1, B1
                GET_COLORS img2.Point(x, y), r2, G2, B2
                dest.PSet (x, y), RGB((R1 + r2) / 2, (G2 + G1) / 2, (B2 + B1) / 2)
        Next y
       frmMain.Pb.Value = x * 100 \ (img1.ScaleWidth)
    Next x
frmMain.Pb.Value = 0
Exit Sub
fin:
End Sub
Private Sub GET_COLORS(Color As Long, ByRef R As Integer, ByRef G As Integer, ByRef b As Integer)
    Dim temp As Long
    temp = (Color And 255)
    R = temp And 255
    temp = Int(Color / 256)
    G = temp And 255
    temp = Int(Color / 65536)
    b = temp And 255
End Sub
Public Sub ImgResize(dest As PictureBox, src As PictureBox, w As Long, H As Long)
src.Height = H
src.Width = w
StretchBlt src.hdc, 0, 0, src.Width / Screen.TwipsPerPixelX, src.Height / Screen.TwipsPerPixelY, dest.hdc, 0, 0, dest.Width, dest.Height, SRCCOPY
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.AutoSize = True
dest.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub Rotate(Form As Long, src As PictureBox, dest As PictureBox, Rfactor As Double)
Dim c1x As Integer, c1y As Integer
Dim c2x As Integer, c2y As Integer
Dim a As Double
Dim p1x As Integer, p1y As Integer
Dim p2x As Integer, p2y As Integer
Dim n As Integer, R As Integer
Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long
dest.AutoSize = False
dest.Height = src.Width
dest.Width = src.Height
dest.Picture = LoadPicture()
Rfactor = 360 - Rfactor
Rfactor = Rfactor * (3.14159265358979 / 180)
c1x = src.ScaleWidth \ 2
c1y = src.ScaleHeight \ 2
c2x = dest.ScaleWidth \ 2
c2y = dest.ScaleHeight \ 2
If c2x < c2y Then n = c2y Else n = c2x
n = (n - 1) * 2
SrchDc = src.hdc
DesthDc = dest.hdc
For p2x = 0 To n / 2
    For p2y = 0 To n / 2
        If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
        R = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
        p1x = R * Cos(a + Rfactor)
        p1y = R * Sin(a + Rfactor)
        c0& = GetPixel(SrchDc, c1x + p1x, c1y + p1y)
        c1& = GetPixel(SrchDc, c1x - p1x, c1y - p1y)
        c2& = GetPixel(SrchDc, c1x + p1y, c1y - p1x)
        c3& = GetPixel(SrchDc, c1x - p1y, c1y + p1x)
        If c0& <> -1 Then SetPixel DesthDc, c2x + p2x, c2y + p2y, c0&
        If c1& <> -1 Then SetPixel DesthDc, c2x - p2x, c2y - p2y, c1&
        If c2& <> -1 Then SetPixel DesthDc, c2x + p2y, c2y - p2x, c2&
        If c3& <> -1 Then SetPixel DesthDc, c2x - p2y, c2y + p2x, c3&
    Next
Next
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.AutoSize = True
dest.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub MyColAdjust(src As PictureBox, dest As PictureBox, buz As Integer, pg As ProgressBar)
pg.Max = (src.ScaleWidth - 1) * 2
PrepareImg src, pg
For i = 0 To src.ScaleWidth - 1
    For j = 0 To src.ScaleHeight - 1
        If buz = 0 Then
            R = Abs(larrCol(0, i, j) + 20)
            G = Abs(larrCol(1, i, j))
            b = Abs(larrCol(2, i, j))
        End If
        If buz = 1 Then
            R = Abs(larrCol(0, i, j))
            G = Abs(larrCol(1, i, j) + 20)
            b = Abs(larrCol(2, i, j))
        End If
        If buz = 2 Then
            R = Abs(larrCol(0, i, j))
            G = Abs(larrCol(1, i, j))
            b = Abs(larrCol(2, i, j) + 20)
        End If
        If buz = 3 Then
            R = Abs(larrCol(0, i, j) - 20)
            G = Abs(larrCol(1, i, j))
            b = Abs(larrCol(2, i, j))
        End If
        If buz = 4 Then
            R = Abs(larrCol(0, i, j))
            G = Abs(larrCol(1, i, j) - 20)
            b = Abs(larrCol(2, i, j))
        End If
        If buz = 5 Then
            R = Abs(larrCol(0, i, j))
            G = Abs(larrCol(1, i, j))
            b = Abs(larrCol(2, i, j) - 20)
        End If
        If larrCol(0, i, j) > 240 And larrCol(1, i, j) > 240 Then R = 255: G = 255: b = 255
        If larrCol(2, i, j) > 240 And larrCol(1, i, j) > 240 Then R = 255: G = 255: b = 255
        If larrCol(0, i, j) > 240 And larrCol(2, i, j) > 240 Then R = 255: G = 255: b = 255
        If R > 240 And b > 240 Then G = 255
        If G > 240 And b > 240 Then R = 255
        If R > 240 And G > 240 Then b = 255
        SetPixel src.hdc, i, j, RGB(R, G, b)
    Next j
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (src.ScaleWidth - 1) + i
    End If
Next i
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub MyDiffuse(src As PictureBox, dest As PictureBox, pg As ProgressBar, alan As Integer)
pg.Max = src.ScaleHeight - 3
For i = 1 To src.ScaleHeight - 3
    For j = 1 To src.ScaleWidth - 3
        pixval = src.Point(j + CInt(Rnd * alan), i + CInt(Rnd * alan))
        red_val = "&h" & Mid$(CStr(Hex(pixval)), 5, 2)
        green_val = "&h" & Mid$(CStr(Hex(pixval)), 3, 2)
        blue_val = "&h" & Mid$(CStr(Hex(pixval)), 1, 2)
        If red_val = "&h" Then red_val = "&h0"
        If green_val = "&h" Then green_val = "&h0"
        If blue_val = "&h" Then blue_val = "&h0"
        newred_val = 0
        newgreen_val = 0
        newblue_val = 0
        If red_val > 240 Then newred_val = 200
        If green_val > 240 Then newgreen_val = 200
        If blue_val > 240 Then newblue_val = 200
        If newred_val > 0 Then red_val = newred_val
        If newred_val > 0 Then green_val = newgreen_val
        If newred_val > 0 Then blue_val = newblue_val
        If freeselection = True Then
            FS = GetPixel(src.hdc, i, j)
            If FS <> 8950944 Then src.PSet (j, i), RGB(red_val, green_val, blue_val)
        Else
            src.PSet (j, i), RGB(red_val, green_val, blue_val)
        End If
    Next
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = i
    End If
Next
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub MyOutline(src As PictureBox, dest As PictureBox, pg As ProgressBar, alan As Integer)
Dim col As Long
pg.Max = src.ScaleWidth - 1
src.Picture = LoadPicture()
For x = 1 To dest.ScaleWidth - 1
    For y = 1 To dest.ScaleHeight - 1
    If freeselection = True Then
        FS = GetPixel(src.hdc, i, j)
        If FS = 8950944 Then GoTo freddy
    Else
     col = Abs(GetPixel(dest.hdc, x, y) - GetPixel(dest.hdc, x, y - 1))
     If col > alan ^ 3 Then col = 0 Else col = vbWhite
     If col = 0 Then SetPixel src.hdc, x, y, col
     col = Abs(GetPixel(dest.hdc, x, y) - GetPixel(dest.hdc, x - 1, y))
     If col > alan ^ 3 Then col = 0 Else col = vbWhite
     If col = 0 Then SetPixel src.hdc, x, y, col
    End If
freddy:
    Next y
    If x > 0 Then
        If pg.Value < pg.Max Then pg.Value = x
    End If
Next x
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Private Sub PrepareImg(src As PictureBox, pg As ProgressBar)
ReDim larrCol(2, src.ScaleWidth, src.ScaleHeight)
For i = 0 To src.ScaleWidth - 1
    For j = 0 To src.ScaleHeight - 1
        tmpCol = GetPixel(src.hdc, i, j)
        R = tmpCol Mod 256
        G = (tmpCol / 256) Mod 256
        b = tmpCol / 256 / 256
        larrCol(0, i, j) = R
        larrCol(1, i, j) = G
        larrCol(2, i, j) = b
    Next j
If i > 0 Then
    If pg.Value < pg.Max Then pg.Value = i
End If
Next i
End Sub
Public Sub MyBrightness(src As PictureBox, dest As PictureBox, mylevel As Integer, pg As ProgressBar)
pg.Max = (src.ScaleWidth - 1) * 2
Dim FS As Long
PrepareImg src, pg
For i = 0 To src.ScaleWidth - 1
    For j = 0 To src.ScaleHeight - 1
        R = Abs(larrCol(0, i, j) + mylevel)
        G = Abs(larrCol(1, i, j) + mylevel)
        b = Abs(larrCol(2, i, j) + mylevel)
        If larrCol(0, i, j) > 240 And larrCol(1, i, j) > 240 Then R = 255: G = 255: b = 255
        If larrCol(2, i, j) > 240 And larrCol(1, i, j) > 240 Then R = 255: G = 255: b = 255
        If larrCol(0, i, j) > 240 And larrCol(2, i, j) > 240 Then R = 255: G = 255: b = 255
        If R > 240 And b > 240 Then G = 255
        If G > 240 And b > 240 Then R = 255
        If R > 240 And G > 240 Then b = 255
        If freeselection = True Then
            FS = GetPixel(src.hdc, i, j)
            If FS <> 8950944 Then SetPixel src.hdc, i, j, RGB(R, G, b)
        Else
            SetPixel src.hdc, i, j, RGB(R, G, b)
        End If
    Next j
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (src.ScaleWidth - 1) + i
    End If
Next i
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub MyGreyscale(src As PictureBox, dest As PictureBox, pg As ProgressBar)
Dim c As Long
pg.Max = (src.ScaleWidth - 1) * 2
PrepareImg src, pg
For i = 0 To src.ScaleWidth - 1
    For j = 0 To src.ScaleHeight - 1
        c = larrCol(0, i, j) * 0.3 + larrCol(1, i, j) * 0.59 + larrCol(2, i, j) * 0.11
        If larrCol(0, i, j) > 240 And larrCol(1, i, j) > 240 Then c = 255
        If larrCol(2, i, j) > 240 And larrCol(1, i, j) > 240 Then c = 255
        If larrCol(0, i, j) > 240 And larrCol(2, i, j) > 240 Then c = 255
        If freeselection = True Then
            FS = GetPixel(src.hdc, i, j)
            If FS <> 8950944 Then SetPixel src.hdc, i, j, RGB(c, c, c)
        Else
            SetPixel src.hdc, i, j, RGB(c, c, c)
        End If
    Next j
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (src.ScaleWidth - 1) + i
    End If
Next i
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub MyEmboss(src As PictureBox, dest As PictureBox, pg As ProgressBar, alan As Integer)
pg.Max = (src.ScaleWidth - 1) * 2
PrepareImg src, pg
For i = 0 To src.ScaleWidth - 1
    For j = 0 To src.ScaleHeight - 1
                R = Abs(larrCol(0, i, j) - larrCol(0, i + 1, j + 1) + alan)
                G = Abs(larrCol(1, i, j) - larrCol(1, i + 1, j + 1) + alan)
                b = Abs(larrCol(2, i, j) - larrCol(2, i + 1, j + 1) + alan)
                If R > 220 Then R = 220
                If G > 220 Then G = 220
                If b > 220 Then b = 220
                If freeselection = True Then
                    FS = GetPixel(src.hdc, i, j)
                    If FS <> 8950944 Then SetPixel src.hdc, i, j, RGB(R, G, b)
                Else
                    SetPixel src.hdc, i, j, RGB(R, G, b)
                End If
            Next j
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (src.ScaleWidth - 1) + i
    End If
Next i
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub MySharpen(src As PictureBox, dest As PictureBox, pg As ProgressBar, alan As Integer)
pg.Max = (src.ScaleWidth - 1) * 2
PrepareImg src, pg
For i = 1 To src.ScaleWidth - 1
    For j = 1 To src.ScaleHeight - 1
        R = larrCol(0, i, j) + 0.5 * alan * (larrCol(0, i, j) - larrCol(0, i - 1, j - 1))
        G = larrCol(1, i, j) + 0.5 * alan * (larrCol(1, i, j) - larrCol(1, i - 1, j - 1))
        b = larrCol(2, i, j) + 0.5 * alan * (larrCol(2, i, j) - larrCol(2, i - 1, j - 1))
        If R > 200 Then R = 240
        If R < 0 Then R = 0
        If G > 200 Then G = 240
        If G < 0 Then G = 0
        If b > 200 Then b = 240
        If b < 0 Then b = 0
        If freeselection = True Then
            FS = GetPixel(src.hdc, i, j)
            If FS <> 8950944 Then SetPixel src.hdc, i, j, RGB(R, G, b)
        Else
            SetPixel src.hdc, i, j, RGB(R, G, b)
        End If
    Next j
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (src.ScaleWidth - 1) + i
    End If
Next i
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub MyBlur(src As PictureBox, dest As PictureBox, pg As ProgressBar, alan As Integer)
pg.Max = src.ScaleWidth - 1 + src.ScaleWidth - alan
PrepareImg src, pg
Dim jeanette As Integer
jeanette = alan
For i = 1 To src.ScaleWidth - alan
    For j = 1 To src.ScaleHeight - alan
        If alan > i Then alan = i
        If alan > j Then alan = j
        R = Abs(larrCol(0, i - alan, j - alan) + larrCol(0, i, j - alan) + larrCol(0, i + alan, j - alan) + larrCol(0, i - alan, j) + larrCol(0, i, j) + larrCol(0, i + alan, j) + larrCol(0, i - alan, j + alan) + larrCol(0, i, j + alan) + larrCol(0, i + alan, j + alan))
        G = Abs(larrCol(1, i - alan, j - alan) + larrCol(1, i, j - alan) + larrCol(1, i + alan, j - alan) + larrCol(1, i - alan, j) + larrCol(1, i, j) + larrCol(1, i + alan, j) + larrCol(1, i - alan, j + alan) + larrCol(1, i, j + alan) + larrCol(1, i + alan, j + alan))
        b = Abs(larrCol(2, i - alan, j - alan) + larrCol(2, i, j - alan) + larrCol(2, i + alan, j - alan) + larrCol(2, i - alan, j) + larrCol(2, i, j) + larrCol(2, i + alan, j) + larrCol(2, i - alan, j + alan) + larrCol(2, i, j + alan) + larrCol(2, i + alan, j + alan))
        If freeselection = True Then
            FS = GetPixel(src.hdc, i, j)
            If FS <> 8950944 Then SetPixel src.hdc, i, j, RGB(R / 10, G / 10, b / 10)
        Else
            SetPixel src.hdc, i, j, RGB(R / 10, G / 10, b / 10)
        End If
        alan = jeanette
    Next j
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (src.ScaleWidth - 1) + i
    End If
Next i
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture src.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub MyInvert(src As PictureBox, dest As PictureBox)
dest.PaintPicture src.Picture, 0, 0, src.ScaleWidth, src.ScaleHeight, 0, 0, src.ScaleWidth, src.ScaleHeight, vbNotSrcCopy
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub DrawLight(pic As PictureBox, pic2 As PictureBox, Target As Long, x As Single, y As Single, RedB As Long, GreenB As Long, BlueB As Long, Radius As Long, NumberOfSteps As Long)
Dim cx As Long
Dim cy As Long
Dim TempColor As Long
Dim TempRadius As Integer
Dim Done() As Boolean
ReDim Done(-Radius To Radius, -Radius To Radius)
For i = 1 To NumberOfSteps
    If CWcancel = True Then
        Exit For
        Exit Sub
    End If
    TempRadius = Radius / NumberOfSteps * i
    For cx = -TempRadius To TempRadius
        If CWcancel = True Then
            Exit For
            Exit Sub
        End If
        For cy = -TempRadius To TempRadius
            If CWcancel = True Then
                Exit For
                Exit Sub
            End If
            If Not Done(cx, cy) Then
                If (cx * cx) + (cy * cy) <= TempRadius * TempRadius Then
                    If CWcancel = True Then
                        Exit For
                        Exit Sub
                    End If
                    TempColor = GetPixel(Target, cx + x, cy + y)
                    GetRGBs TempColor
                    Red = Red + RedB * (NumberOfSteps - i)
                    If Red > 255 Then Red = 255
                    If Red < 0 Then Red = 0
                    Green = Green + GreenB * (NumberOfSteps - i)
                    If Green > 255 Then Green = 255
                    If Green < 0 Then Green = 0
                    Blue = Blue + BlueB * (NumberOfSteps - i)
                    If Blue > 255 Then Blue = 255
                    If Blue < 0 Then Blue = 0
                    SetPixel Target, cx + x, cy + y, RGB(Red, Green, Blue)
                    Done(cx, cy) = True
                    If CWcancel = True Then
                        Exit For
                        Exit Sub
                    End If
                End If
            End If
            If CWcancel = True Then
                Exit For
                Exit Sub
            End If
        Next cy
        If CWcancel = True Then
            Exit For
            Exit Sub
        End If
    Next cx
    If CWcancel = True Then
        Exit For
        Exit Sub
    End If
Next i
If GetTempFile("", "BI", 0, sfilename) Then SavePicture frmMain.ActiveForm.picsource.Image, sfilename
pic.Picture = LoadPicture(sfilename)
pic2.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub GetRGBs(RGBVal As Long)
    Red = RGBVal And 255
    Green = (RGBVal And 65280) \ 256
    Blue = (RGBVal And 16711680) \ 65535
End Sub
Public Sub MyPixelate(src As PictureBox, dest As PictureBox, pg As ProgressBar, alan As Integer)
Dim simon As Integer
simon = alan
pg.Max = src.ScaleWidth + alan
For i = 0 To src.ScaleWidth Step alan
    For j = 0 To src.ScaleHeight Step alan
        If j + 1 > src.ScaleHeight Then j = src.ScaleHeight - 1
        If i + 1 > src.ScaleWidth Then i = src.ScaleWidth - 1
        tempval = src.Point(i + 1, j + 1)
        dest.Line (i, j)-(i + alan, j + alan), tempval, BF
        alan = simon
    Next j
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = i
    End If
Next i
pg.Value = 0
End Sub
Public Sub Filling(pic As PictureBox, col As Long, ByVal FStyle As Long, x, y)
    Dim a As Long
    pic.FillStyle = FStyle
    pic.FillColor = frmMain.TheColor.BackColor
    a = ExtFloodFill(pic.hdc, x, y, col, 1)
    pic.FillStyle = 1
End Sub
Public Sub Filling2(pic As PictureBox, col As Long, ByVal FStyle As Long, x, y)
    Dim a As Long
    pic.FillStyle = FStyle
    pic.FillColor = vbBlack
    a = ExtFloodFill(pic.hdc, x, y, col, 1)
    pic.FillStyle = 1
End Sub
Public Sub MyEdgeinnerline(src As PictureBox, dest As PictureBox, mylevel As Integer, mywidth As Integer, pg As ProgressBar, mycol As Long)
For j = mywidth To dest.ScaleWidth - mywidth
    SetPixel dest.hdc, j, mywidth, mycol
    SetPixel dest.hdc, j, dest.ScaleHeight - mywidth, mycol
Next j
For j = mywidth To dest.ScaleHeight - mywidth
    SetPixel dest.hdc, mywidth, j, mycol
    SetPixel dest.hdc, dest.ScaleWidth - mywidth, j, mycol
Next j
End Sub
Public Sub MyEdgeouterline(src As PictureBox, dest As PictureBox, mylevel As Integer, mywidth As Integer, pg As ProgressBar, mycol As Long)
For j = 0 To dest.ScaleWidth - 1
    SetPixel dest.hdc, j, 0, vbBlack
    SetPixel dest.hdc, j, dest.ScaleHeight - 1, vbBlack
Next j
For j = 0 To dest.ScaleHeight - 1
    SetPixel dest.hdc, 0, j, vbBlack
    SetPixel dest.hdc, dest.ScaleWidth - 1, j, vbBlack
Next j
End Sub
Public Sub Butt1(src As PictureBox, dest As PictureBox, mywidth As Integer, mylevel1 As Integer, mylevel2 As Integer, outer As Boolean, inner As Boolean, pg As ProgressBar)
pg.Max = (src.ScaleWidth - 1) * 2
Dim mw As Integer
Dim mz As Integer
Dim tempsat As Integer
Dim temphue As Integer
Dim templum As Integer
Dim HSLV As HSLCol
mw = mylevel1
mz = mylevel2
mylevel1 = mylevel1 / 3
mylevel2 = mylevel2 / 3
curborderlevel2 = mw
curborderlevel3 = mz
mw = 0
mz = 0
For i = 1 To dest.ScaleWidth - 1
    If i < mywidth Then mw = mw + 1
    If mw > mywidth Then mw = mywidth
    If i > dest.ScaleWidth - mywidth - 1 Then mw = mw - 1
    For j = 1 To mw
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat + mylevel1
        HSLV.Lum = templum
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
    Next j
    For j = dest.ScaleHeight - mw To dest.ScaleHeight - 1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat
        HSLV.Lum = templum - mylevel2
        tempval = HSLtoRGB(HSLV)
        Red = RGBRed(tempval)
        Green = RGBGreen(tempval)
        Blue = RGBBlue(tempval)
        If Red > 160 Then Red = 160
        If Green > 160 Then Green = 160
        If Blue > 160 Then Blue = 160
        tempval = RGB(Red, Green, Blue)
        SetPixel dest.hdc, i, j, tempval
    Next j
    If i < mywidth Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat + mylevel1
        HSLV.Lum = templum
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        Next j
    End If
    If i > dest.ScaleWidth - mywidth - 1 Then
        If mw < 0 Then mw = 0
        For j = mw + 1 To dest.ScaleHeight - mw - 1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat
        HSLV.Lum = templum - mylevel2
        tempval = HSLtoRGB(HSLV)
        Red = RGBRed(tempval)
        Green = RGBGreen(tempval)
        Blue = RGBBlue(tempval)
        If Red > 160 Then Red = 160
        If Green > 160 Then Green = 160
        If Blue > 160 Then Blue = 160
        tempval = RGB(Red, Green, Blue)
        SetPixel dest.hdc, i, j, tempval
        Next j
    End If
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (dest.ScaleWidth - 1) + i
    End If
Next i
If outline Then MyEdgeouterline src, dest, mylevel1, mywidth, pg, vbBlack
If inline Then MyEdgeinnerline src, dest, mylevel1, mywidth, pg, vbBlack
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub Butt2(src As PictureBox, dest As PictureBox, mywidth As Integer, mylevel1 As Integer, mylevel2 As Integer, outer As Boolean, inner As Boolean, pg As ProgressBar)
pg.Max = (src.ScaleWidth - 1) * 2
Dim mw As Integer
Dim mt As Long
Dim mb As Long
Dim ml As Long
Dim mr As Long
Dim tempsat As Integer
Dim temphue As Integer
Dim templum As Integer
Dim HSLV As HSLCol
mw = mylevel1
mt = mylevel2
mylevel1 = mylevel1 * 2
curborderlevel2 = mw
curborderlevel3 = mt
mw = 0
mt = 0
ml = mylevel1
For i = 1 To dest.ScaleWidth - 1
    If i < mywidth Then mw = mw + 1
    If mw > mywidth Then mw = mywidth
    If i > dest.ScaleWidth - mywidth - 1 Then mw = mw - 1
    mt = mylevel1
    For j = 1 To mw
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat + mt
        HSLV.Lum = templum
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        mt = mt - (mylevel1) / mywidth
    Next j
    mb = mylevel2
    For j = dest.ScaleHeight - 1 To dest.ScaleHeight - mw Step -1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat
        HSLV.Lum = templum - mb
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        mb = mb - (mylevel2) / mywidth
    Next j
    If i < mywidth Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat + ml
        HSLV.Lum = templum
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        Next j
    End If
    ml = ml - (mylevel1) / mywidth
    If i > dest.ScaleWidth - mywidth - 1 Then
            mr = mr + (mylevel2) / mywidth
        If mw < 0 Then mw = 0
        For j = mw + 1 To dest.ScaleHeight - mw - 1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat
        HSLV.Lum = templum - mr
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        Next j
    End If
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (dest.ScaleWidth - 1) + i
    End If
Next i
If outline Then MyEdgeouterline src, dest, mylevel1, mywidth, pg, vbBlack
If inline Then MyEdgeinnerline src, dest, mylevel1, mywidth, pg, vbBlack
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub FlatBorder(src As PictureBox, dest As PictureBox, mywidth As Integer, mycol As Long, outer As Boolean, inner As Boolean, pg As ProgressBar)
pg.Max = (src.ScaleWidth - 1) * 2
Dim mw As Integer
Dim mz As Integer
Dim tempsat As Integer
Dim temphue As Integer
Dim templum As Integer
Dim HSLV As HSLCol
mw = 0
mz = 0
For i = 1 To dest.ScaleWidth - 1
    If i < mywidth Then mw = mw + 1
    If mw > mywidth Then mw = mywidth
    If i > dest.ScaleWidth - mywidth - 1 Then mw = mw - 1
    For j = 1 To mw
        SetPixel dest.hdc, i, j, mycol
    Next j
    For j = dest.ScaleHeight - mw To dest.ScaleHeight - 1
        SetPixel dest.hdc, i, j, mycol
    Next j
    If i < mywidth Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
        SetPixel dest.hdc, i, j, mycol
        Next j
    End If
    If i > dest.ScaleWidth - mywidth - 1 Then
        If mw < 0 Then mw = 0
        For j = mw + 1 To dest.ScaleHeight - mw - 1
        SetPixel dest.hdc, i, j, mycol
        Next j
    End If
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (dest.ScaleWidth - 1) + i
    End If
Next i
If outline Then MyEdgeouterline src, dest, 0, 0, pg, vbBlack
If inline Then MyEdgeinnerline src, dest, 0, mywidth, pg, vbBlack
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub
Public Sub ButtBW(src As PictureBox, dest As PictureBox, mywidth As Integer, mylevel1 As Integer, mylevel2 As Integer, outer As Boolean, inner As Boolean, raised As Boolean, pg As ProgressBar)
pg.Max = (src.ScaleWidth - 1) * 2
Dim mw As Integer
Dim mz As Integer
Dim mq As Integer
mq = mywidth
mywidth = mywidth / 8
mw = 0
mz = 0
For i = 1 To dest.ScaleWidth - 1
    If i < mywidth Then mw = mw + 1
    If mw > mywidth Then mw = mywidth
    If i > dest.ScaleWidth - mywidth - 1 Then mw = mw - 1
    For j = 1 To mw
        If j < mywidth / 2 + 1 Then
            If raised = True Then
                SetPixel dest.hdc, i, j, RGB(255, 255, 255)
            Else
                SetPixel dest.hdc, i, j, RGB(0, 0, 0)
            End If
        End If
        If j > mywidth / 2 And j < mywidth Then
            If raised = True Then
                SetPixel dest.hdc, i, j, RGB(192, 192, 192)
            Else
                SetPixel dest.hdc, i, j, RGB(128, 128, 128)
            End If
        End If
    Next j
    For j = dest.ScaleHeight - mw To dest.ScaleHeight - 1
        If j > dest.ScaleHeight - mywidth And j < dest.ScaleHeight - mywidth / 2 Then
            If raised = True Then
                SetPixel dest.hdc, i, j, RGB(128, 128, 128)
            Else
                SetPixel dest.hdc, i, j, RGB(192, 192, 192)
            End If
        End If
        If j > dest.ScaleHeight - mywidth / 2 - 1 And j < dest.ScaleHeight Then
            If raised = True Then
                SetPixel dest.hdc, i, j, RGB(0, 0, 0)
            Else
                SetPixel dest.hdc, i, j, RGB(255, 255, 255)
            End If
        End If
    Next j
    If i < mywidth / 2 + 1 Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
            If raised = True Then
                SetPixel dest.hdc, i, j, RGB(255, 255, 255)
            Else
                SetPixel dest.hdc, i, j, RGB(0, 0, 0)
            End If
        Next j
    End If
    If i > mywidth / 2 And i < mywidth Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
            If raised = True Then
                SetPixel dest.hdc, i, j, RGB(192, 192, 192)
            Else
                SetPixel dest.hdc, i, j, RGB(128, 128, 128)
            End If
        Next j
    End If
    If i < dest.ScaleWidth - mywidth / 2 And i > dest.ScaleWidth - mywidth - 1 Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
            If raised = True Then
                SetPixel dest.hdc, i, j, RGB(128, 128, 128)
            Else
                SetPixel dest.hdc, i, j, RGB(192, 192, 192)
            End If
        Next j
    End If
    If i > dest.ScaleWidth - mywidth / 2 - 1 Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
            If raised = True Then
                SetPixel dest.hdc, i, j, RGB(0, 0, 0)
            Else
                SetPixel dest.hdc, i, j, RGB(255, 255, 255)
            End If
        Next j
    End If
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (dest.ScaleWidth - 1) + i
    End If
Next i
If outline Then MyEdgeouterline src, dest, mylevel1, mywidth, pg, vbBlack
If inline Then MyEdgeinnerline src, dest, mylevel1, mywidth, pg, vbBlack
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
borderwidth = mq
End Sub
Public Sub Frame3D(src As PictureBox, dest As PictureBox, mywidth As Integer, myFwidth As Integer, col As Long, pg As ProgressBar)
pg.Max = (src.ScaleWidth - 1) * 2
Dim mw As Integer
Dim mz As Integer
Dim mq As Integer
mq = mywidth
mywidth = mywidth / 8
mw = 0
mz = 0
For i = 1 To dest.ScaleWidth - 1
    If i < mywidth Then mw = mw + 1
    If mw > mywidth Then mw = mywidth
    If i > dest.ScaleWidth - mywidth - 1 Then mw = mw - 1
    For j = 1 To mw
        If j < mywidth / 2 Then
                SetPixel dest.hdc, i, j, RGB(255, 255, 255)
        End If
        If j > mywidth / 2 - 1 And j < mywidth Then
                SetPixel dest.hdc, i, j, RGB(204, 204, 204)
        End If
    Next j
    For j = dest.ScaleHeight - mw To dest.ScaleHeight - 1
        If j > dest.ScaleHeight - mywidth And j < dest.ScaleHeight - mywidth / 2 Then
                SetPixel dest.hdc, i, j, RGB(128, 128, 128)
        End If
        If j > dest.ScaleHeight - mywidth / 2 - 1 And j < dest.ScaleHeight Then
                SetPixel dest.hdc, i, j, RGB(64, 64, 64)
        End If
    Next j
    If i < mywidth / 2 Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
                SetPixel dest.hdc, i, j, RGB(255, 255, 255)
        Next j
    End If
    If i > mywidth / 2 - 1 And i < mywidth Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
                SetPixel dest.hdc, i, j, RGB(204, 204, 204)
        Next j
    End If
    If i < dest.ScaleWidth - mywidth / 2 And i > dest.ScaleWidth - mywidth - 1 Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
                SetPixel dest.hdc, i, j, RGB(128, 128, 128)
        Next j
    End If
    If i > dest.ScaleWidth - mywidth / 2 - 1 Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
                SetPixel dest.hdc, i, j, RGB(64, 64, 64)
        Next j
    End If
    If i > mywidth - 1 And i < myFwidth Then
        For j = mywidth To dest.ScaleHeight - mywidth
            SetPixel dest.hdc, i, j, col
        Next j
    ElseIf i > dest.ScaleWidth - myFwidth - 1 And i < dest.ScaleWidth - mywidth Then
        For j = mywidth To dest.ScaleHeight - mywidth
            SetPixel dest.hdc, i, j, col
        Next j
    ElseIf i > myFwidth - 1 And i < dest.ScaleWidth - myFwidth + 1 Then
        For j = mywidth To myFwidth
            SetPixel dest.hdc, i, j, col
        Next j
        For j = dest.ScaleHeight - myFwidth To dest.ScaleHeight - mywidth
            SetPixel dest.hdc, i, j, col
        Next j
    End If
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (dest.ScaleWidth - 1) + i
    End If
Next i
innerFrame3D src, dest, mywidth, myFwidth, pg
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
borderwidth = mq
End Sub
Public Sub innerFrame3D(src As PictureBox, dest As PictureBox, mywidth As Integer, myFwidth As Integer, pg As ProgressBar)
Dim mw As Integer
Dim mz As Integer
mw = 0
mz = 0
For i = myFwidth To dest.ScaleWidth - myFwidth - 1
    If i < mywidth + myFwidth Then mw = mw + 1
    If mw > mywidth + myFwidth Then mw = mywidth
    If i > dest.ScaleWidth - mywidth - myFwidth - 1 Then mw = mw - 1
    For j = myFwidth To mw + myFwidth
        If j < mywidth / 2 + 1 + myFwidth Then
                SetPixel dest.hdc, i, j, RGB(128, 128, 128)
        End If
        If j > mywidth / 2 + myFwidth And j < mywidth + myFwidth + 1 Then
                SetPixel dest.hdc, i, j, RGB(64, 64, 64)
        End If
    Next j
    For j = dest.ScaleHeight - myFwidth - mw To dest.ScaleHeight - myFwidth - 1
        If j > dest.ScaleHeight - mywidth - myFwidth And j < dest.ScaleHeight - myFwidth - mywidth / 2 - 1 Then
                SetPixel dest.hdc, i, j, RGB(255, 255, 255)
        End If
        If j > dest.ScaleHeight - myFwidth - mywidth / 2 - 2 And j < dest.ScaleHeight - myFwidth Then
                SetPixel dest.hdc, i, j, RGB(204, 204, 204)
        End If
    Next j
    If i < mywidth / 2 + myFwidth Then
        For j = mw + 1 + myFwidth To dest.ScaleHeight - mw - 1 - myFwidth
                SetPixel dest.hdc, i, j, RGB(128, 128, 128)
        Next j
    End If
    If i > mywidth / 2 + myFwidth - 1 And i < mywidth + myFwidth Then
        For j = mw + 1 + myFwidth To dest.ScaleHeight - mw - 1 - myFwidth
                SetPixel dest.hdc, i, j, RGB(64, 64, 64)
        Next j
    End If
    If i < dest.ScaleWidth - mywidth / 2 - myFwidth - 1 And i > dest.ScaleWidth - mywidth - myFwidth - 1 Then
        For j = mw + myFwidth + 1 To dest.ScaleHeight - myFwidth - mw - 1
                SetPixel dest.hdc, i, j, RGB(255, 255, 255)
        Next j
    End If
    If i > dest.ScaleWidth - mywidth / 2 - myFwidth - 2 And i < dest.ScaleWidth - myFwidth Then
        For j = mw + myFwidth + 1 To dest.ScaleHeight - mw - 1 - myFwidth
                SetPixel dest.hdc, i, j, RGB(204, 204, 204)
        Next j
    End If
Next i
End Sub
Public Sub Frame3D2(src As PictureBox, dest As PictureBox, mywidth As Integer, mylevel1 As Integer, mylevel2 As Integer, outer As Boolean, inner As Boolean, myFwidth As Integer, pg As ProgressBar)
pg.Max = (src.ScaleWidth - 1) * 2
Dim mw As Integer
Dim mt As Long
Dim mb As Long
Dim ml As Long
Dim mr As Long
Dim tempsat As Integer
Dim temphue As Integer
Dim templum As Integer
Dim HSLV As HSLCol
mw = mylevel1
mt = mylevel2
mylevel1 = mylevel1 * 2
curborderlevel2 = mw
curborderlevel3 = mt
mw = 0
mt = 0
ml = mylevel1
For i = 1 To dest.ScaleWidth - 1
    If i < mywidth Then mw = mw + 1
    If mw > mywidth Then mw = mywidth
    If i > dest.ScaleWidth - mywidth - 1 Then mw = mw - 1
    mt = mylevel1
    For j = 1 To mw
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat + mt
        HSLV.Lum = templum
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        mt = mt - (mylevel1) / mywidth
    Next j
    mb = mylevel2
    For j = dest.ScaleHeight - 1 To dest.ScaleHeight - mw Step -1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat
        HSLV.Lum = templum - mb
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        mb = mb - (mylevel2) / mywidth
    Next j
    If i < mywidth Then
        For j = mw + 1 To dest.ScaleHeight - mw - 1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat + ml
        HSLV.Lum = templum
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        Next j
    End If
    ml = ml - (mylevel1) / mywidth
    If i > dest.ScaleWidth - mywidth - 1 Then
            mr = mr + (mylevel2) / mywidth
        If mw < 0 Then mw = 0
        For j = mw + 1 To dest.ScaleHeight - mw - 1
        tempval = GetPixel(dest.hdc, i, j)
        If tempval < 0 Then Exit For
        tempsat = RGBtoHSL(tempval).Sat
        temphue = RGBtoHSL(tempval).Hue
        templum = RGBtoHSL(tempval).Lum
        HSLV.Hue = temphue
        HSLV.Sat = tempsat
        HSLV.Lum = templum - mr
        tempval = HSLtoRGB(HSLV)
        SetPixel dest.hdc, i, j, tempval
        Next j
    End If
    If i > 0 Then
        If pg.Value < pg.Max Then pg.Value = (dest.ScaleWidth - 1) + i
    End If
Next i
If outline Then MyEdgeouterline src, dest, mylevel1, mywidth, pg, vbBlack
If inline Then MyEdgeinnerline src, dest, mylevel1, mywidth, pg, vbBlack
pg.Value = 0
If GetTempFile("", "BI", 0, sfilename) Then SavePicture dest.Image, sfilename
dest.Picture = LoadPicture(sfilename)
src.Picture = LoadPicture(sfilename)
Kill sfilename
End Sub

