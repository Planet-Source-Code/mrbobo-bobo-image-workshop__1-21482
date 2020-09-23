Attribute VB_Name = "ModClipHook"
'Standard MS procedure to hook the clipboard
Public DontAdd As Boolean
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetClipboardViewer Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const WM_DRAWCLIPBOARD = &H308
Dim PrevProc As Long
Public Sub HookForm(f As Form)
    PrevProc = SetWindowLong(f.Hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHookForm(f As Form)
    SetWindowLong f.Hwnd, GWL_WNDPROC, PrevProc
End Sub
Public Function WindowProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
WindowProc = CallWindowProc(PrevProc, Hwnd, uMsg, wParam, lParam)
On Error GoTo woops
    If uMsg = WM_DRAWCLIPBOARD Then
        If IsClipboardFormatAvailable(vbCFBitmap) <> 0 Then
            If DontAdd = False Then
                 frmMain.LoadClipboard
            End If
            If frmMain.ChClipLock.Value = 1 Then frmMain.ShClip.Visible = False
        End If
    End If
woops:
End Function

