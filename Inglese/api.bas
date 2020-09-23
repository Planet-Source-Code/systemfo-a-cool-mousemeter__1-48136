Attribute VB_Name = "api"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Const Twip_m As Single = 56692.854479
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE

Public Function StayOnTop(frm As Form)
Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_TOPMOST, frm.ScaleLeft, frm.ScaleTop, frm.ScaleWidth, frm.ScaleHeight, FLAGS)
End Function

Public Function NotOnTop(frm As Form)
Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, frm.Left, frm.Top, frm.ScaleWidth, frm.ScaleHeight, FLAGS)
End Function
