Attribute VB_Name = "basRTB"

Option Explicit

' This is used by several functions Some may be pointless but I did not feel like weeding out the ones
' used in this code. You can if you want.
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageByRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETFIRSTVISIBLELINE = &HCE

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Public Const EM_LINEFROMCHAR = &HC9
'Public Const EM_LINEINDEX = &HBB
'Public Const EM_LINELENGTH = &HC1
'Public Const EM_GETLINE = &HC4
'Public Const EM_GETLINECOUNT = &HBA
'Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_GETRECT = &HB2
Public Const WM_GETFONT = &H31
Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1

Public Type TEXTMETRIC
  tmHeight As Long
  tmAscent As Long
  tmDescent As Long
  tmInternalLeading As Long
  tmExternalLeading As Long
  tmAveCharWidth As Long
  tmMaxCharWidth As Long
  tmWeight As Long
  tmOverhang As Long
  tmDigitizedAspectX As Long
  tmDigitizedAspectY As Long
  tmFirstChar As Byte
  tmLastChar As Byte
  tmDefaultChar As Byte
  tmBreakChar As Byte
  tmItalic As Byte
  tmUnderlined As Byte
  tmStruckOut As Byte
  tmPitchAndFamily As Byte
  tmCharSet As Byte
End Type

Public Const SFF_SELECTION = &H8000&
Public Const WM_USER = &H400

Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_POSFROMCHAR = &HD6&
Public Const EM_CHARFROMPOS = &HD7&
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const EM_STREAMIN = (WM_USER + 73)
Public Const EM_HIDESELECTION = WM_USER + 63

Public Const PS_SOLID = 0
Public Const DT_CALCRECT = &H400
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20

Public Const GWL_WNDPROC = (-4)
Private Const WM_VSCROLL = &H115
Public lPrevWndProc As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_VSCROLL Then
        ' This takes the scroll method and scrolls the gutter correctly
        frmMain.DrawLines frmMain.picLines
    End If
    WindowProc = CallWindowProc(lPrevWndProc, hwnd, Msg, wParam, ByVal lParam)
End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Public Function GetFirstVisibleLine(ByVal hwnd As Long) As Long
    GetFirstVisibleLine = SendMessage(hwnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
End Function

Public Function LastVisibleLine(ByVal hwnd As Long) As Long
    LastVisibleLine = GetVisibleLines(hwnd) + GetFirstVisibleLine(hwnd) - 1
End Function

Public Function LineCount(ByVal hwnd As Long) As Long
    LineCount = SendMessageByRef(hwnd, EM_GETLINECOUNT, 0&, 0&)
End Function

Public Function LineForCharacterIndex(lIndex As Long, ByVal hwnd As Long) As Long
   LineForCharacterIndex = SendMessageByLong(hwnd, EM_LINEFROMCHAR, lIndex, 0)
End Function

Public Function GetVisibleLines(ByVal hwnd As Long) As Long
    Dim rc As RECT
    Dim hdc As Long
    Dim lFont As Long
    Dim OldFont As Long
    Dim di As Long
    Dim tm As TEXTMETRIC
    Dim lc As Long
    lc = SendMessage(hwnd, EM_GETRECT, 0, rc)
    lFont = SendMessage(hwnd, WM_GETFONT, 0, 0)
    hdc = GetDC(hwnd)
    If lFont <> 0 Then OldFont = SelectObject(hdc, lFont)
    di = GetTextMetrics(hdc, tm)
    If lFont <> 0 Then lFont = SelectObject(hdc, OldFont)
    GetVisibleLines = (rc.Bottom - rc.Top) / tm.tmHeight
    di = ReleaseDC(hwnd, hdc)
End Function
