Attribute VB_Name = "modClient"
Option Explicit


Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const SM_CXSCREEN = 0        ' Width of screen
Public Const SM_CYSCREEN = 1        ' Height of screen
Public Const SM_CXFULLSCREEN = 16   ' Width of window client area
Public Const SM_CYFULLSCREEN = 17   ' Height of window client area

Private Declare Function GetSystemMetrics& Lib "User32" (ByVal nIndex As Long)

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "User32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long


'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uId As Long
 uFlags As Long
 uCallBackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

Public Type Size
    Width As Long
    Height As Long
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Declare Function SetForegroundWindow Lib "User32" _
(ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public nid          As NOTIFYICONDATA

Public objCrypto    As New clsEncrypt

Public Function GetLoggedInUserName() As String
    On Error Resume Next
    Dim s As String * 256
    GetUserName s, 256
    GetLoggedInUserName = Left(s, InStr(1, s, Chr(0)) - 1)
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function


Public Function GetStartBarSize() As Size
    Dim X As Long
    Dim Y As Long
    Dim fX As Long
    Dim fY As Long
    Dim s As Size

    X = GetSystemMetrics(SM_CXSCREEN)
    Y = GetSystemMetrics(SM_CYSCREEN)
    fX = GetSystemMetrics(SM_CXFULLSCREEN)
    fY = GetSystemMetrics(SM_CYFULLSCREEN)
    
    If fX <> X Then
        X = Abs(fX - X)
    End If
    
    If fY <> Y Then
        Y = Abs(fY - Y)
    End If
    
    s.Width = X
    s.Height = Y
    
    GetStartBarSize = s
End Function

