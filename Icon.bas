Attribute VB_Name = "Module2"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetForegroundWindow Lib "user32 " (ByVal hWnd As Long) As Long

Public Const WM_USER = &H400
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSEHWHEEL = &H20E
Public Const TRAY_CALLBACK = (WM_USER + 1001&)
Public Const GWL_WNDPROC = (-4)
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_MESSAGE = &H1
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private OldWindowProc As Long
Private TheForm As Form
Private TheMenu As Menu
Private TheData As NOTIFYICONDATA

Public Sub AddToTray(frm As Form, mnu As Menu)
    Set TheForm = frm
    Set TheMenu = mnu
    
    OldWindowProc = SetWindowLong(frm.hWnd, GWL_WNDPROC, AddressOf NewWindowProc)
    
    With TheData
        .cbSize = Len(TheData)
        .hWnd = frm.hWnd
        .uID = 0
        .uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
        .uCallbackMessage = TRAY_CALLBACK
        .hIcon = frm.Icon.Handle
        .szTip = frm.Caption & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, TheData
End Sub

Public Function NewWindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = TRAY_CALLBACK Then
        If lParam = WM_LBUTTONDOWN Then
            If TheForm.WindowState = vbMinimized Then
                TheForm.WindowState = vbNormal
                TheForm.Show
            Else
                TheForm.WindowState = vbMinimized
            End If
            Exit Function
        End If
        If lParam = WM_RBUTTONDOWN Then
            SetForegroundWindow TheForm.hWnd
            TheForm.PopupMenu TheMenu
            Exit Function
        End If
    End If
    
    NewWindowProc = CallWindowProc(OldWindowProc, hWnd, Msg, wParam, lParam)
End Function

Public Sub SetTrayTip(tip As String)
    With TheData
        .szTip = Left(tip, 30) & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

Public Sub SetTrayIcon(pic As Picture)
    If pic.Type <> vbPicTypeIcon Then Exit Sub

    With TheData
        .hIcon = pic.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

Public Sub RemoveFromTray()
    With TheData
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, TheData
    SetWindowLong TheForm.hWnd, GWL_WNDPROC, OldWindowProc
End Sub

