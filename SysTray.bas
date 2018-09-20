Attribute VB_Name = "SysTray"
Option Explicit

Public Declare Function Shell_NotifyIcon _
   Lib "shell32.dll" _
   Alias "Shell_NotifyIconA" _
   (ByVal dwMessage As Long, _
    lpData As NotifyIconData) As Long
    
'Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
      
Public Type NotifyIconData
    cbSize As Long
    hwnd   As Long
    uId    As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Global sysTrayIcon As NotifyIconData

Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200
'Global Const WM_LBUTTONDBLCLK = &H203   'Double-click
'Global Const WM_LBUTTONDOWN = &H201     'Button down
Global Const WM_LBUTTONUP = &H202       'Button up
'Global Const WM_RBUTTONDBLCLK = &H206   'Double-click
'Global Const WM_RBUTTONDOWN = &H204     'Button down
Global Const WM_RBUTTONUP = &H205       'Button up
Public Const ProductName As String = "FastWrite"

Public Sub InitTrayIcon()
  'put the tray icon
  sysTrayIcon.cbSize = Len(sysTrayIcon)
  sysTrayIcon.hwnd = frmMain.hwnd
  sysTrayIcon.uId = 1&
  sysTrayIcon.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
  sysTrayIcon.uCallbackMessage = WM_MOUSEMOVE
  'sysTrayIcon.hIcon = frmMain.ImageContainer.ListImages("trayGray").Picture
  sysTrayIcon.szTip = ProductName & " " & frmMain.GetVer & " (" & frmMain.lblMode & ")" & vbNullChar
  Shell_NotifyIcon NIM_ADD, sysTrayIcon
End Sub

Public Sub UnloadTrayIcon()
  Shell_NotifyIcon NIM_DELETE, sysTrayIcon
End Sub

Public Sub UpdateTrayIcon()
  sysTrayIcon.szTip = ProductName & " " & frmMain.GetVer & " (" & frmMain.lblMode & ")" & vbNullChar
  Shell_NotifyIcon NIM_MODIFY, sysTrayIcon
End Sub
