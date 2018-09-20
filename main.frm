VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FastWrite"
   ClientHeight    =   3735
   ClientLeft      =   12570
   ClientTop       =   8145
   ClientWidth     =   2550
   ClipControls    =   0   'False
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer langTimer 
      Interval        =   100
      Left            =   2760
      Top             =   2640
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer unregTmr 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   2040
      Top             =   2640
   End
   Begin VB.Timer tmrClrInfo 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   2040
      Top             =   2160
   End
   Begin MSComctlLib.ImageList ImageContainer 
      Left            =   4080
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0CCA
            Key             =   "key0a"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1199
            Key             =   "key0b"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1821
            Key             =   "key1a"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1C25
            Key             =   "key1b"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2037
            Key             =   "key2a"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2423
            Key             =   "key2b"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":281C
            Key             =   "key3a"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2C31
            Key             =   "key3b"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":31F2
            Key             =   "key4a"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":35E7
            Key             =   "key4b"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3B88
            Key             =   "key5a"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":3F5C
            Key             =   "key5b"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":433D
            Key             =   "key6a"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4738
            Key             =   "key6b"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4B40
            Key             =   "key7a"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":4F44
            Key             =   "key7b"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5353
            Key             =   "key8a"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5751
            Key             =   "key8b"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5B5B
            Key             =   "key9a"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":5F56
            Key             =   "key9b"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":635E
            Key             =   "keyadda"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":6743
            Key             =   "keyaddb"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":6B38
            Key             =   "keybacka"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":6F33
            Key             =   "keybackb"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":7339
            Key             =   "keycapa"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":77BF
            Key             =   "keycapb"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":7DF6
            Key             =   "keynexta"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":81EC
            Key             =   "keynextb"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":85ED
            Key             =   "keypreva"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":89E7
            Key             =   "keyprevb"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":8F8D
            Key             =   "trayYellow"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":9327
            Key             =   "trayBlue"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":96C1
            Key             =   "trayGray"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":9A5B
            Key             =   "trayGreen"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":9DF5
            Key             =   "trayOlive"
         EndProperty
      EndProperty
   End
   Begin VB.Timer helpTimer 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   2760
      Top             =   2160
   End
   Begin VB.Timer IntTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   2640
   End
   Begin VB.Timer typeTmr 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2400
      Top             =   2400
   End
   Begin VB.Timer msgTmr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   2160
   End
   Begin VB.Label lblTestVersion 
      Caption         =   "test version"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "'yaniv' was added"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AC4493&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   3270
      Width           =   2535
   End
   Begin VB.Image keypic 
      Height          =   1200
      Index           =   13
      Left            =   1830
      Top             =   840
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   0
      Left            =   120
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   1
      Left            =   120
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   2
      Left            =   690
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   3
      Left            =   1260
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   4
      Left            =   120
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   5
      Left            =   690
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   6
      Left            =   1260
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   11
      Left            =   1260
      Top             =   240
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   12
      Left            =   1830
      Top             =   240
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   14
      Left            =   1260
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   9
      Left            =   1260
      Top             =   840
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   8
      Left            =   690
      Top             =   840
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   7
      Left            =   120
      Top             =   840
      Width           =   600
   End
   Begin VB.Image keypic 
      Height          =   600
      Index           =   10
      Left            =   690
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lblReg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trial Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label lblMode 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mode: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu traymenu 
      Caption         =   "TrayMenu"
      Begin VB.Menu mnuShow 
         Caption         =   "Main Window"
      End
      Begin VB.Menu mnuTypes 
         Caption         =   "Type Mode"
         Begin VB.Menu mnuMultiTap 
            Caption         =   "MultiTap"
         End
         Begin VB.Menu mnuSingleTap 
            Caption         =   "SingleTap"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSingleTapPlus 
            Caption         =   "SingleTap+"
         End
         Begin VB.Menu mnuNumbers 
            Caption         =   "Numbers"
         End
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu SettingsMenu 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Begin VB.Menu mnuHelpContents 
            Caption         =   "Help Contents"
            Shortcut        =   {F1}
         End
         Begin VB.Menu sep11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRegister 
            Caption         =   "Buy FastWrite"
         End
         Begin VB.Menu mnuSerialNum 
            Caption         =   "Already have a Serial Number ?"
         End
         Begin VB.Menu bugmnu 
            Caption         =   "Report a bug/Send feedback"
         End
         Begin VB.Menu sep12 
            Caption         =   "-"
         End
         Begin VB.Menu updatesMenu 
            Caption         =   "Check for Updates"
         End
         Begin VB.Menu mnuAbout 
            Caption         =   "About"
         End
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu exitTrayMenu 
         Caption         =   "Exit"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'run shell
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'transparency
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
'window on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'check for another open application
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const SW_RESTORE = 9

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_SHOWWINDOW = &H40
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1

'private
Private unregMsgCount As Integer
Private unregMsgNextCount As Integer

'public
Public errorMsg As String
Public realyExit As Boolean
Public unreg As Boolean



  
' Set a form always on the top.
'
' the form can be specified as a Form or object
' or through its hWnd property
' If OnTop=False the always on the top mode is de-activated.

Private Sub SetAlwaysOnTopMode(hWndOrForm As Variant, Optional ByVal OnTop As Boolean = True)
    Dim hwnd As Long
    ' get the hWnd of the form to be move on top
    If VarType(hWndOrForm) = vbLong Then
        hwnd = hWndOrForm
    Else
        hwnd = hWndOrForm.hwnd
    End If
    If OnTop Then
      SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
      SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub


Private Sub bugmnu_Click()
On Error GoTo errH
  ShellExecute 0&, "open", "mailto:fastwrite@gmail.com?subject=FastWrite bug/feedback", _
      vbNullString, vbNullString, vbNormalFocus
  Exit Sub
errH:
  MsgBox "Could not open a new email message window, please send your but/feedback to:" & vbCrLf & "fastwrite@gmail.com" & vbCrLf & "Thank You", vbOKOnly
End Sub

Private Sub exitTrayMenu_Click()
  realyExit = True
  Unload Me
  
End Sub

Private Sub UnloadPictures()
    keypic(0).Picture = Nothing
    keypic(1).Picture = Nothing
    keypic(2).Picture = Nothing
    keypic(3).Picture = Nothing
    keypic(4).Picture = Nothing
    keypic(5).Picture = Nothing
    keypic(6).Picture = Nothing
    keypic(7).Picture = Nothing
    keypic(8).Picture = Nothing
    keypic(9).Picture = Nothing
    keypic(12).Picture = Nothing
    keypic(11).Picture = Nothing
    keypic(10).Picture = Nothing
    keypic(13).Picture = Nothing
    keypic(14).Picture = Nothing
End Sub

Private Sub LoadPictures()
  
    keypic(0).Picture = ImageContainer.ListImages("key0a").Picture
    keypic(1).Picture = ImageContainer.ListImages("key1a").Picture
    keypic(2).Picture = ImageContainer.ListImages("key2a").Picture
    keypic(3).Picture = ImageContainer.ListImages("key3a").Picture
    keypic(4).Picture = ImageContainer.ListImages("key4a").Picture
    keypic(5).Picture = ImageContainer.ListImages("key5a").Picture
    keypic(6).Picture = ImageContainer.ListImages("key6a").Picture
    keypic(7).Picture = ImageContainer.ListImages("key7a").Picture
    keypic(8).Picture = ImageContainer.ListImages("key8a").Picture
    keypic(9).Picture = ImageContainer.ListImages("key9a").Picture
    keypic(12).Picture = ImageContainer.ListImages("keybacka").Picture
    keypic(11).Picture = ImageContainer.ListImages("keynexta").Picture
    keypic(10).Picture = ImageContainer.ListImages("keypreva").Picture
    keypic(13).Picture = ImageContainer.ListImages("keycapa").Picture
    keypic(14).Picture = ImageContainer.ListImages("keyadda").Picture
   
End Sub

Private Sub Form_Load()
Dim hwnd As Long
Dim sTitle As String
'Dim samebios As Boolean
Dim Registration As New RegClass
Dim i As Integer

  'check for another instance of the program
  'show it and close this one
  If App.PrevInstance = True Then
    sTitle = Me.Caption
    App.Title = "newcopy"
    Me.Caption = "newcopy"
    hwnd = FindWindow(ProductName, sTitle)
    If hwnd <> 0 Then
      ShowWindow hwnd, SW_RESTORE
      SetForegroundWindow hwnd
    End If
    End
  End If
   
  ReadINI
  
'  samebios = IsSameBios
'  unreg = IsUnReged
'  If unreg And samebios Then
'    MsgBox "It seems that you have changed your hardware/computer." & vbCrLf & "Since your registration code is based on your hardware, you will have to reregister." & vbCrLf & "You may try to reenter the code you've received by mail, if it doesn't work just send the mail again." & vbCrLf & "Sorry for the inconvenience...", vbOKOnly, ProductName
'    ReWriteIni
'    lblReg.Caption = "Unregistered (" & UnregDaysLeft & " days left)"
'    unregTmr.Enabled = True
'    mnuRegister.Visible = True
'  ElseIf samebios = False Then
'    If UnregDaysLeft < 0 Then
'      Dim reg As String
'      MsgBox "Registration period is over." & vbCrLf & "If you liked the program, please buy it - it's very cheap :)" & vbCrLf & "Otherwise - Goodbye...", vbOKOnly, ProductName
'      reg = ReadRegistry(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\Dhcp", "Description")
'      If Right$(reg, 1) = "." Then
'        WriteRegistry HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\Dhcp", "Description", ValString, Left$(reg, Len(reg) - 1)
'      End If
'      frmRegistration.Show
'      frmRegistration.onlyMe = True
'      UnloadTrayIcon
'      Unload Me
'      Exit Sub
'    Else
'      lblReg.Caption = "Unregistered (" & UnregDaysLeft & " days left)"
'      unregTmr.Enabled = True
'      mnuRegister.Visible = True
'    End If
'  ElseIf unreg = False Then
'    mnuRegister.Visible = False
'  End If
'  unregMsgCount = 0
'  unregMsgNextCount = 5

  Dim t
  t = Timer
  For i = 1 To 100
    i = i + 1
    Call donothing
  Next 'i
  If Timer - t > 1 Then End
  
  If Registration.CheckOnLoad = False Then
    unregMsgCount = 0
    unregMsgNextCount = 5
    'lblReg.Caption = "TRIAL VERSION"
    'unregTmr.Enabled = True
    'mnuRegister.Visible = True
    mnuRegisterNotOK
    unreg = True
  Else
    mnuRegister.Visible = False
    unreg = False
  End If
  
  OpenDatabase
  
  LoadPictures
    
  
  Dim bytOpacity As Byte
  'Set the transparency level
  bytOpacity = db_transparency
  Dim prevStyle As Long
  
  If bytOpacity < 255 Then 'max opacity - disable alpha
    prevStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, prevStyle Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(Me.hwnd, 0, bytOpacity, LWA_ALPHA)
  End If
  
  If unreg = True Then 'unregistered
    lblReg.Visible = True
    Me.Height = 4200
    'Me.Top = Screen.Height - Me.Height - 450
    '7110
  Else 'registered
    lblReg.Visible = False
    Me.Height = 3970
    'Me.Top = Screen.Height - Me.Height - 450
    '7290
  End If
  Me.Top = Screen.Height - Me.Height - 450
  Me.Left = Screen.Width - Me.Width - 105
  
  traymenu.Visible = False
  lblInfo.Caption = vbNullString
  lblInfo.Visible = True
    
  mainHandler.mainHandlerInit
  
  SwitchMode db_defaultTypeMode
  
  InitTrayIcon
  
  
  If db_startMinimized Then
    Me.Hide
  Else
    Me.Visible = True
  End If
  
  Call SetAlwaysOnTopMode(Me.hwnd, True)
   
  realyExit = False
  
  KeyboardHook
End Sub

Private Sub donothing()
  'just nothing to do
  'and do morenothing too
  Call morenothing
End Sub

Private Sub morenothing()
  'nothing to do here too
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long

Static b As Boolean

  If b = False Then
      b = True
      If Me.ScaleMode = vbPixels Then
        msg = X
      Else
        msg = X / Screen.TwipsPerPixelX
      End If
       
      Select Case msg
        Case WM_RBUTTONUP
          SetForegroundWindow Me.hwnd
          Me.PopupMenu traymenu
        Case WM_LBUTTONUP
          Call keybhook.SwitchMode
      End Select
      b = False
  End If
  
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If realyExit = True Then
  If db_saveLastMode = 1 Then
    WriteINI
  End If
  UnloadTrayIcon
  Unload frmSettings
  'Unload frmRegistration
  Unhook
  CloseDatabase
Else
  Cancel = True
  UnloadPictures
  Me.Visible = False
  Me.Hide
End If
End Sub


Private Sub helpTimer_Timer()
  helpTimer.Enabled = False
  Call ShowHelp
End Sub

Private Sub IntTimer_Timer()
  IntTimer.Enabled = False
  'Debug.Print "in IntTimer"
  Call keybhook.HandleKeyPress
End Sub

Public Function GetVer() As String
  GetVer = App.Major & "." & App.Minor & ".0." & App.Revision
End Function

Private Sub langTimer_Timer()
  DisableIfNotEnglish
End Sub

Private Sub mnuAbout_Click()
Dim msg As String

  msg = ProductName & " " & GetVer & _
         vbCrLf & "Created by Yaniv Kalsky"
  If unreg = True Then
    msg = msg & vbCrLf & "This is a Trial Version"
  Else
    msg = msg & vbCrLf & "Registerd version: " & GetSetting("FastWrite", vbNull, "serialNumber", "")
  End If
  MsgBox msg, vbOKOnly, ProductName
End Sub

Private Sub mnuHelpContents_Click()
  Call ShowHelp
End Sub

Private Sub UncheckAllModes()
  mnuMultiTap.Checked = False
  mnuSingleTap.Checked = False
  mnuSingleTapPlus.Checked = False
  mnuNumbers.Checked = False
End Sub

Private Sub mnuMultiTap_Click()
  UncheckAllModes
  mnuMultiTap.Checked = True
  Call SwitchMode(TypeModes_MULTITAP)
End Sub

Private Sub mnuNumbers_Click()
  UncheckAllModes
  mnuNumbers.Checked = True
  Call SwitchMode(TypeModes_NUMBERS)
End Sub

'Private Sub mnuOff_Click()
'  UncheckAllModes
'  mnuOff.Checked = True
'  Call SwitchMode(TypeModes_OFF)
'
'End Sub

Public Sub mnuRegister_Click()
Dim Registration As New RegClass
  'frmRegistration.Show vbModal
  If Registration.BuyFastWrite Then
    unreg = unreg
  End If
End Sub

Public Sub mnuRegisterOK()
    lblReg.Visible = False
    mnuRegister.Visible = False
    mnuSerialNum.Visible = False
    unregTmr.Enabled = False
    unregTmr.Enabled = True
    Me.Height = 3970
    Me.Top = Screen.Height - Me.Height - 450
    Me.Left = Screen.Width - Me.Width - 105
    frmSettings.defaultOnReg.Visible = False
    frmSettings.defaultsFrame.Enabled = True
    frmSettings.autoaddOnReg.Visible = False
    frmSettings.AutoAddChk.Enabled = True
    unreg = False
End Sub

Public Sub mnuRegisterNotOK()
    lblReg.Visible = True
    mnuRegister.Visible = True
    mnuSerialNum.Visible = True
    unregTmr.Enabled = True
    Me.Height = 4200
    Me.Top = Screen.Height - Me.Height - 450
    Me.Left = Screen.Width - Me.Width - 105
    frmSettings.defaultOnReg.Visible = True
    frmSettings.defaultsFrame.Enabled = False
    frmSettings.autoaddOnReg.Visible = True
    frmSettings.AutoAddChk.Enabled = False
    unreg = True
End Sub

Private Sub mnuSerialNum_Click()
Dim serial As String
Dim Registration As New RegClass

  serial = InputBox("If you've already got a serial number (from online purchase)," & vbCrLf & "please enter it exactly as shown on your receipt, including case (uppercase or lowercase).", "FastWrite - Activation", "fwkeyfw")
  If serial > "" And serial <> "fwkeyfw" Then
    Call Registration.ActivateFastWrite(serial)
  End If
End Sub

Private Sub mnuShow_Click()
  LoadPictures
  Me.Visible = True
  Me.Show
End Sub

Private Sub mnuSingleTap_Click()
  UncheckAllModes
  mnuSingleTap.Checked = True
  Call SwitchMode(TypeModes_SINGLETAP)
End Sub

Private Sub mnuSingleTapPlus_Click()
  UncheckAllModes
  mnuSingleTapPlus.Checked = True
  Call SwitchMode(TypeModes_SINGLETAPPLUS)
End Sub

Private Sub msgTmr_Timer()
  msgTmr.Enabled = False
  If errorMsg > "" Then
    MsgBox errorMsg, vbCritical, ProductName
    errorMsg = vbNullString
  End If
End Sub

Public Sub ChangeOpacity(Value As Integer)
  Dim bytOpacity As Byte
  'Set the transparency level
  bytOpacity = Value
  If (Value < 255) Then
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(Me.hwnd, 0, bytOpacity, LWA_ALPHA)
  Else
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, &H40101)
  End If
  Call SetAlwaysOnTopMode(Me.hwnd, True)
End Sub

Private Sub SettingsMenu_Click()
  frmSettings.Show
End Sub

Private Sub tmrClrInfo_Timer()
  tmrClrInfo.Enabled = False
  lblInfo.Caption = vbNullString
End Sub

Private Sub typeTmr_Timer()
  typeTmr.Enabled = False
  'debug.print Time
  Call TypeTmrHandler
End Sub

Public Sub UpdateCaption(newCap As String)
  'Me.Caption = productname ' (" & trim$(newCap) & ")"
  Me.lblMode.Caption = Trim$(newCap)
End Sub

Public Sub ShowHelp(Optional showSettings As Boolean = False)
On Error GoTo errH
Dim i As InternetExplorer
'Dim help As String

  Set i = New InternetExplorer
  i.Visible = False
  Me.MousePointer = vbHourglass
  If showSettings Then
    i.Navigate App.Path & "\help\help.htm#settings"
  Else
    i.Navigate App.Path & "\help\help.htm"
  End If
  Do While i.ReadyState <> READYSTATE_COMPLETE
    DoEvents
  Loop
  
  i.AddressBar = False
  i.ToolBar = False
  i.Resizable = False
  i.StatusBar = False

  i.Width = 800
  i.Height = 500
  i.Top = 100
  i.Left = 100
  
  Me.MousePointer = vbNormal
  i.Visible = True
  Set i = Nothing
  Exit Sub
errH:
  If Not (i Is Nothing) Then
    i.Quit
    Set i = Nothing
    Me.MousePointer = vbNormal
    MsgBox "Could not open the help file, please open it through the FastWrite shortcuts group", vbOKOnly, "FastWrite - Help"
  End If
End Sub

Public Sub ColorKey(key As Integer, Value As Boolean)
  If Me.Visible = False Then Exit Sub

  Select Case key
    Case vbKeyNumpad0
      If Value = False Then
        keypic(0).Picture = ImageContainer.ListImages("key0a").Picture
      Else
        keypic(0).Picture = ImageContainer.ListImages("key0b").Picture
      End If
      keypic(0).Refresh
    Case vbKeyNumpad1
      If Value = False Then
        keypic(1).Picture = ImageContainer.ListImages("key1a").Picture
      Else
        keypic(1).Picture = ImageContainer.ListImages("key1b").Picture
      End If
      keypic(1).Refresh
    Case vbKeyNumpad2
      If Value = False Then
        keypic(2).Picture = ImageContainer.ListImages("key2a").Picture
      Else
        keypic(2).Picture = ImageContainer.ListImages("key2b").Picture
      End If
      keypic(2).Refresh
    Case vbKeyNumpad3
      If Value = False Then
        keypic(3).Picture = ImageContainer.ListImages("key3a").Picture
      Else
        keypic(3).Picture = ImageContainer.ListImages("key3b").Picture
      End If
      keypic(3).Refresh
    Case vbKeyNumpad4
      If Value = False Then
        keypic(4).Picture = ImageContainer.ListImages("key4a").Picture
      Else
        keypic(4).Picture = ImageContainer.ListImages("key4b").Picture
      End If
      keypic(4).Refresh
    Case vbKeyNumpad5
      If Value = False Then
        keypic(5).Picture = ImageContainer.ListImages("key5a").Picture
      Else
        keypic(5).Picture = ImageContainer.ListImages("key5b").Picture
      End If
      keypic(5).Refresh
    Case vbKeyNumpad6
      If Value = False Then
        keypic(6).Picture = ImageContainer.ListImages("key6a").Picture
      Else
        keypic(6).Picture = ImageContainer.ListImages("key6b").Picture
      End If
      keypic(6).Refresh
    Case vbKeyNumpad7
      If Value = False Then
        keypic(7).Picture = ImageContainer.ListImages("key7a").Picture
      Else
        keypic(7).Picture = ImageContainer.ListImages("key7b").Picture
      End If
      keypic(7).Refresh
    Case vbKeyNumpad8
      If Value = False Then
        keypic(8).Picture = ImageContainer.ListImages("key8a").Picture
      Else
        keypic(8).Picture = ImageContainer.ListImages("key8b").Picture
      End If
      keypic(8).Refresh
    Case vbKeyNumpad9
      If Value = False Then
        keypic(9).Picture = ImageContainer.ListImages("key9a").Picture
      Else
        keypic(9).Picture = ImageContainer.ListImages("key9b").Picture
      End If
      keypic(9).Refresh
    Case vbKeySubtract
      If Value = False Then
        keypic(12).Picture = ImageContainer.ListImages("keybacka").Picture
      Else
        keypic(12).Picture = ImageContainer.ListImages("keybackb").Picture
      End If
      keypic(12).Refresh
    Case vbKeyMultiply
      If Value = False Then
        keypic(11).Picture = ImageContainer.ListImages("keynexta").Picture
      Else
        keypic(11).Picture = ImageContainer.ListImages("keynextb").Picture
      End If
      keypic(11).Refresh
    Case vbKeyDivide
      If Value = False Then
        keypic(10).Picture = ImageContainer.ListImages("keypreva").Picture
      Else
        keypic(10).Picture = ImageContainer.ListImages("keyprevb").Picture
      End If
      keypic(10).Refresh
    Case vbKeyAdd
      If Value = False Then
        keypic(13).Picture = ImageContainer.ListImages("keycapa").Picture
      Else
        keypic(13).Picture = ImageContainer.ListImages("keycapb").Picture
      End If
      keypic(13).Refresh
    Case vbKeyDecimal
      If Value = False Then
        keypic(14).Picture = ImageContainer.ListImages("keyadda").Picture
      Else
        keypic(14).Picture = ImageContainer.ListImages("keyaddb").Picture
      End If
      keypic(14).Refresh
  End Select
End Sub

Public Sub SetTrayIcon(mode As TypeModesEnum)
Dim iconKey As String

  Select Case mode
    Case TypeModes_MULTITAP
      iconKey = "trayBlue"
    Case TypeModes_NUMBERS
      iconKey = "trayYellow"
    Case TypeModes_SINGLETAP
      iconKey = "trayGreen"
    Case TypeModes_SINGLETAPPLUS
      iconKey = "trayOlive"
    Case TypeModes_OFF
      iconKey = "trayGray"
  End Select
  sysTrayIcon.hIcon = ImageContainer.ListImages(iconKey).Picture
  Me.Icon = ImageContainer.ListImages(iconKey).Picture
  
  If db_saveLastMode = 1 Then
    db_defaultTypeMode = mode
  End If
  UpdateTrayIcon
End Sub

Public Sub ShowInfo(info As String, Optional endless As Boolean = False)
  lblInfo.Caption = info
  If endless = False Then tmrClrInfo.Enabled = True
End Sub

Private Sub unregTmr_Timer()

  If unregMsgCount = unregMsgNextCount Then
    If unreg = False Then 'registered
      'verify again that this program is legal
      Dim Registration As New RegClass
      If Registration.CheckOnLoad = False Then
        'not legal!!!
        Registration.Unregister
      Else
        unregTmr.Enabled = False
        Exit Sub
      End If
    End If
    unregTmr.Enabled = False
    MsgBox "This is an unregistered version of FastWrite" & vbCrLf & "Please register and have this annoying message disappear...", vbOKOnly, ProductName
    'Randomize
    unregMsgNextCount = 20 '10 minutes nag
    unregTmr.Enabled = True
  Else
    unregMsgCount = unregMsgCount + 1
  End If
  
  
End Sub


Private Sub CheckUpdatesNow()
On Error GoTo eH
Dim Source1 As String
Dim msg As Integer

  Source1 = Inet1.OpenURL("http://fastwrite.bambuk.co.il/setup/lastver.txt")
  If LenB(Source1) > 0 Then
    If Source1 > GetVer Then
        msg = MsgBox("There is a new version (" & Source1 & "), your version is " & GetVer & vbCrLf _
                   & "Do you want to go to FastWrite website to see the changes?", vbYesNo, "FastWrite Update Available")
        If msg = vbYes Then
          ShellExecute 0&, "open", "http://beam.to/fastwrite", _
                       vbNullString, vbNullString, vbNormalFocus
        End If
    Else
        MsgBox "You have the latest version", vbOKOnly, "FastWrite"
    End If
  Else
eH:
    msg = MsgBox("Could not connect FastWrite server, try again later...", vbInformation, "FastWrite Update")
    Err.Clear
  End If
  
End Sub

Private Sub updatesMenu_Click()
  CheckUpdatesNow
End Sub
