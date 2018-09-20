VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   ClipControls    =   0   'False
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton helpBtn 
      Caption         =   "Help"
      Height          =   375
      Left            =   7680
      TabIndex        =   24
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Functionality"
      Height          =   975
      Left            =   4680
      TabIndex        =   19
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox AutoAddChk 
         Caption         =   "Auto add new words"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox FastDelChk 
         Caption         =   "Fast Delete (remembers only the current word)"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label autoaddOnReg 
         Caption         =   "(Only in registered version)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Startup"
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox StartWinChk 
         Caption         =   "Start with Windows"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox StartMiniChk 
         Caption         =   "Start Minimized (for advanced typers)"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox UpdatesChk 
         Caption         =   "Check for updates on startup"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Word's priority database "
      ClipControls    =   0   'False
      Height          =   855
      Left            =   4680
      TabIndex        =   12
      Top             =   1200
      Width           =   4455
      Begin VB.CommandButton btnLearnPrio 
         Caption         =   "Teach FastWrite your words"
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton btnDefaultPrio 
         Caption         =   "Set to default"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Tips"
      Enabled         =   0   'False
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame defaultsFrame 
      Caption         =   "Default Type Mode"
      ClipControls    =   0   'False
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   4455
      Begin VB.CheckBox saveLastChk 
         Caption         =   "or Save last used mode"
         Height          =   255
         Left            =   1920
         TabIndex        =   23
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton opTypeMode 
         Caption         =   "Off"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton opTypeMode 
         Caption         =   "Numbers"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton opTypeMode 
         Caption         =   "SingleTap+"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton opTypeMode 
         Caption         =   "SingleTap"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton opTypeMode 
         Caption         =   "MultiTap"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   280
         Width           =   1095
      End
      Begin VB.Label defaultOnReg 
         Caption         =   "These options are only available in the registered version"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Image OffImage 
         Height          =   240
         Left            =   120
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image NumImage 
         Height          =   240
         Left            =   120
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image SinglePImage 
         Height          =   240
         Left            =   120
         Top             =   960
         Width           =   240
      End
      Begin VB.Image SingleImage 
         Height          =   240
         Left            =   120
         Top             =   600
         Width           =   240
      End
      Begin VB.Image MultiImage 
         Height          =   240
         Left            =   120
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.CommandButton cancelBtn 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton okBtn 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main Window Transparency"
      ClipControls    =   0   'False
      Height          =   735
      Left            =   4680
      TabIndex        =   0
      Top             =   2400
      Width           =   4455
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         SmallChange     =   5
         Min             =   140
         Max             =   255
         SelStart        =   225
         TickStyle       =   3
         Value           =   225
      End
      Begin VB.Label Label2 
         Caption         =   "Opaque"
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Transparent"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public onLoad As Boolean

Private Sub btnDefaultPrio_Click()
'Dim i As Long

  If vbYes = MsgBox("Are you sure you want to resore the default priority database ?", vbYesNo, "Set default priorities") Then
    Me.MousePointer = vbHourglass
    DB.Execute "UPDATE words " & "SET priority=default_pri"
  
    mainHandler.StartNewWord
    Me.MousePointer = vbNormal
  End If
End Sub

Private Sub btnLearnPrio_Click()
  frmLearn.Show
End Sub

Private Sub cancelBtn_Click()
  Me.Hide
End Sub

Private Sub Form_Load()
  onLoad = True
  Slider1.Value = db_transparency
  opTypeMode(db_defaultTypeMode).Value = True
  
  MultiImage.Picture = frmMain.ImageContainer.ListImages("trayBlue").Picture
  SingleImage.Picture = frmMain.ImageContainer.ListImages("trayGreen").Picture
  SinglePImage.Picture = frmMain.ImageContainer.ListImages("trayOlive").Picture
  NumImage.Picture = frmMain.ImageContainer.ListImages("trayYellow").Picture
  OffImage.Picture = frmMain.ImageContainer.ListImages("trayGray").Picture
  
  If frmMain.unreg = True Then
    defaultOnReg.Visible = True
    defaultsFrame.Enabled = False
    autoaddOnReg.Visible = True
    AutoAddChk.Enabled = False
  End If
   
  onLoad = False
End Sub

Private Sub UpdateDBValues()
Dim i As Integer

  db_transparency = Slider1.Value
  
  For i = 0 To 4
    If opTypeMode(i).Value = True Then
      db_defaultTypeMode = i
      Exit For
    End If
  Next 'i
  
  db_updatesOnStart = UpdatesChk.Value
  db_startMinimized = StartMiniChk.Value
  If db_startWithWindows <> StartWinChk.Value Then
    db_startWithWindows = StartWinChk.Value
    If db_startWithWindows Then
      RunAtStartup App.Title, App.Path & "\" & App.EXEName & ".EXE"
    Else
      RemoveFromStartup App.Title, App.Path & "\" & App.EXEName & ".EXE"
    End If
  End If
  
  db_fastDelete = FastDelChk.Value
  db_autoAddNew = AutoAddChk.Value
  db_saveLastMode = saveLastChk.Value
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Unload frmLearn
End Sub

Private Sub helpBtn_Click()
  Call frmMain.ShowHelp(True)
End Sub

Private Sub okBtn_Click()
  UpdateDBValues
  WriteINI
  Me.Hide
End Sub

Private Sub saveLastChk_Click()
  If saveLastChk.Value = 0 Then
    opTypeMode(0).Enabled = True
    opTypeMode(1).Enabled = True
    opTypeMode(2).Enabled = True
    opTypeMode(3).Enabled = True
    opTypeMode(4).Enabled = True
  Else
    opTypeMode(0).Enabled = False
    opTypeMode(1).Enabled = False
    opTypeMode(2).Enabled = False
    opTypeMode(3).Enabled = False
    opTypeMode(4).Enabled = False
  End If
End Sub

Private Sub Slider1_Change()
  If onLoad = False Then
    Call frmMain.ChangeOpacity(Slider1.Value)
  End If
End Sub

