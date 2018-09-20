VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLearn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show me your world..."
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
   ClipControls    =   0   'False
   Icon            =   "frmLearn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton browseBtn 
      Caption         =   "Browse"
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox filesLocation 
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   7575
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.OptionButton opAdd 
      Caption         =   "Add to the current priorities database"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Value           =   -1  'True
      Width           =   3975
   End
   Begin VB.OptionButton opNew 
      Caption         =   "Start a new priorities database"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   4560
      Width           =   3855
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton btnLearn 
      Caption         =   "Learn text"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtLearn 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
   Begin VB.Label Label2 
      Caption         =   "or Choose a TEXT files for automatic learn"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Put in the following text box a piece of text that represents your daily use"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmLearn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub browseBtn_Click()

  On Error GoTo e_Trap
  
  Dim sFile$
  
  ' call the function.  a zero length return indicates that the user did not select a file name.
  sFile = GetOpenName(App.Path)
  
  If LenB(sFile) > 0 Then
    filesLocation.text = sFile
  End If
  
e_Trap:
    Exit Sub
    Resume
End Sub

Private Sub btnClose_Click()
On Error Resume Next
  Unload Me
End Sub

Private Function CleanWord(w As String) As String
On Error GoTo errH

  w = LCase$(w)
  Do While InStr(PUNCT, Right$(w, 1))
    w = Mid$(w, 1, Len(w) - 1)
  Loop
  Do While InStr(PUNCT, Left$(w, 1))
    w = Mid$(w, 2)
  Loop
  CleanWord = Trim$(w)
  Exit Function
errH:
  CleanWord = ""
End Function

Private Sub AnalyzeText(text As String)
Dim i As Long
Dim sp() As String
Dim usp As Long
Dim key As String
Dim rs As Recordset

On Error GoTo errH

  sp = Split(text, " ")
  usp = UBound(sp) - 1
  If frmMain.unreg = True Then
    If usp > 2 Then
      usp = 2
      MsgBox "This is a Trial version of FastWrite so only the first 2 words will be learned", vbOKOnly
    End If
  End If
  PB.Max = usp + 1
  PB.Value = 0
  Me.MousePointer = vbHourglass
  On Error Resume Next
  'With frmMain.Data1
    For i = 0 To usp
      sp(i) = CleanWord(sp(i))
      If LenB(sp(i)) > 0 Then
      Set rs = DB.OpenRecordset("select word,priority from words where word=" & Chr$(34) & sp(i) & Chr$(34))
      With rs
        If rs.RecordCount = 1 Then
          rs.Edit
          rs.Fields(1) = rs.Fields(1) + 1
          rs.Update
        Else
          key = CalculateKey(sp(i))
          If LenB(key) > 0 Then
            Dim rs1 As Recordset
            Set rs1 = DB.OpenRecordset("select * from words where word='" & sp(i) & "'")
              With rs1
                .MoveLast
                .MoveFirst
                .AddNew
                !word = sp(i)
                !priority = 1
                !default_pri = 0
                !key = key
                .Update
                .Close
              End With
          End If
        End If
      End With
      rs.Close
      End If
      If i Mod 10 = 0 Then
        PB.Value = i
      End If
    Next 'i
    
errH:
  rs.Close
  PB.Value = PB.Max
  Me.MousePointer = vbNormal
End Sub

Private Sub btnLearn_Click()

Dim text As String
Dim files() As String
Dim i As Integer
Dim fileN As Byte
Dim oneline As String

  If opNew.Value = True Then
    If vbYes = MsgBox("All previous priorities will be deleted (since installation)" & vbCrLf & "Are you sure ? (default priorities can still be restored)", vbYesNo, ProductName) Then
      Me.MousePointer = vbHourglass
     
      DB.Execute "UPDATE words " & "SET priority=0"

      Me.MousePointer = vbNormal
    Else
      Exit Sub
    End If
  End If
  
  txtLearn.text = txtLearn.text & " "
  text = Replace(txtLearn.text, vbCrLf, " ")
  If LenB(text) > 0 Then AnalyzeText text
  
  
  files = Split(filesLocation & "; ", "; ")
  oneline = ""
  i = 0
  If filesLocation <> "" Then
  'For i = 0 To UBound(files) - 1
    fileN = FreeFile
    text = ""
    Open files(i) For Input As #fileN
    On Error GoTo closeFile
      Do While Not EOF(fileN)
        Line Input #fileN, oneline
        text = text & " " & oneline
      Loop
closeFile:
    Close #fileN
    On Error Resume Next
    If LenB(text) > 0 Then AnalyzeText text
  End If
     
  'Next 'i
  
  mainHandler.StartNewWord
End Sub

Private Sub filesLocation_Change()
On Error GoTo errH
  If filesLocation.text <> "" Then
    btnLearn.Enabled = True
  ElseIf txtLearn.text = "" Then
    btnLearn.Enabled = False
  End If
errH:
End Sub

Private Sub txtLearn_Change()
On Error GoTo errH
  If txtLearn.text <> "" Then
    btnLearn.Enabled = True
  ElseIf filesLocation.text = "" Then
    btnLearn.Enabled = False
  End If
errH:
End Sub
