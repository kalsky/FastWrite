Attribute VB_Name = "iniHandler"
Option Explicit

'Public Type Record   ' Define user-defined type.
'   user As String * 30
'   pass As String * 20
'End Type



Private MydsEncrypt As dsEncrypt
'Private EncUpgraded As Boolean

'details
Public db_transparency As Integer
Public db_defaultTypeMode As TypeModesEnum
Public db_updatesOnStart As Byte
Public db_startMinimized As Byte
Public db_startWithWindows As Byte
Public db_fastDelete As Byte
Public db_autoAddNew As Byte
Public db_saveLastMode As Byte
''''''''''''''''''''''''''''''''''''


Public Sub ReadINI()
Dim s As String
'Dim l As Integer
Dim fileN As Byte
On Error GoTo sof

  If Dir$(App.Path & "\" & "me.dat") > "" Then
    Set MydsEncrypt = New dsEncrypt
    MydsEncrypt.KeyString = ("KATHER")
    fileN = FreeFile
    Open App.Path & "\" & "me.dat" For Input As #fileN
      
      Line Input #fileN, s
      db_transparency = CInt(Decrypt(s))
      Line Input #fileN, s
      db_defaultTypeMode = CByte(Decrypt(s))
      
      If Not EOF(fileN) Then 'added on version 0.90.0.2
        Line Input #fileN, s
        db_updatesOnStart = CByte(Decrypt(s))
        frmSettings.UpdatesChk.Value = db_updatesOnStart
        Line Input #fileN, s
        db_startMinimized = CByte(Decrypt(s))
        frmSettings.StartMiniChk.Value = db_startMinimized
        Line Input #fileN, s
        db_startWithWindows = CByte(Decrypt(s))
        frmSettings.StartWinChk.Value = db_startWithWindows
      Else
        db_updatesOnStart = 0
        db_startMinimized = 0
        db_startWithWindows = 0
      End If
      
      If Not EOF(fileN) Then 'added on version 0.90.0.3
        Line Input #fileN, s
        db_fastDelete = CByte(Decrypt(s))
        frmSettings.FastDelChk.Value = db_fastDelete
        Line Input #fileN, s
        db_autoAddNew = CByte(Decrypt(s))
        frmSettings.AutoAddChk.Value = db_autoAddNew
        Line Input #fileN, s
        db_saveLastMode = CByte(Decrypt(s))
        frmSettings.saveLastChk.Value = db_saveLastMode
      Else
        db_fastDelete = 1
        db_autoAddNew = 0
        db_saveLastMode = 0
      End If
       
    Close #fileN
    Set MydsEncrypt = Nothing
    
  Else 'set defaults
    db_transparency = 225
    db_defaultTypeMode = TypeModes_SINGLETAP
    db_updatesOnStart = False
    'WriteRegistry HKEY_LOCAL_MACHINE, "SOFTWARE\FastWrite\", "reg", ValString, GetUnregKey
    WriteINI
  End If
  Exit Sub
sof:
  Close #fileN
End Sub

Public Sub WriteINI()
Dim fileN As Byte

  fileN = FreeFile
  
  If MydsEncrypt Is Nothing Then
    Set MydsEncrypt = New dsEncrypt
    MydsEncrypt.KeyString = ("KATHER")
  End If
  
  Open App.Path & "\" & "me.dat" For Output As #fileN
    Print #fileN, Encrypt(db_transparency)
    Print #fileN, Encrypt(db_defaultTypeMode)
    Print #fileN, Encrypt(db_updatesOnStart)
    Print #fileN, Encrypt(db_startMinimized)
    Print #fileN, Encrypt(db_startWithWindows)
    Print #fileN, Encrypt(db_fastDelete)
    Print #fileN, Encrypt(db_autoAddNew)
    Print #fileN, Encrypt(db_saveLastMode)
  Close #fileN
  Set MydsEncrypt = Nothing
End Sub

Public Sub ReWriteIni()
  Kill App.Path & "\" & "me.dat"
  'WriteRegistry HKEY_LOCAL_MACHINE, "SOFTWARE\FastWrite\", "reg", ValString, GetUnregKey
  WriteINI
End Sub

Public Function Encrypt(ByVal Plain As String) As String

  Encrypt = MydsEncrypt.Encrypt(Plain)

End Function

Public Function Decrypt(ByVal Encrypted As String) As String

  Decrypt = MydsEncrypt.Encrypt(Encrypted)

End Function



