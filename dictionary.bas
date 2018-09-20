Attribute VB_Name = "dicMdl"
Option Explicit

Private letters() As String
Public Const LTTRS As String = "8,8,8,9,9,9,4,4,4,5,5,5,6,6,6,1,1,1,1,2,2,2,3,3,3,3"
Public DB As Database

Public Sub AddNewWord(w As String)

Dim rs As Recordset

  If LenB(w) = 0 Then Exit Sub
  If InStr(LTTRS, Left$(w, 1)) And InStr(PUNCT, w) Then Exit Sub
  
  letters = Split(LTTRS, ",")
  
  'Debug.Print w; " is trying to be added"
  Set rs = DB.OpenRecordset("select * from words where word=" & Chr$(34) & LCase$(Trim$(w)) & Chr$(34))
  On Error Resume Next
  If rs.RecordCount = 0 Then
    Dim key As String
    Dim i As Integer
    Dim ww As String
    Dim lenw As Integer
    
    ww = w
    w = LCase$(Trim$(w))
    If wordOK(w) = True Then
      key = vbNullString
      lenw = Len(w)
      For i = 1 To lenw
        key = key & getKey(Mid$(w, i, 1))
      Next 'i
      With rs
        .AddNew
        !word = ww
        !key = key
        !priority = 1
        !default_pri = 0
        .Update
      End With
      'Debug.Print ww & " was added to the dictionary"
      frmMain.ShowInfo "'" & ww & "' was added"
    Else
      frmMain.errorMsg = "Word must be with only letters (a-z;A-Z) and without spaces"
    End If
  Else
    If db_autoAddNew = 0 Then
    'Debug.Print w & " - Word already exists"
      frmMain.ShowInfo "'" & w & "' already exists"
    End If
  End If
End Sub

Private Function wordOK(ByVal w As String) As Boolean
Dim i As Integer
On Error GoTo errH

  For i = 33 To 126
    w = Replace(w, ChrW$(i), vbNullString)
  Next 'i
  
  If LenB(w) > 0 Then
    wordOK = False
  Else
    wordOK = True
  End If
  Exit Function
errH:
wordOK = False
End Function

Private Function getKey(c As String) As String
Dim dif As Integer
On Error GoTo errH

  dif = AscW(c) - AscWa
  If ((dif < 25) And (dif >= 0)) Then
    getKey = letters(dif)
  End If
  If InStr(PUNCT, c) Then
errH:
    getKey = 7
  End If
  
End Function

Public Function CalculateKey(w As String) As String
Dim key As String
Dim i As Integer
Dim lenw As Integer

  letters = Split(LTTRS, ",")
On Error GoTo errH
  key = vbNullString
  If wordOK(w) = True Then
    lenw = Len(w)
    For i = 1 To lenw
      key = key & getKey(Mid$(w, i, 1))
    Next 'i
  End If
errH:
  CalculateKey = key
End Function

Public Sub OpenDatabase()
On Error Resume Next

Dim MydsEncrypt As New dsEncrypt
MydsEncrypt.KeyString = ("KATHER")

'the password is encrypted!!
    Set DB = DBEngine.OpenDatabase(App.Path & "\dic.dat", False, False, _
        "MS Access;PWD=" & MydsEncrypt.Encrypt("Bsug{tueNblHs"))
End Sub

Public Sub CloseDatabase()
On Error Resume Next
    DB.Close
End Sub


