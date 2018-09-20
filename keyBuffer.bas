Attribute VB_Name = "keyBuffer"
Option Explicit

Private Const BUFF_SIZE = 500

Private Buffer(0 To BUFF_SIZE) As Integer
Private startIndex As Integer
Private endIndex As Integer
Private inKeys As Integer

Public Sub CleanBuffer()
  startIndex = 1
  endIndex = 1
  inKeys = 0
End Sub

Public Function IsEmpty() As Boolean
  IsEmpty = (startIndex = endIndex)
End Function

Public Function IsFull() As Boolean
  IsFull = (((endIndex + 1) Mod BUFF_SIZE) = startIndex)
End Function

Public Function AddKey(key As Long) As Boolean
  If IsFull = False Then
    Buffer(endIndex) = key
    If ((endIndex + 1) Mod BUFF_SIZE) <> startIndex Then
      endIndex = (endIndex + 1) Mod BUFF_SIZE
    End If
    AddKey = True
    inKeys = inKeys + 1
    'Debug.Print endIndex, startIndex
  Else
    'Debug.Print "addkey failed"
    AddKey = False
  End If
End Function

Public Function ReadKey() As Integer
  If IsEmpty = False Then
    ReadKey = Buffer(startIndex)
    startIndex = (startIndex + 1) Mod BUFF_SIZE
    inKeys = inKeys - 1
  Else
    ReadKey = -1
  End If
End Function

Public Function CopyKey() As Integer
  If IsEmpty = False Then
    CopyKey = Buffer(startIndex)
  Else
    CopyKey = -1
  End If
End Function


Public Function HasSome() As Boolean
  HasSome = (inKeys > 0)
End Function

Public Function GetUnfreeSize() As Integer
  GetUnfreeSize = inKeys
End Function
