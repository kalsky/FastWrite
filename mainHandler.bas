Attribute VB_Name = "mainHandler"
Option Explicit

'local
Private Const CONV As String = "7,8,9,4,5,6,1,2,3"
Private ftiter As FTiterator
Private currentString As String
Private lastString As String
Private Enum LettersPosEnum
  LettersPos_ABC = 1
  LettersPos_DEF = 2
  LettersPos_GHI = 3
  LettersPos_JKL = 4
  LettersPos_MNO = 5
  LettersPos_PQRS = 6
  LettersPos_TUV = 7
  LettersPos_WXYZ = 8
  LettersPos_DOT = 0
  LettersPos_SPACE = 10
  LettersPos_PLUS = 13
  LettersPos_MINUS = 12
  LettersPos_MULTIPLY = 9
  LettersPos_DIVIDE = 11
  LettersPos_ADDWORD = 14
End Enum


Public Sub mainHandlerInit()
  Set ftiter = New FTiterator
  currentString = vbNullString
  lastString = vbNullString
End Sub

Public Function getDeepLevel() As Integer
  getDeepLevel = ftiter.deepnessLevel
End Function

Public Function getLastString() As String
  getLastString = lastString
End Function

Public Sub StartNewWord()
  ftiter.newword
End Sub

Public Sub SetPrevWord(word As String)
Dim key As Integer
Dim letters() As String
Dim dif As Integer
Dim i As Integer
Dim char As String
Dim lenw As Integer
  
  ftiter.newword
  currentString = vbNullString
  lenw = Len(word)
  
  letters = Split(LTTRS, ",")
  For i = 1 To lenw
    char = Mid$(word, i, 1)
    If char >= "a" And char <= "z" Then
        
        dif = AscW(char) - AscWa
        If (dif < 25) And (dif >= 0) Then
            key = letters(dif)
            ftiter.nextLevel (key)
        End If
    End If
  Next 'i
  Call ftiter.MoveTo(word)
  currentString = word
End Sub


Public Function handleClick(index As Integer, newword As Boolean) As String
Dim key As Integer
Dim letters() As String
Dim exactSize As Boolean
  
  If newword = True Then
    Call ftiter.newword
  End If
  
  
  If GetTypeMode = TypeModes_SINGLETAPPLUS Then
    exactSize = False
  Else
    exactSize = True
  End If
 
  Select Case index
    Case LettersPos_MULTIPLY  'flip forward
      If getDeepLevel > 0 Then
        If ftiter.MoveNext(exactSize) Then
            currentString = ftiter.Current(exactSize)
        Else
            Call ftiter.StartOver
            currentString = ftiter.Current(exactSize)
        End If
      End If
           
    Case LettersPos_DIVIDE 'flip backwards
      If getDeepLevel > 0 Then
        If ftiter.MoveBack Then
            currentString = ftiter.Last
        End If
      End If
      
    Case LettersPos_MINUS 'simulate backspace

      ftiter.prevLevel
      If InStr(currentString, vbTab) Then
        lastString = Left$(currentString, InStr(currentString, vbTab) - 1)
      Else
        lastString = vbNullString
      End If
      currentString = vbNullString
      
    Case LettersPos_SPACE 'simulate spacebar
      Call ftiter.addPriority(Replace(currentString, vbTab, vbNullString))
      Call ftiter.newword
      lastString = currentString
      currentString = vbNullString
      
    Case Else
      
      If index >= LettersPos_ABC And index <= LettersPos_WXYZ Then
           
        letters = Split(CONV, ",")
        key = letters(index)
        ftiter.nextLevel (key)
        If ftiter.MoveNext() Then
            currentString = ftiter.Current(exactSize)
        Else
            Select Case index
            Case LettersPos_DEF
              currentString = currentString & "e"
            Case LettersPos_PQRS
              currentString = currentString & "s"
            Case Else
              currentString = currentString & "?"
              frmMain.ShowInfo "can't find match"
            End Select
            currentString = Replace(currentString, vbTab, vbNullString)
        End If

      Else
        If index = LettersPos_DOT Then
          Call ftiter.addPriority(Replace(currentString, vbTab, vbNullString))
          Call ftiter.newword
          key = 7
          ftiter.nextLevel (key)
          If ftiter.MoveNext(exactSize) Then
              currentString = ftiter.Current(exactSize)
          Else
              ftiter.prevLevel
          End If
        End If
      End If
      
  End Select
  
addTab:
  If (Len(currentString) > ftiter.deepnessLevel) And (exactSize = False) Then
    currentString = Left$(currentString, ftiter.deepnessLevel) & vbTab & Mid$(currentString, ftiter.deepnessLevel + 1)
  ElseIf (Len(currentString) > ftiter.deepnessLevel) Then
    currentString = Left$(currentString, ftiter.deepnessLevel)
  End If
  
exitfunc:
  handleClick = currentString
End Function
