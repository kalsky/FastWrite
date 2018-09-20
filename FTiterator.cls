VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTiterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public currentIndex As Long
Public deepnessLevel As Integer
Public lastMatchedWord As Dictionary
Public levelString As String
Private currentSql As String
Private currentSet As Recordset

Public Sub newword()
Dim i As Long
Dim recCount As Long

  levelString = vbNullString
  lastMatchedWord.RemoveAll
  deepnessLevel = 0
  currentIndex = 0
  recCount = DB.Recordsets.Count - 1
  For i = 0 To recCount
    DB.Recordsets(i).Close
  Next 'i
  Set currentSet = Nothing
  
'  If Not (currentSet Is Nothing) Then
'    currentSet.Close
'  End If
  
End Sub

Public Sub StartOver()
  currentIndex = 0
  lastMatchedWord.RemoveAll
End Sub

Public Sub addPriority(curStr As String)
Dim rs As Recordset
  
  With currentSet
    If currentIndex > 0 Then
      If .RecordCount >= (currentIndex - 1) Then
        .AbsolutePosition = currentIndex - 1
        If !word = curStr Then
          Set rs = DB.OpenRecordset(Replace(currentSql, "word,key", "word,key,priority"))
          With rs
            .AbsolutePosition = currentIndex - 1
            .Edit
            .Fields(2).Value = rs.Fields(2).Value + 1
            'Debug.Print "priority for the word: '" & .Fields(0).Value & "' is now " & .Fields(2).Value
            .Update
            .Close
          End With
        End If
      End If
    End If
  End With
End Sub


Public Sub MoveTo(w As String)
Dim s As String
Dim recCount As Long

   currentIndex = 0
   If currentSet Is Nothing Then Exit Sub
   lastMatchedWord.RemoveAll
   
   With currentSet
     recCount = .RecordCount
     Do While currentIndex < recCount
        .AbsolutePosition = currentIndex
        s = !word 'recset.Fields(0).Value 'frmMain.DBGrid1.Columns(0).Value
  
        currentIndex = currentIndex + 1
        If Trim$(w) = Trim$(s) Then
           lastMatchedWord.Add (s), 1
           Exit Sub
        End If
    Loop
  End With
    
End Sub


'        /**
'         * Returns the next element in the iteration.
'         *
'         * @return next element in the iteration.
'         *
'         * @exception NoSuchElementException iteration has no more elements.
'         */
Public Function Current(Optional exactSize As Boolean = False) As String
    
Dim tempCurrentIndex As Long
Dim s As String
Dim retVal As String
Dim recCount As Long
      
      retVal = vbNullString
      tempCurrentIndex = currentIndex
      
      With currentSet
        recCount = .RecordCount
              
        If exactSize = True Then
          
          Do While currentIndex < recCount
            .AbsolutePosition = currentIndex
            s = !word 'recSet.Fields(0).Value
            
            currentIndex = currentIndex + 1
            If deepnessLevel = Len(s) Then
              If Not lastMatchedWord.Exists(s) Then
                lastMatchedWord.Add (s), currentIndex - 1
                retVal = s
                Exit Do
              End If
            End If
          Loop
        End If
  
        If (LenB(retVal) = 0) Then
        
          'Dim ss As String
          currentIndex = tempCurrentIndex
          Do While currentIndex < recCount
              .AbsolutePosition = currentIndex
              s = !word 'recSet.Fields(0).Value 'frmMain.DBGrid1.Columns(0).Value
  
              currentIndex = currentIndex + 1
              If deepnessLevel <= Len(s) Then
  
                  If Not lastMatchedWord.Exists(s) Then
                      lastMatchedWord.Add (s), currentIndex - 1
                      retVal = s
                      Exit Do
                  End If
              End If
          Loop
        End If
      End With
exitfunc:
Current = retVal


End Function

'        /**
'         * Returns the element that was in the last iteration.
'         *
'         *
'         *
'         *
'         */
Public Function Last() As String
    
Dim tempCurrentIndex As Long
Dim key
Dim beforelast As String
Dim lastKey As String

      tempCurrentIndex = 0
      For Each key In lastMatchedWord.keys
        tempCurrentIndex = tempCurrentIndex + 1
        If tempCurrentIndex = lastMatchedWord.Count - 1 Then
          beforelast = key
        ElseIf tempCurrentIndex = lastMatchedWord.Count Then
          lastKey = key
        End If
      Next 'key
      currentIndex = lastMatchedWord.Item(beforelast) + 1
      lastMatchedWord.Remove lastKey
      Last = beforelast
        
    
End Function

'         /**
'         * Returns true if the iteration has more elements. (In other words,
'         * returns <code>true</code> if <code>next</code> would return an
'         * element rather than throwing an exception.)
'         *
'         * @return true if the iterator has more elements.
'         */

Public Function MoveNext(Optional wholeWord As Boolean = False) As Boolean
 
Dim tempCurrentIndex As Integer
Dim s As String
Dim ss As String
Dim recCount As Long
       
    MoveNext = False
    tempCurrentIndex = currentIndex
    
    With currentSet
    
      recCount = .RecordCount
      
      While currentIndex < recCount
          
          .AbsolutePosition = currentIndex
          s = !word 'recSet.Fields(0).Value 'frmMain.DBGrid1.Columns(0).Value
          currentIndex = currentIndex + 1
          If deepnessLevel = Len(s) Then
              If Not lastMatchedWord.Exists(s) Then
                  currentIndex = tempCurrentIndex
                  MoveNext = True
                  Exit Function
              End If
          End If
      Wend
      If wholeWord = True Then Exit Function
      
      currentIndex = tempCurrentIndex
      While currentIndex < .RecordCount
          
          .AbsolutePosition = currentIndex
          s = !word 'recSet.Fields(0).Value 'frmMain.DBGrid1.Columns(0).Value
          currentIndex = currentIndex + 1
          If deepnessLevel <= Len(s) Then
              
              ss = Mid$(s, 1, deepnessLevel)
              If Not lastMatchedWord.Exists(ss) Then
                  currentIndex = tempCurrentIndex
                  MoveNext = True
                  Exit Function
              End If
          End If
      Wend
    End With
    MoveNext = False
End Function




'         /**
'         * Returns true if the iteration has more elements. (In other words,
'         * returns <code>true</code> if <code>last</code> would return an
'         * element rather than throwing an exception.)
'         *
'         * @return true if the iterator has more elements.
'         */

Public Function MoveBack() As Boolean

   If lastMatchedWord.Count > 1 Then
     MoveBack = True
   Else
     MoveBack = False
   End If
    
End Function


'        /**
'         * Go to next level in dictionary. This means that user adds a
'         * key to word that is searched. Iteration continues from start
'         * of next level.
'         *
'         * @param key a button pressed from UI.
'         */
Public Sub nextLevel(ByVal key As Integer)

Dim sqlS As String

    lastMatchedWord.RemoveAll
    
    If key <> 7 And Right$(levelString, 1) = "7" Then
      Call newword
    End If
    deepnessLevel = deepnessLevel + 1
    
    levelString = levelString & key
    sqlS = "select word,key from words where key like '" & levelString & "*'"
    If deepnessLevel < 3 Then
      sqlS = sqlS & " and len(key)<=" & Trim$(Str$(Len(levelString) + 2))
    End If
    sqlS = sqlS & " order by priority DESC"
    currentSql = sqlS
    If Not (currentSet Is Nothing) Then
      currentSet.Close
      Set currentSet = Nothing
    End If
    Set currentSet = DB.OpenRecordset(sqlS)
    'Debug.Print "chacesize: " & currentSet.CacheSize
    currentIndex = 0
    
End Sub

'         /**
'         * Go to previous level in dictionary. Iteration continues from start
'         * of previous level.
'         */
Public Sub prevLevel()
Dim sqlS As String

     deepnessLevel = deepnessLevel - 1
     If deepnessLevel <= 0 Then
        currentIndex = 0
        deepnessLevel = 0
        levelString = vbNullString
        Exit Sub
     End If
     levelString = Left$(levelString, Len(levelString) - 1)
     sqlS = "select word,key from words where key like '" & levelString & "*'"
     If deepnessLevel < 3 Then
       sqlS = sqlS & " and len(key)<=" & Trim$(Str$(Len(levelString) + 2))
     End If
     sqlS = sqlS & " order by priority DESC"
     currentSql = sqlS
     If Not (currentSet Is Nothing) Then
      currentSet.Close
      Set currentSet = Nothing
     End If
     Set currentSet = DB.OpenRecordset(sqlS)
     currentIndex = 0
     lastMatchedWord.RemoveAll
     
End Sub

        
Private Sub Class_Initialize()
    Set lastMatchedWord = New Dictionary
    deepnessLevel = 0
    currentIndex = 0
End Sub

Private Sub Class_Terminate()
On Error Resume Next
    Set lastMatchedWord = Nothing
End Sub
