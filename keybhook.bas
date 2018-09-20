Attribute VB_Name = "keybhook"
Option Explicit

'This code can be used for a keyboard or mouse hook.
'Any questions or comments, email me at itcdr@yahoo.com

'Type for Keyboard Hook
Private Type KBDLLHOOKSTRUCT
 code As Long
End Type
'Type for Mouse
'Private Type POINTAPI
' X As Long
' Y As Long
'End Type


'Win 32 API Functions found in API Viewer - see APIViewer in my folder for upgraded API Viewer
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function VkKeyScanW Lib "user32" (ByVal cChar As Integer) As Integer
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'keyboard
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetLocaleInfo& Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale&, _
 ByVal LCType&, ByVal lpLCData$, ByVal cchData&) ' ANSI version
Private Const LOCALE_SLANGUAGE& = &H2 '  localized name of language

Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&) ' not NT?
Private Const DWL_ANYTHREAD& = 0 ' or pass the thread you want to test for
Private KBLayout& ' track current
Private Const bufsize = 256 ' for API string handling

'num lock
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'fw window
Public Declare Function GetForegroundWindow Lib "user32" () As Long
'window
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'cursor
'Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
'caret
'Private Declare Function GetCaretPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
'Private Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
'Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

''accesibility
'Private Type UUID
'   Data1 As Long
'   Data2 As Integer
'   Data3 As Integer
'   Data4(0 To 7) As Byte
'End Type
'
'Declare Function AccessibleObjectFromPoint Lib "oleacc" (ByVal lx As Long, ByVal ly As Long, ppacc As IAccessible, pvarChild As Variant) As Long
''Active Accessibility declarations
'Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As UUID, ppvObject As Object) As Long
'Private Declare Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As IAccessible, ByVal iChildStart As Long, ByVal cChildren As Long, rgvarChildren As Variant, pcObtained As Long) As Long
'Private Declare Function GetRoleText Lib "oleacc" Alias "GetRoleTextA" (ByVal dwRole As Long, ByVal szRole As String, ByVal cchRoleMax As Integer) As Long
'Private Declare Function GetStateText Lib "oleacc" Alias "GetStateTextA" (ByVal dwStateBit As Long, ByVal szState As String, ByVal cchStateBitMax As Integer) As Long


' Constant declarations:
Private Const VK_NUMLOCK = &H90
Private Const KEYEVENTF_KEYUP = &H2
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const EXTENDEDKEY = &H1

'Constants found in API Viewer under constants
'Check if action occured. Only possible parameter for ncode in hookProc
Private Const HC_ACTION As Long = 0
Private Const WH_KEYBOARD_LL As Long = 13

'HookProc wParam Parameters:

'Keyboard event constants
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101

Private keyboard As Long  'Variable for set hook
'Keyboard Hook Variable
Dim hookKey As KBDLLHOOKSTRUCT
'Variable to turn on and off mechanizem

'ascii consts
Public Const AscWa As Integer = 97
Public Const AscWz As Integer = 122
Public Const AscW0 As Integer = 48
Public Const AscWCapA As Integer = 65
Public Const AscWCapZ As Integer = 90


Private arr(0 To 14) As Integer
Private lastActiveWindow As Long
Private inKeyb As Boolean
'Private inDisableKeyb As Boolean
'for use without dictionary
Private RotLetters() As String
Private FirstLetters() As String
Private NumLetters() As String
Private Const FIRSTLTTRS As String = "p,t,w,g,j,m,.,a,d"
Private Const ROTATELTTRS As String = "b,c,a,e,f,d,h,i,g,k,l,j,n,o,m,q,r,s,p,u,v,t,x,y,z,w"
Public Const PUNCT As String = ".,?;:!@#$%^&*()-+~'`/\|<>{}[]_="""
Private Const SHIFTEDPUNCT As String = "?:!@#$%^&*()+~|<>{}_+"""
Private Const NOTSHIFTED As String = "abcdefghijklmnopqrstuvwxyz,./;'[]\=-`1234567890"
Private Enum TypeStateEnum
  TypeState_INIT = 0
  TypeState_WAIT = 1
  TypeState_DONE = 2
End Enum
Public Enum TypeModesEnum
  TypeModes_MULTITAP = 0
  TypeModes_SINGLETAP = 1
  TypeModes_SINGLETAPPLUS = 2
  TypeModes_NUMBERS = 3
  TypeModes_OFF = 4 'off
End Enum

'modes
Private TypeState As TypeStateEnum
Private TypeMode As TypeModesEnum
Private LastMode As TypeModesEnum

Public Function GetTypeMode() As TypeModesEnum
  GetTypeMode = TypeMode
End Function

Private Sub KeyDown(ByVal key As Integer, Optional extended As Integer = 0)
On Error Resume Next
  keybd_event key, 0, extended, 0  'key down
End Sub

Private Sub KeyUp(ByVal key As Integer, Optional extended As Integer = 0)
On Error Resume Next
  keybd_event key, 0, extended Or &H2, 0 'key up
End Sub

Private Sub KeyPress(ByVal key As Integer, Optional extended As Integer = 0)
On Error GoTo errH
  keybd_event key, 0, extended, 0 'key down
  keybd_event key, 0, extended Or &H2, 0 'key up
errH:
End Sub

Private Sub KeyPressShifted(ByVal key As Integer, Optional extended As Integer = 0)
On Error GoTo errH
  KeyDown vbKeyShift
  KeyPress key, extended
  KeyUp vbKeyShift
errH:
End Sub

Private Sub KeyPressCtrlC()
On Error GoTo errH
  KeyDown vbKeyControl
  KeyPress vbKeyC
  KeyUp vbKeyControl
errH:
End Sub

Private Sub KeyPressCtrlV()
On Error GoTo errH
  KeyDown vbKeyControl
  KeyPress vbKeyV
  KeyUp vbKeyControl
errH:
End Sub

Private Sub KeyPressCtrlLeft()
On Error GoTo errH
  KeyDown vbKeyControl
  KeyPress vbKeyLeft, EXTENDEDKEY
  KeyUp vbKeyControl
errH:
End Sub

Private Sub KeyPressCtrlRight()
On Error GoTo errH
  KeyDown vbKeyControl
  KeyPress vbKeyRight, EXTENDEDKEY
  KeyUp vbKeyControl
errH:
End Sub

Private Sub KeyPressDeleteWord(wLen As Integer)
On Error Resume Next
Dim i As Integer

  For i = 1 To wLen
    KeyPress vbKeyBack
  Next 'i
  
End Sub


Private Function IsSameWindow() As Boolean
Dim foregound_hwnd As Long

  foregound_hwnd = GetForegroundWindow()

  If lastActiveWindow <> foregound_hwnd Then
    IsSameWindow = True
    lastActiveWindow = foregound_hwnd
  Else
    IsSameWindow = False
  End If
End Function

Private Sub SingleTap()

Dim s As String
Dim lenS As Integer
Dim currentKey As Integer
Dim i As Integer
Dim ll As Integer
Dim code As Integer
Dim newword As Boolean
Dim curLetter As String
Dim tempS As String
Dim wasEnter As Boolean
Dim theword As String
     
On Error GoTo errH

    'currentKey = hookKey.code 'use this one to igonre run-overs
    currentKey = keyBuffer.ReadKey  'use this one to igonre run-overs
    wasEnter = False
    
    newword = IsSameWindow()

    Select Case currentKey
    Case vbKeyDivide '/
      code = arr(11)
    Case vbKeyDecimal  '. add new word
      If IsOfficeWindow(GetForegroundWindow) = False Then
        s = Trim$(GetPrevWord(0))
        Call dicMdl.AddNewWord(s)
      End If
    Case vbKeySubtract '-
      code = arr(12)
    Case vbKeyAdd '+
      If IsOfficeWindow(GetForegroundWindow) = False Then
        s = LCase$(Trim$(GetPrevWord(1)))
        Dim ss As String
        If s >= "a" And s <= "z" Then
          'debug.print s
          
          SelLastWord
          tempS = ClipGetText
          ClipSetText " "
          KeyPressCtrlC
          s = ClipGetText
          ClipSetText tempS
          ss = Left$(s, 1)
          If AscW(ss) >= AscWa And AscW(ss) <= AscWz Then
            s = UCase$(Left$(s, 1)) & Mid$(s, 2)
          ElseIf AscW(ss) >= AscWCapA And AscW(ss) <= AscWCapZ Then
            If UCase$(s) <> s Then
              s = UCase$(s)
            Else
              s = LCase$(s)
            End If
          End If
          lenS = Len(s)
          If lenS > 0 Then
            For i = 1 To lenS
              ss = Mid$(s, i, 1)
              If Trim$(ss) = UCase$(Trim$(ss)) Then
                KeyDown vbKeyShift
              End If
              KeyPress VkKeyScanW(AscW(LCase$(ss)))
              If Trim$(ss) = UCase$(Trim$(ss)) Then
                KeyUp vbKeyShift
              End If
            Next 'i
          End If
          
        End If
      Else 'office window
        'if mainHandler.getLastString
      End If
      
    Case Else
      If currentKey = vbKeySeparator Or currentKey = vbKeyReturn Then
        wasEnter = True
        currentKey = vbKeyNumpad0
      End If
      code = arr(currentKey - vbKeyNumpad0)
    End Select
    
    If currentKey = vbKeyAdd Or currentKey = vbKeyDecimal Then '+ .
      Exit Sub 'keypress was handled - exit sub
    End If
    
'    If newword = False And lastKey = vbKeySubtract And currentKey <> vbKeySubtract Then
'
'    End If

    s = mainHandler.handleClick(code, newword)
    theword = mainHandler.getLastString
    lenS = Len(s)
           
AfterHandleClick:
    If lenS > 0 Then
       If currentKey = vbKeyMultiply Or currentKey = vbKeyDivide Then '* /
         ll = mainHandler.getDeepLevel
       Else
         ll = mainHandler.getDeepLevel - 1
       End If
              
       If currentKey = vbKeyNumpad7 Then
         KeyPress vbKeyRight
         KeyPress vbKeySpace
         KeyPress vbKeyBack
       Else
         KeyPress vbKeySpace
         KeyPressDeleteWord ll + 1
       End If
                  
       
        
       'debug.print s
       For i = 1 To lenS
         'in use when word completion is enabled
         If Mid$(s, i, 1) <> vbTab Then
           curLetter = Mid$(s, i, 1)
           If (InStr(NOTSHIFTED, curLetter)) Then
  
             KeyPress VkKeyScanW(AscW(curLetter))
             
           ElseIf (curLetter >= "A" And curLetter <= "Z") Then
             
             KeyPressShifted VkKeyScanW(AscW(LCase$(curLetter)))
             
           Else
             
             KeyPressShifted VkKeyScanW(AscW(GetPreShiftedKey(curLetter)))
             
           End If
         End If
         
       Next 'i
       
       If InStr(s, vbTab) Then
         KeyDown vbKeyShift
         lenS = Len(s) - InStrRev(s, vbTab)
         For i = 1 To lenS
           KeyPress vbKeyLeft, EXTENDEDKEY
         Next 'i
         KeyUp vbKeyShift
       End If
     
     
    Else 'lenS=0
       If currentKey = vbKeyNumpad0 Then  '_ space
         If db_autoAddNew = 1 Then
           Call dicMdl.AddNewWord(theword)
           'Debug.Print theword
         End If
         If IsOfficeWindow(GetForegroundWindow) = False Then
           tempS = ClipGetText
           ClipSetText " "
           KeyPressCtrlC
           s = ClipGetText
  
           If Len(Trim$(s)) > 0 Then
             KeyPressCtrlV
           End If
  
           ClipSetText tempS
           
           If wasEnter = True Then
             KeyPress vbKeyReturn
           Else
             KeyPress vbKeySpace
           End If
         Else
           tempS = mainHandler.getLastString
           If InStr(tempS, vbTab) Then
             tempS = Mid$(tempS, InStr(tempS, vbTab) + Len(vbTab))
             KeyPress vbKeyRight, EXTENDEDKEY
             
           End If
           If wasEnter = True Then
             KeyPress vbKeyReturn
           Else
             KeyPress vbKeySpace
           End If
         End If
                   
       ElseIf currentKey = vbKeySubtract Then '- backspace
                            
         KeyPress vbKeyBack
         
         If IsOfficeWindow(GetForegroundWindow) = False And (db_fastDelete = 0) Then
          
           s = LCase$(GetPrevWord(1))
           
           s = Trim$(s)
           If LenB(s) = 0 Or s < "a" Or s > "z" Then
             Exit Sub
           End If
           
           SelLastWord
           
           tempS = ClipGetText
           'Debug.Print tempS
           ClipSetText " "
           KeyPressCtrlC
           s = ClipGetText
  
           If Len(Trim$(s)) > 0 Then
             KeyPressCtrlV
           End If
           ClipSetText tempS
           
           Call mainHandler.SetPrevWord(LCase$(Trim$(s)))
         Else
           'no extra treatment right now...
           If Len(mainHandler.getLastString) > 0 Then
             Call mainHandler.SetPrevWord(mainHandler.getLastString)
           End If
         End If
       End If
    End If 'len(s)
errH:
    
End Sub



Private Sub MultiTap()
Dim newkey As String
Dim s As String
Dim currentKey As Integer

     On Error Resume Next
     'currentKey = hookKey.code 'use this one to igonre run-overs
     currentKey = keyBuffer.ReadKey  'use this one to igonre run-overs
     
     If (TypeState = TypeState_DONE) Or (TypeState = TypeState_INIT) Then
       s = vbNullString
     Else
       KeyPress vbKeyRight, EXTENDEDKEY
       s = GetPrevWord(1)
     End If
     
     If LenB(Trim$(s)) = 0 Then TypeState = TypeState_INIT
     
     If TypeState = TypeState_WAIT Then
       If IsPunctuation(s) Then
         'Debug.Print currentKey, s
         If currentKey = vbKeyNumpad7 Then
           KeyPress vbKeyBack
         Else
           s = vbNullString
         End If
       Else
         If NumLetters(AscW(LCase$(s)) - AscWa) = (currentKey - AscWa + 1) Then
           KeyPress vbKeyBack
         Else
           s = vbNullString
         End If
       End If
     End If
       
     If LenB(Trim$(s)) > 0 Then
     
       s = Trim$(s)
       If s >= "a" And s <= "z" Then
         newkey = RotLetters(AscW(s) - AscWa)
         KeyPress VkKeyScanW(AscW(newkey))
       ElseIf s >= "A" And s <= "Z" Then
         newkey = RotLetters(AscW(LCase$(s)) - AscWa)
         KeyPressShifted VkKeyScanW(AscW(newkey))
       ElseIf IsPunctuation(s) Then
         newkey = Mid$(PUNCT, InStr(PUNCT, s) + 1, 1)
         If LenB(newkey) = 0 Then newkey = Left$(PUNCT, 1)
         If IsShifted(newkey) Then
           newkey = GetPreShiftedKey(newkey)
           KeyPressShifted VkKeyScanW(AscW(newkey))
         Else
           KeyPress VkKeyScanW(AscW(newkey))
         End If
       Else 'numbers - i guess
       
       End If
       
       
       
     ElseIf currentKey = vbKeyNumpad0 Then  '_ space
       If db_autoAddNew = 1 And IsOfficeWindow(GetForegroundWindow) = False Then
          s = Trim$(GetPrevWord(0))
          Call dicMdl.AddNewWord(s)
          
          TypeState = TypeState_DONE
       End If
       KeyPress vbKeySpace
       GoTo waitstate
     ElseIf currentKey = vbKeyAdd Then '+ capital letter
       If IsOfficeWindow(GetForegroundWindow) = False Then
          s = GetPrevWord(1)
          'debug.print s
          If s >= "a" And s <= "z" Then
            KeyPress vbKeyBack
            KeyPressShifted VkKeyScanW(AscW(s))
          ElseIf s >= "A" And s <= "Z" Then
            KeyPress vbKeyBack
            KeyPress VkKeyScanW(AscW(LCase$(s)))
          End If
          
          GoTo waitstate
       Else
          Exit Sub
       End If
     ElseIf currentKey = vbKeySubtract Then '- backspace
       KeyPress vbKeyBack
       
       TypeState = TypeState_DONE
       Exit Sub
     ElseIf currentKey = vbKeyDecimal Then '. add word
       If IsOfficeWindow(GetForegroundWindow) = False Then
          s = Trim$(GetPrevWord(0))
          Call dicMdl.AddNewWord(s)
          
          TypeState = TypeState_DONE
       End If
       Exit Sub
     
     Else 'letters
       If currentKey >= vbKeyNumpad1 And currentKey <= vbKeyNumpad9 Then
         newkey = FirstLetters(currentKey - vbKeyNumpad1)
         KeyPress VkKeyScanW(AscW(newkey))
       End If
       If currentKey = vbKeyReturn Then
         KeyPress vbKeyReturn
         TypeState = TypeState_DONE
         Exit Sub
       End If
     End If
     
     KeyPressShifted vbKeyLeft, EXTENDEDKEY
          
waitstate:
     TypeState = TypeState_WAIT
     frmMain.typeTmr.Enabled = True
      
End Sub

Public Sub DisableIfNotEnglish()
'Exit Sub
    'If TypeMode = TypeModes_OFF Then Exit Sub
    
    ' Display KeyBoard Locale
    ' (this is great code to put on a status bar!)
    Dim r&
    
    r = GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow(), 0)) '0&
    If KBLayout = 0 Then
      KBLayout = r
      Exit Sub
    End If
    If r <> KBLayout Then
    Debug.Print "layout changed " & Time
        KBLayout = r
        'Debug.Print r
        r = val("&H" & Right$(Hex$(r), 4)) ' lower 16bits
        'Debug.Print r
        Dim strng$, Buffer As String * bufsize
        r = GetLocaleInfo(r, LOCALE_SLANGUAGE, Buffer, _
            bufsize - 1)
        strng = Buffer
        If InStr(strng, "English") = 0 Then 'not english
          If TypeMode <> TypeModes_OFF Then
            LastMode = TypeMode
            Debug.Print "in change mode to " & strng
            
            SwitchMode TypeModes_OFF
            
          End If
          frmMain.ShowInfo "Not in English layout", True
          
        Else 'back to english
          If LastMode <> TypeModes_OFF Then
            Debug.Print "in change mode to english"
            If NumLockOn = False Then
              ToggleNumLock 1
            End If
            SwitchMode LastMode
          End If
          
          frmMain.ShowInfo "", True
        End If
        

    End If


    
End Sub

Public Function CanChangeMode() As Boolean
  Dim r&
    
    r = GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow(), 0)) '0&

    r = val("&H" & Right$(Hex$(r), 4)) ' lower 16bits
    Dim strng$, Buffer As String * bufsize
    r = GetLocaleInfo(r, LOCALE_SLANGUAGE, Buffer, _
        bufsize - 1)
    strng = Buffer
    If InStr(strng, "English") = 0 Then
      CanChangeMode = False
    Else
      CanChangeMode = True
    End If
     
End Function

Public Sub HandleKeyPress()
Dim tempkey As Integer
Dim t As Single
Dim t1 As Single

  inKeyb = True
  tempkey = keyBuffer.CopyKey
  If tempkey <> -1 Then
    Call frmMain.ColorKey(tempkey, True)
    t = Timer
  End If
  
  If TypeMode = TypeModes_SINGLETAP Or TypeMode = TypeModes_SINGLETAPPLUS Then
  
    Call SingleTap
    
  ElseIf TypeMode = TypeModes_MULTITAP Then
  
    Call MultiTap
     
  End If 'typemode
  
  If tempkey <> -1 Then
    
    t1 = (Timer - t) * 1000
    
    If 80 - t1 > 0 Then
      'debug.print 80 - t1
      Sleep 80 - t1
    End If
    Call frmMain.ColorKey(tempkey, False)
  End If
  inKeyb = False
  DoEvents
  If keyBuffer.IsEmpty = False Then
    frmMain.IntTimer.Enabled = True
  End If
errormessage:
  If LenB(frmMain.errorMsg) > 0 Then
    frmMain.msgTmr.Enabled = True
  End If
End Sub

Public Sub SwitchMode(Optional newtype As TypeModesEnum = 99)
  
  If newtype <> TypeModes_OFF And CanChangeMode = False Then Exit Sub
  
  If newtype = 99 Then
    TypeMode = (TypeMode + 1) Mod (TypeModes_OFF + 1)
  Else
    TypeMode = newtype
  End If
  mainHandler.StartNewWord
  
  frmMain.mnuMultiTap.Checked = False
  frmMain.mnuNumbers.Checked = False
'  frmMain.mnuOff.Checked = False
  frmMain.mnuSingleTap.Checked = False
  frmMain.mnuSingleTapPlus.Checked = False
  
  Select Case TypeMode
    Case TypeModes_MULTITAP
      frmMain.UpdateCaption "MultiTap"
      frmMain.mnuMultiTap.Checked = True
    Case TypeModes_SINGLETAP
      frmMain.UpdateCaption "SingleTap"
      frmMain.mnuSingleTap.Checked = True
    Case TypeModes_SINGLETAPPLUS
      frmMain.UpdateCaption "SingleTap+"
      frmMain.mnuSingleTapPlus.Checked = True
    Case TypeModes_NUMBERS
      frmMain.UpdateCaption "Numbers"
      frmMain.mnuNumbers.Checked = True
    Case TypeModes_OFF
      frmMain.UpdateCaption "Off"
'      frmMain.mnuOff.Checked = True
  End Select
  frmMain.SetTrayIcon TypeMode
  
End Sub

Private Sub SetNumLockState(State As Integer)
Dim keys(0 To 255) As Byte
  
  GetKeyboardState keys(0)
  keys(VK_NUMLOCK) = State
  SetKeyboardState keys(0)
  
End Sub

Private Function NumLockOn() As Boolean
Dim iKeyState As Integer
    
    iKeyState = GetKeyState(vbKeyNumlock)
    NumLockOn = (iKeyState = 1 Or iKeyState = -127)
  
End Function

Private Sub ToggleNumLock(mode As Integer)
'Dim keys(0 To 255) As Byte
'Dim keys1(0 To 255) As Byte
  
  'GetKeyboardState keys(0)
  
  
Debug.Print "numlock state is " & NumLockOn & ", new state should be " & mode
  If NumLockOn <> CBool(mode) Then
    'SetNumLockState mode
    KeyPress vbKeyNumlock, EXTENDEDKEY
'    KeyUp vbKeyNumlock, EXTENDEDKEY
'    KeyDown vbKeyNumlock, EXTENDEDKEY
'    SetNumLockState mode
    
'    keys(VK_NUMLOCK) = mode
'    SetKeyboardState keys(0)
    Debug.Print "num pressed"
    
  End If
  
Debug.Print "new num lock state is " & NumLockOn
End Sub

Public Function KeyboardProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo errH

 'Check if action occured
 
 If ncode = HC_ACTION Then
    
    If wParam = WM_KEYUP Then
      'remove any extra key presses
      If keyBuffer.GetUnfreeSize > 1 Then
        'Debug.Print keyBuffer.GetUnfreeSize
        keyBuffer.CleanBuffer
      End If
    End If
    
    If wParam = WM_KEYDOWN Then
    
      'If inKeyb = True Then GoTo nexthook
      'If inDisableKeyb = True Then GoTo nexthook
      
      frmMain.typeTmr.Enabled = False
      frmMain.IntTimer.Enabled = False
      
      'Copy Memory address of action to hook variable
      Call CopyMemory(hookKey, ByVal lParam, Len(hookKey))
      
      'Block all keyboard input
      KeyboardProc = 1
            
      'If F1 is pressed show help
      If hookKey.code = vbKeyF1 Then
        If GetForegroundWindow = frmMain.hwnd Then
          frmMain.helpTimer.Enabled = True
          Exit Function
        Else
          GoTo nexthook
        End If
      End If

      'If vbKeyNumlock is pressed toggle modes
      If hookKey.code = vbKeyNumlock Then
        If inKeyb Then Exit Function
        Call SwitchMode
        
        If TypeMode = TypeModes_OFF Then
          GoTo nexthook
        Else
          If TypeMode = TypeModes_MULTITAP Then
'            If NumLockOn = False Then
'            Debug.Print "repress numlock"
'              inKeyb = True
'              KeyUp vbKeyNumlock, EXTENDEDKEY
'              KeyDown vbKeyNumlock, EXTENDEDKEY
'              inKeyb = False
'              SetNumLockState 1
'            End If
            GoTo nexthook
          Else
            inKeyb = True
            'KeyPress vbKeyNumlock, EXTENDEDKEY
            'ToggleNumLock 1
            KeyUp vbKeyNumlock, EXTENDEDKEY
            KeyDown vbKeyNumlock, EXTENDEDKEY
            SetNumLockState 1
    
            inKeyb = False
          End If
        End If
        Exit Function
      End If
      
      'If F12 is pressed Exit program
      If hookKey.code = vbKeyF12 Then
        Unhook
        frmMain.realyExit = True
        Unload frmMain
        Exit Function
      End If
      

        
      If TypeMode = TypeModes_NUMBERS Then GoTo nexthook
      If TypeMode = TypeModes_OFF Then GoTo nexthook
      

      
      If ((hookKey.code >= vbKeyNumpad0 And hookKey.code <= vbKeyDivide) And (hookKey.code <> vbKeySeparator)) Or (hookKey.code = vbKeyReturn) Then
        If inKeyb = True And hookKey.code = vbKeyReturn Then GoTo nexthook
        
        If keyBuffer.AddKey(hookKey.code) = True Then
          'If keyBuffer.HasSome = False Then
            'frmMain.IntTimer.Enabled = True
          'End If
        End If
        frmMain.IntTimer.Enabled = True
        Exit Function
        
      End If 'hookey.code
      
      If (hookKey.code = vbKeySpace And inKeyb = False) Then
        If keyBuffer.AddKey(vbKeyNumpad0) = True Then
          'If keyBuffer.HasSome = False Then
            'frmMain.IntTimer.Enabled = True
          'End If
        End If
        frmMain.IntTimer.Enabled = True
        Exit Function
      End If
      
      If (hookKey.code = vbKeyBack And inKeyb = False) Then
        If keyBuffer.AddKey(vbKeySubtract) = True Then
          'If keyBuffer.HasSome = False Then
            'frmMain.IntTimer.Enabled = True
          'End If
        End If
        frmMain.IntTimer.Enabled = True
        Exit Function
      End If

      
   End If
 End If
  
 GoTo nexthook
 
errH:
 'debug.print Err.Description
 Err.Clear

nexthook:
 'If the message is not one we want to trap, pass it along
 'Debug.Print "next hook key"
 KeyboardProc = CallNextHookEx(keyboard, ncode, wParam, lParam)
 Exit Function
End Function

Private Function IsOfficeWindow(hwnd As Long) As Boolean
Dim r As Long
Dim clsName As String
Const officeClasses As String = "OMain,XLMAIN,FrontPageExplorerWindow40,rctrl_renwnd32,PP7FrameClass,PP97FrameClass,PP9FrameClass,PP10FrameClass,JWinproj-WhimperMainClass,OpusApp"
    
    clsName = Space$(256)
    r = GetClassName(hwnd, clsName, 255)

    'debug.print clsName
    If InStr(officeClasses, Left$(clsName, r)) Then
        IsOfficeWindow = True
    Else
        IsOfficeWindow = False
    End If
  
End Function

Private Function IsShifted(l As String) As Boolean
  If InStr(SHIFTEDPUNCT, l) Then
    IsShifted = True
  Else
    IsShifted = False
  End If
End Function

Private Function IsPunctuation(l As String) As Boolean
  If InStr(PUNCT, l) Then
    IsPunctuation = True
  Else
    IsPunctuation = False
  End If
End Function

Private Function GetPreShiftedKey(l As String) As String
  Dim kp As String
  Select Case l
    Case ":"
      kp = ";"
    Case "?"
      kp = "/"
    Case "<"
      kp = ","
    Case ">"
      kp = "."
    Case "~"
      kp = "`"
    Case "!"
      kp = "1"
    Case "@"
      kp = "2"
    Case "#"
      kp = "3"
    Case "$"
      kp = "4"
    Case "%"
      kp = "5"
    Case "^"
      kp = "6"
    Case "&"
      kp = "7"
    Case "*"
      kp = "8"
    Case "("
      kp = "9"
    Case ")"
      kp = "0"
    Case "_"
      kp = "-"
    Case "+"
      kp = "="
    Case ChrW$(34)  ' "
      kp = "'"
    Case "{"
      kp = "["
    Case "}"
      kp = "]"
    Case "|"
      kp = "\"
    Case Else
      kp = "?"
  End Select
  GetPreShiftedKey = kp
End Function


Public Sub TypeTmrHandler()
  TypeState = TypeState_DONE
  KeyPress vbKeyRight, EXTENDEDKEY
End Sub

Private Function ClipGetText() As String
Dim ret As String

  On Error Resume Next
  Do
    Err.Clear
    ret = Clipboard.GetText
    DoEvents
  Loop Until Err.Number = 0
  Err.Clear
  On Error GoTo 0
  ClipGetText = ret
End Function

Private Sub ClipSetText(s As String)

  On Error Resume Next
  Do
    Err.Clear
    Clipboard.SetText s
    DoEvents
  Loop Until Err.Number = 0
  Err.Clear
  On Error GoTo 0
  
End Sub

Public Function GetPrevWord(size As Integer) As String
Dim tempS As String
Dim i As Integer
Dim lenS As Integer
Dim s As String

  If size = 0 Then 'word
  
'    For i = 1 To 3
'      'KeyPress vbKeyRight, EXTENDEDKEY
'      KeyDown vbKeyLeft ', 0, EXTENDEDKEY, 0
'      Sleep 100
'    Next i
'    For i = 1 To 3
'      'KeyPress vbKeyRight, EXTENDEDKEY
'      KeyDown vbKeyRight ', 0, EXTENDEDKEY, 0
'      Sleep 100
'    Next i
  
  
    KeyDown vbKeyControl
  End If
  
  KeyPressShifted vbKeyLeft, EXTENDEDKEY
  KeyUp vbKeyControl
  
  'backup clipboard
  tempS = ClipGetText
  ClipSetText " "
  
  'press control+c to copy selected
  'If size <> 0 Then 'not word
    KeyDown vbKeyControl
    KeyPress vbKeyC
    KeyUp vbKeyControl
  'Else
  '  KeyPress vbKeyC
  'End If
  
  'move back to the start position
  If size = 0 Then
    
    'KeyUp vbKeyControl
'    KeyPress vbKeyControl
'    KeyUp vbKeyControl
'    DoEvents
    lenS = Len(Trim$(ClipGetText))
    For i = 1 To lenS
      'KeyPress vbKeyRight, EXTENDEDKEY
      KeyDown vbKeyRight ', 0, EXTENDEDKEY, 0
      'Sleep 100
    Next 'i
    'Stop
    'KeyUp vbKeyRight, EXTENDEDKEY
  Else
    KeyPress vbKeyRight, EXTENDEDKEY
  End If
     
  DoEvents
  
  s = ClipGetText
  If s = " " Then
    GetPrevWord = vbNullString
  Else
    GetPrevWord = s
  End If
    
    
  ClipSetText tempS
  
  'Debug.Print Clipboard.GetText, temps, GetPrevWord
End Function

Public Sub SelLastWord()

  KeyDown vbKeyControl
   
  KeyPressShifted vbKeyLeft, EXTENDEDKEY
  
  KeyUp vbKeyControl
 
End Sub

Public Function KeyboardHook(Optional rehook As Boolean = False) As Integer
If keyboard <> 0 Then Exit Function
  
Debug.Print "hooking"

  If TypeMode <> TypeModes_OFF Then
    ToggleNumLock 1
  End If
  
  keyboard = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardProc, App.hInstance, 0&)
  
  keyBuffer.CleanBuffer
  
  If rehook = False Then
  
    arr(0) = 10 'space
    arr(1) = 6
    arr(2) = 7
    arr(3) = 8
    arr(4) = 3
    arr(5) = 4
    arr(6) = 5
    arr(7) = 0
    arr(8) = 1
    arr(9) = 2
    arr(10) = 9
    arr(11) = 11
    arr(12) = 12
    arr(13) = 13
    arr(14) = 14
    
    RotLetters = Split(ROTATELTTRS, ",")
    FirstLetters = Split(FIRSTLTTRS, ",")
    NumLetters = Split(LTTRS, ",")
    TypeState = TypeState_INIT
    'TypeMode = TypeModes_SINGLETAP
    inKeyb = False
    'inDisableKeyb = False
    
    LastMode = TypeMode
  End If
End Function

Public Function Unhook() As Integer
  If keyboard = 0 Then Exit Function

Debug.Print "unhooking"

  Call UnhookWindowsHookEx(keyboard)
  keyboard = 0
  Unhook = 1
  
  ' NumLock handling: - turn off num lock
 ' ToggleNumLock 0
  If NumLockOn = True Then
    KeyPress vbKeyNumlock, EXTENDEDKEY
  End If

End Function

'Private Sub PasteText(s As String)
''  Dim tempC As String
''  Dim ccc As New ApiClipboard
''
''
''  Call ccc.GetTextData(EPredefinedClipboardFormatConstants.CF_TEXT, tempC)
''  Call ccc.SetTextData(EPredefinedClipboardFormatConstants.CF_TEXT, s)
''  Debug.Print s, tempC
''
''
''  keybd_event VK_CONTROL, 0, 0, 0 'key down
''  DoEvents
''  keybd_event 86, 0, 0, 0  'key down
''  keybd_event 86, 0, &H2, 0 'key up
''
''  keybd_event VK_CONTROL, 0, &H2, 0 'key up
''
''  Call ccc.SetTextData(EPredefinedClipboardFormatConstants.CF_TEXT, tempC)
''  ccc.ClipboardClose
'
'  Dim tempC As String
'  tempC = Clipboard.GetText(vbCFText)
'  Clipboard.Clear
'  'Debug.Print Clipboard.GetFormat(vbCFText)
'  Call Clipboard.SetText(trim$(s))
'  'Debug.Print s, tempC, Clipboard.GetText, Clipboard.GetFormat(vbCFText)
'  keybd_event VK_CONTROL, 0, 0, 0 'key down
'  DoEvents
'  keybd_event 86, 0, 0, 0  'key down
'  keybd_event 86, 0, &H2, 0 'key up
'
'  keybd_event VK_CONTROL, 0, &H2, 0 'key up
'
'  DoEvents
'  Call Clipboard.SetText(tempC)
'End Sub



'Public Function CopySelected() As String
'  Dim temps As String
'  temps = Clipboard.GetText
'
'
'  keybd_event VK_CONTROL, 0, 0, 0 'key down
'
'  keybd_event 67, 1, 0, 0 'key down
'  keybd_event 67, 1, &H2, 0 'key up
'
'  keybd_event VK_CONTROL, 0, &H2, 0 'key up
'
'  If Clipboard.GetText = temps Then
'    CopySelected = ""
'  Else
'    CopySelected = Clipboard.GetText
'  End If
'  Clipboard.SetText temps
'
'  Debug.Print temps, Clipboard.GetText
'End Function




