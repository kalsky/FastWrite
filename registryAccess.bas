Attribute VB_Name = "registryAccess"
'
' Created by E.Spencer - This code is public domain.
'
Option Explicit
'Security Mask constants
Private Const READ_CONTROL As Long = &H20000
Private Const SYNCHRONIZE  As Long = &H100000
Private Const STANDARD_RIGHTS_ALL  As Long = &H1F0000
Private Const STANDARD_RIGHTS_READ  As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE  As Long = READ_CONTROL
Private Const KEY_QUERY_VALUE  As Long = &H1
Private Const KEY_SET_VALUE  As Long = &H2
Private Const KEY_CREATE_SUB_KEY  As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS  As Long = &H8
Private Const KEY_NOTIFY  As Long = &H10
Private Const KEY_CREATE_LINK  As Long = &H20
Private Const KEY_ALL_ACCESS  As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ  As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE  As Long = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE  As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const pvpRunHKey = "Software\Microsoft\Windows\CurrentVersion\Run"

' Possible registry data types
Public Enum InTypes
   ValNull = 0
   ValString = 1
   ValXString = 2
   ValBinary = 3
   ValDWord = 4
   ValLink = 6
   ValMultiString = 7
   ValResList = 8
End Enum
' Registry value type definitions
Private Const REG_NONE As Long = 0
Private Const REG_SZ As Long = 1
Private Const REG_EXPAND_SZ As Long = 2
Private Const REG_BINARY As Long = 3
Private Const REG_DWORD As Long = 4
Private Const REG_LINK As Long = 6
Private Const REG_MULTI_SZ As Long = 7
Private Const REG_RESOURCE_LIST As Long = 8
' Registry section definitions
Private Const HKEY_CLASSES_ROOT  As Long = &H80000000
Public Const HKEY_CURRENT_USER  As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE   As Long = &H80000002
Private Const HKEY_USERS  As Long = &H80000003
Private Const HKEY_PERFORMANCE_DATA  As Long = &H80000004
Private Const HKEY_CURRENT_CONFIG  As Long = &H80000005
Private Const HKEY_DYN_DATA  As Long = &H80000006
' Codes returned by Reg API calls
Private Const ERROR_NONE  As Long = 0
Private Const ERROR_BADDB  As Long = 1
Private Const ERROR_BADKEY  As Long = 2
Private Const ERROR_CANTOPEN  As Long = 3
Private Const ERROR_CANTREAD  As Long = 4
Private Const ERROR_CANTWRITE  As Long = 5
Private Const ERROR_OUTOFMEMORY  As Long = 6
Private Const ERROR_INVALID_PARAMETER  As Long = 7
Private Const ERROR_ACCESS_DENIED  As Long = 8
Private Const ERROR_INVALID_PARAMETERS  As Long = 87
Private Const ERROR_NO_MORE_ITEMS  As Long = 259
' Registry API functions used in this module (there are more of them)
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
'Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
'Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
'Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long

' This routine allows you to get values from anywhere in the Registry, it currently
' only handles string, double word and binary values. Binary values are returned as
' hex strings.
'
' Example
' Text1.Text = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultUserName")
'
Public Function ReadRegistry(ByVal Group As Long, ByVal Section As String, ByVal key As String) As String
Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
Dim TStr2 As String, TStr1 As String, i As Long
On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
sValue = Space$(2048)
lValueLength = Len(sValue)
lResult = RegQueryValueEx(lKeyValue, key, 0&, lDataTypeValue, sValue, lValueLength)
If (lResult = 0) And (Err.Number = 0) Then
   If lDataTypeValue = REG_DWORD Then
      td = AscW(Mid$(sValue, 1, 1)) + &H100& * AscW(Mid$(sValue, 2, 1)) + &H10000 * AscW(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(AscW(Mid$(sValue, 4, 1)))
      sValue = Format$(td, "000")
   End If
   If lDataTypeValue = REG_BINARY Then
       ' Return a binary field as a hex string (2 chars per byte)
       TStr2 = vbNullString
       For i = 1 To lValueLength
          TStr1 = Hex$(AscW(Mid$(sValue, i, 1)))
          If Len(TStr1) = 1 Then TStr1 = "0" & TStr1
          TStr2 = TStr2 + TStr1
       Next
       sValue = TStr2
   Else
      sValue = Left$(sValue, lValueLength - 1)
   End If
Else
   sValue = "Not Found"
End If
lResult = RegCloseKey(lKeyValue)
ReadRegistry = sValue
End Function

' This routine allows you to write values into the entire Registry, it currently
' only handles string and double word values.
'
' Example
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValString, "NewValueHere"
' WriteRegistry HKEY_CURRENT_USER, "SOFTWARE\My Name\My App\", "NewSubKey", ValDWord, "31"
'
Public Sub WriteRegistry(ByVal Group As Long, ByVal Section As String, ByVal key As String, ByVal ValType As InTypes, ByVal Value As Variant)
Dim lResult As Long
Dim lKeyValue As Long
Dim InLen As Long
Dim lNewVal As Long
Dim sNewVal As String
On Error Resume Next
lResult = RegCreateKey(Group, Section, lKeyValue)
If ValType = ValDWord Then
   lNewVal = CLng(Value)
   InLen = 4
   lResult = RegSetValueExLong(lKeyValue, key, 0&, ValType, lNewVal, InLen)
Else
   ' Fixes empty string bug - spotted by Marcus Jansson
   If ValType = ValString Then Value = Value + ChrW$(0)
   sNewVal = Value
   InLen = Len(sNewVal)
   lResult = RegSetValueExString(lKeyValue, key, 0&, 1&, sNewVal, InLen)
End If
lResult = RegFlushKey(lKeyValue)
lResult = RegCloseKey(lKeyValue)
End Sub

Private Sub SaveString(ByVal Hkey As Long, strPath As String, strValue As String, strData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueExString(keyhand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If r = 87 Then
        DeleteValue Hkey, strPath, strValue
    End If
    r = RegCloseKey(keyhand)
End Sub

Private Sub DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
On Error Resume Next
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Sub

Public Sub RunAtStartup(sAppTitle As String, strsAppName As String)
  SaveString HKEY_CURRENT_USER, pvpRunHKey, sAppTitle, strsAppName
End Sub

Public Sub RemoveFromStartup(sAppTitle As String, strsAppName As String)
  DeleteValue HKEY_CURRENT_USER, pvpRunHKey, sAppTitle
End Sub



'' This routine enumerates the subkeys under any given key
'' Call repeatedly until "Not Found" is returned - store values in array or something
''
'' Example - this example just adds all the subkeys to a string - you will probably want to
'' save then into an array or something.
''
'' Dim Res As String
'' Dim i As Long
'' Res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", i)
'' Do Until Res = "Not Found"
''   Text1.Text = Text1.Text & " " & Res
''   i = i + 1
''   Res = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", i)
'' Loop
'
'Public Function ReadRegistryGetSubkey(ByVal Group As Long, ByVal Section As String, Idx As Long) As String
'Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long, lValueLength As Long, sValue As String, td As Double
'On Error Resume Next
'lResult = RegOpenKey(Group, Section, lKeyValue)
'sValue = Space$(2048)
'lValueLength = Len(sValue)
'lResult = RegEnumKey(lKeyValue, Idx, sValue, lValueLength)
'If (lResult = 0) And (Err.Number = 0) Then
'   sValue = Left$(sValue, InStr(sValue, ChrW(0)) - 1)
'Else
'   sValue = "Not Found"
'End If
'lResult = RegCloseKey(lKeyValue)
'ReadRegistryGetSubkey = sValue
'End Function
'
'' This routine allows you to get all the values from anywhere in the Registry under any
'' given subkey, it currently only returns string and double word values.
''
'' Example - returns list of names/values to multiline text box
'' Dim Res As Variant
'' Dim i As Long
'' Res = ReadRegistryGetAll(HKEY_CURRENT_USER, "Software\Microsoft\Notepad", i)
'' Do Until Res(2) = "Not Found"
''    Text1.Text = Text1.Text & chrw(13) & chrw(10) & Res(1) & " " & Res(2)
''    i = i + 1
''    Res = ReadRegistryGetAll(HKEY_CURRENT_USER, "Software\Microsoft\Notepad", i)
'' Loop
''
'Public Function ReadRegistryGetAll(ByVal Group As Long, ByVal Section As String, Idx As Long) As Variant
'Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long
'Dim lValueLength As Long, lValueNameLength As Long
'Dim sValueName As String, sValue As String
'Dim td As Double
'On Error Resume Next
'lResult = RegOpenKey(Group, Section, lKeyValue)
'sValue = Space$(2048)
'sValueName = Space$(2048)
'lValueLength = Len(sValue)
'lValueNameLength = Len(sValueName)
'lResult = RegEnumValue(lKeyValue, Idx, sValueName, lValueNameLength, 0&, lDataTypeValue, sValue, lValueLength)
'If (lResult = 0) And (Err.Number = 0) Then
'   If lDataTypeValue = REG_DWORD Then
'      td = AscW(Mid$(sValue, 1, 1)) + &H100& * AscW(Mid$(sValue, 2, 1)) + &H10000 * AscW(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(AscW(Mid$(sValue, 4, 1)))
'      sValue = Format$(td, "000")
'   End If
'   sValue = Left$(sValue, lValueLength - 1)
'   sValueName = Left$(sValueName, lValueNameLength)
'Else
'   sValue = "Not Found"
'End If
'lResult = RegCloseKey(lKeyValue)
'' Return the datatype, value name and value as an array
'ReadRegistryGetAll = Array(lDataTypeValue, sValueName, sValue)
'End Function
'
'' This routine deletes a specified key (and all its subkeys and values if on Win95) from the registry.
'' Be very careful using this function.
''
'' Example
'' DeleteSubkey HKEY_CURRENT_USER, "Software\My Name\My App"
''
'Public Function DeleteSubkey(ByVal Group As Long, ByVal Section As String) As String
'Dim lResult As Long, lKeyValue As Long
'On Error Resume Next
'lResult = RegOpenKeyEx(Group, vbNullChar, 0&, KEY_ALL_ACCESS, lKeyValue)
'lResult = RegDeleteKey(lKeyValue, Section)
'lResult = RegCloseKey(lKeyValue)
'End Function
'
'' This routine deletes a specified value from below a specified subkey.
'' Be very careful using this function.
''
'' Example
'' DeleteValue HKEY_CURRENT_USER, "Software\My Name\My App", "NewSubKey"
''
'Public Function DeleteValue(ByVal Group As Long, ByVal Section As String, ByVal key As String) As String
'Dim lResult As Long, lKeyValue As Long
'On Error Resume Next
'lResult = RegOpenKey(Group, Section, lKeyValue)
'lResult = RegDeleteValue(lKeyValue, key)
'lResult = RegCloseKey(lKeyValue)
'End Function

