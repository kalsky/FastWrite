VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Private bios As String
'Private mac As String
'
'
'Public Function GetBiosInfo() As String
'Dim retVal As String
'Dim WMIService, Items, SubItems
'On Error Resume Next
'
'  If LenB(bios) > 0 Then
'    retVal = bios
'  Else
'    Set WMIService = GetObject("winmgmts:\\.\root\cimv2")
'    Set Items = WMIService.ExecQuery("Select * from Win32_BIOS", , 48)
'    For Each SubItems In Items
'
'        'retVal = retVal & "Manufacturer: " & SubItems.Manufacturer & ","
'        'retVal = retVal & "SMBIOSBIOSVersion: " & SubItems.SMBIOSBIOSVersion
'        'retVal = retVal & "Version: " & SubItems.Version & ","
'        retVal = retVal & SubItems.SMBIOSBIOSVersion
'    Next
'    'debug.print retVal
'    bios = retVal
'  End If
'GetBiosInfo = retVal
'End Function
'
'Private Function GetMACAddress() As String
'Dim obj, objs
'
'  If LenB(mac) > 0 Then
'    GetMACAddress = mac
'  Else
'    Set objs = GetObject("winmgmts:").ExecQuery( _
'     "SELECT MACAddress FROM Win32_NetworkAdapter " & _
'     "WHERE ((MACAddress Is Not NULL) " & _
'     "AND (Manufacturer <> 'Microsoft'))")
'
'    For Each obj In objs
'     mac = obj.MACAddress
'     Exit For
'    Next obj
'    GetMACAddress = mac
'  End If
'End Function
'
'Private Function Hex2Dec(ByVal strHex As String) As Long
'    ' Check to see if string already begins with &H.
'    If Left$(strHex, 2) <> "&H" Then strHex = "&H" & strHex
'
'    ' Check to see if string contains Decimals and strip them out.
'    If InStr(1, strHex, ".") Then strHex = Left$(strHex, (InStr(1, strHex, ".") - 1))
'
'    Hex2Dec = CLng(strHex)
'End Function
'
'Private Function OnlyNumbers(s As String) As String
'Dim i As Integer
'Dim res As String
'
'  For i = 1 To Len(s)
'    If Mid$(s, i, 1) >= "0" And Mid$(s, i, 1) <= "9" Then
'      res = res & Mid$(s, i, 1)
'    End If
'  Next i
'  OnlyNumbers = res
'End Function
'
'Public Function GetUnregKey() As String
'Dim mac As String
''Dim unregkey As String
'Dim sp() As String
'Dim i As Integer
'Dim d As Long
'Dim ds As String
'Dim temp As String
'
'  mac = GetMACAddress
'  sp = Split(mac, ":")
'  temp = "9542"
'  For i = 1 To 5 Step 2
'    d = Hex2Dec(sp(i))
'    ds = Format(Str(d), "0##")
'    temp = temp & Trim$(ds)
'  Next i
'  For i = 0 To 5 Step 2
'    d = Hex2Dec(sp(i))
'    ds = Format(Str(d), "0##")
'    temp = temp & Trim$(ds)
'  Next i
'  temp = temp & "43500"
'  GetUnregKey = temp
'
'End Function
'
'Public Function ValidateRegKey(testKey As String) As Boolean
'Dim bioskey As String
'Dim unregkey As String
'
'  bioskey = Trim$(OnlyNumbers(GetBiosInfo))
'  If Right$(testKey, Len(bioskey)) = bioskey Then
'    If Replace(testKey, bioskey, vbNullString) = md5(GetUnregKey) Then
'      ValidateRegKey = True
'    Else
'      ValidateRegKey = False
'    End If
'  Else
'    unregkey = GetUnregKey
'    If testKey = unregkey Then
'      ValidateRegKey = False
'      Exit Function
'    End If
'    If testKey = md5(unregkey) Then
'      ValidateRegKey = True
'    Else
'      ValidateRegKey = False
'    End If
'
'  End If
'
'End Function
'
'Public Function md5(b As String) As String
'Dim c As New md5Class
'Dim s(8) As Byte
'Dim s1() As Byte
'Dim s2 As String
'Dim i As Integer
'Dim bsp(3) As String
'Const cnst As String = "022371696"
'Dim j As Integer
'
'  bsp(0) = Left$(b, 9)
'  bsp(1) = Mid$(b, 10, 9)
'  bsp(2) = Right$(b, 9)
'  s2 = vbNullString
'  For j = 0 To 2
'   For i = 0 To 8
'     s(i) = (AscW(Mid$(bsp(j), i + 1, 1)) + AscW(Mid$(cnst, i + 1, 1))) Mod 10 + AscW0
'   Next i
'
'   s1() = c.CreateHash(s)
'
'   For i = 0 To 8
'     s1(i) = (s1(i) Mod 10) + AscW0
'     s2 = s2 & ChrW(s1(i))
'   Next i
'  Next j
'
'  md5 = s2
'  Set c = Nothing
'
'End Function
'
'Public Function IsSameBios() As Boolean
'Dim regVal As String
'Dim bioskey As String
'
'  regVal = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\FastWrite\", "reg")
'  bioskey = OnlyNumbers(GetBiosInfo)
'  If Right$(regVal, Len(bioskey)) = bioskey Then
'    IsSameBios = True
'  Else
'    IsSameBios = False
'  End If
'
'End Function
'
'Public Function IsUnReged() As Boolean
'Dim regVal As String
''Dim retVal As Boolean
'Dim tempkey As String
'
'  regVal = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\FastWrite\", "reg")
'  tempkey = GetUnregKey
'  If regVal = "Not Found" Then
'    WriteRegistry HKEY_LOCAL_MACHINE, "SOFTWARE\FastWrite\", "reg", ValString, tempkey
'    IsUnReged = True
'  Else
'
'    If Not ValidateRegKey(regVal) Then 'unreged
'      If IsSameBios = False Then
'        If regVal <> tempkey Then
'          WriteRegistry HKEY_LOCAL_MACHINE, "SOFTWARE\FastWrite\", "reg", ValString, tempkey
'        End If
'      End If
'      IsUnReged = True
'    Else 'reged
'      IsUnReged = False
'    End If
'
'  End If
'
'End Function
'
'Public Function UnregDaysLeft() As Integer
'Dim fso As New FileSystemObject
'Dim fl As File
'Dim d As Integer
'Dim reg As String
'
'  If Dir(App.Path & "\me.dat") > "" Then
'    Set fl = fso.GetFile(App.Path & "\me.dat")
'    d = Now - fl.DateCreated
'  Else
'    d = 16
'  End If
'  If d < 0 Then d = 16
'  If d > 16 Then d = 16
'  reg = ReadRegistry(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\Dhcp", "Description")
'  If Right$(reg, 1) <> "." Then
'    d = 16
'  End If
'  UnregDaysLeft = 15 - d
'  Set fso = Nothing
'End Function
'
'Public Sub UpdateRegistryAsRegistered(newkey As String)
'Dim reg As String
'
'  WriteRegistry HKEY_LOCAL_MACHINE, "SOFTWARE\FastWrite\", "reg", ValString, newkey & Trim$(OnlyNumbers(GetBiosInfo))
'  reg = ReadRegistry(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\Dhcp", "Description")
'  If Right$(reg, 1) <> "." Then
'    WriteRegistry HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\Dhcp", "Description", ValString, reg & "."
'  End If
'End Sub



'*************************************************
'
'This is the use of
'eSellerate Embedded Webstore SDK capabilities to:

'*1* PURCHASE
'*2* CHECK FOR UPDATE
'*3* LAUNCH UPDATE
'*4* VALIDATE SERIAL NUMBER
'*5* VALIDATE ACTIVATION
'*6* ACTIVATE SERIAL NUMBER
'*7* EXTENDED PURCHASE (WITH PREFILL)
'  Copyright (C) 2004 eSellerate Corporation
'       All rights reserved worldwide
'*************************************************

'Private Declare Function GetModuleFileName Lib "kernel32" Alias _
'      "GetModuleFileNameA" (ByVal hModule As Long, _
'      ByVal lpFileName As String, ByVal nSize As Long) As Long


Public Function CheckOnLoad() As Boolean
    Dim serialNumber As String
    Dim version As String
    Dim esellerate As eWebSDK
    Dim result As eWeb_ReturnValues
    Dim prefix As String
    Dim activationResult As eWeb_ReturnValues
    
    ' Create the eSellerate control.
    Set esellerate = CreateObject("eWebSDK")
    
    ' This function fetches the registry key that contains the serial number for FastWrite.
    serialNumber = GetSetting("FastWrite", vbNull, "SerialNumber", "nonreg")
    version = GetSetting("FastWrite", vbNull, "Version", "noVersion")
    
    ' *4* VALIDATE SERIAL NUMBER example *
    If serialNumber = "nonreg" Then
        result = 0
    Else
        ' This function will validate the serial number found in the registry...
        result = esellerate.eWeb_ValidateSerialNumber(serialNumber, "", "", "78289")
        ' Parameter 1 is the serial number string.
        ' Parameter 2 is left empty since the serial number is not a name based serial number.
        ' Parameter 3 is left empty since no Extra-Data Key needed.
        ' Parameter 4, 78289, is the Publisher Key string. It's optional but used for tighter validation.


        'Here, to confirm that the validated serial number was generated for this product,
        'the first seven characters of the serial number are checked against the Prefix defined
        'for the Serial Number Set.
        prefix = "fwkeyfw"
        
        If result > 0 Then
            If prefix <> Left$(serialNumber, 7) Then 'If the prefix defined for the FastWrite serial number does not match the serial
                result = 0                            'number stored in the registry, set result = 0
            End If
        Else
            result = 0  ' Serial Number is invalid
        End If
    
    End If
    
    ' *5*  VALIDATE ACTIVATION example  *
    If result > 0 Then  ' The serial number is valid, so check whether the local activation works correctly.
        ' Now that we have retrieved the serial number, validate our activation against it.
        activationResult = esellerate.eWeb_ValidateActivation("PUB8214398618", "ACT256708285", serialNumber)
        ' Parameter 1, PUB8214398618 is the Publisher ID.
        ' Parameter 2, ACT256708285, is the Activation ID used for product activation.
        ' Parameter 3, serialnumber is a String object to receive the serial number that was activated.
          
        ' *6* ACTIVATE SERIAL NUMBER example *
        If activationResult < 0 Then 'If the local activation fails, call ActivateSerialNumber to activate the serial number with the
            activationResult = esellerate.eWeb_ActivateSerialNumber("PUB8214398618", "ACT256708285", serialNumber, False)
            ' Parameter 1, PUB8214398618, is the Publisher ID.
            ' Parameter 2, ACT256708285, is the Activation ID used for product activation.
            ' Parameter 3, serialnumber, is the serial number to be activated.
            ' Parameter 4, False, means not to ask user for connection if one doesn't exist.
        End If
        If activationResult >= 0 Then
           'is active
            If version = "noVersion" Then
'                txtVersion.Caption = "VERSION 1.0"
'                txtSerialNumber.Caption = serialNumber
'                txtSerialNumber.Alignment = 2
'                mnuBuy.Visible = False
'                mnuBuyEx.Visible = False
'                mnuUpdate.Visible = True
                 frmMain.mnuRegisterOK
            End If
        Else
            Select Case activationResult
               Case eWeb_ReturnValues.E_ACTIVATION_INVALID_ACTIVATION_KEY ' unknown activationID
                    result = MsgBox("Product activation for FastWrite failed. Please contact the makers of FastWrite for further information.", vbOKOnly + vbExclamation, "Error")
                Case eWeb_ReturnValues.E_ACTIVATESN_UNKNOWN_SN ' unknown serialnumber
                    result = MsgBox("Product activation for FastWrite failed. Our server may not have finished processing your order yet. Please recheck your serial number and try again in a few minutes.", vbOKOnly + vbExclamation, "Error")
                Case eWeb_ReturnValues.E_ACTIVATESN_BLACKLISTED_SN ' blacklisted serialnumber
                    result = MsgBox("Product activation for FastWrite failed. Please contact the makers of FastWrite for further information.", vbOKOnly + vbExclamation, "Error")
                Case eWeb_ReturnValues.E_ACTIVATESN_INVALID_ORDER ' invalid order
                    result = MsgBox("Product activation for FastWrite failed. Please contact the makers of FastWrite for further information.", vbOKOnly + vbExclamation, "Error")
                Case eWeb_ReturnValues.E_ACTIVATESN_LIMIT_MET ' activation limit met
                    result = MsgBox("Product activation for FastWrite failed. FastWrite has already been registered on another computer. Please contact the makers of FastWrite for further information.", vbOKOnly + vbExclamation, "Error")
                Case eWeb_ReturnValues.E_ACTIVATESN_NOT_UNIQUE ' non-unique serial number
                    result = MsgBox("Product activation for FastWrite failed. FastWrite has already been registered on another computer. Please contact the makers of FastWrite for further information.", vbOKOnly + vbExclamation, "Error")
            End Select
        End If
    End If
    
    If result = 0 Or activationResult < 0 Then
      CheckOnLoad = False
    Else
      CheckOnLoad = True
    End If

   ' Remember to release object that was created.
   Set esellerate = Nothing
End Function

Public Function BuyFastWrite() As Boolean
   Dim esellerate As eWebSDK
   Dim resultData As eWebResultData
   Dim result As eWeb_ReturnValues
   'Dim serialcount
   Dim serialNumber As String
   Dim activationKey As String
   
'* 1* PURCHASE example *
'* Command the Embedded Webstore to perform the purchase specified.

   Set esellerate = CreateObject("eWebSDK") 'Create the eSellerate control
   Set resultData = CreateObject("eWebResultData")
  
  result = esellerate.eWeb_Purchase("STR9242549065", "SKU15098327618", "", "", "", "", "", "ACT256708285", "", resultData)
   ' STR9242549065                      -- Embedded Webstore eSeller Refnum
   ' SKU15098327618                     -- The SKU to be placed in the cart (optional)
   ' Evx4dB9L                           -- Preview ID
   ' ""                                 -- Layout Certificate
   ' ""                                 -- Tracking ID
   ' ""                                 -- Affiliate ID
   ' ""                                 -- Coupon ID
   ' ACT256708285                       -- Activation Set ID, to activate serial number
   ' ""                                 -- extra data, not needed here
   ' resultData                         -- The result data object that we created above
   
  ' Any positive value returned from eWeb_Purchase could indicate a successful purchase.
   If result >= 0 Then
    ' When retrieving data about the purchase, you will see the terms "Record" and "Field"
    ' used to designate the information you want to receive from the result data. There is
    ' a list of valid records/fields that youl'll have available. The Record indicates what
    ' "object" you want to retrieve data about (Order, OrderItem, Shipping, etc.). The field
    ' represents the data you want to retrieve from the record. In this case, we want to get
    ' the serial number from the item that we ordered. Since I really only need the serial
    ' number, I could also call:
    ' resultData.eWeb_IndexGetResultDataValue("_OrderItem", "SerialNumber", 0, serialNumber)
    
    result = resultData.eWeb_GetOrderItemsByIndex(0, "", "", "", "", serialNumber, "", "", activationKey, "")
    'result = resultData.eWeb_IndexGetResultDataValue("_OrderItem", "SerialNumber", 0, serialNumber)
    'result = resultData.eWeb_IndexGetResultDataValue("_OrderItem", "ActivationResult", 0, activationKey)
    If activationKey <> "" Then
        'activation was a success
        'txtSerialNumber.Caption = serialNumber                      ' Display the serial number in the FastWrite form
        SaveSetting "FastWrite", vbNull, "SerialNumber", serialNumber
        'txtVersion.Caption = "VERSION 1.0"
        'mnuBuy.Visible = False
        'mnuUpdate.Visible = True
        'mnuBuyEx.Visible = False
        frmMain.mnuRegisterOK
    Else
        result = MsgBox("Product activation for FastWrite failed.", vbOKOnly + vbExclamation, "Error")
    End If
   Else
    Select Case result
        Case eWeb_ReturnValues.E_ACTIVATION_INVALID_ACTIVATION_KEY ' unknown activationID
            result = MsgBox("Product activation for FastWrite failed. Please contact the makers of FastWrite for further information.\n\n(Error = E_ACTIVATION_INVALID_ACTIVATION_KEY)", vbOKOnly + vbExclamation, "Error")
        Case eWeb_ReturnValues.E_ACTIVATESN_UNKNOWN_SN ' unknown serialnumber
            result = MsgBox("Product activation for FastWrite failed. Our server may not have finished processing your order yet. Please recheck your serial number and try again in a few minutes.\n\n(Error = E_ACTIVATESN_UNKNOWN_SN", vbOKOnly + vbExclamation, "Error")
        Case eWeb_ReturnValues.E_ACTIVATESN_BLACKLISTED_SN ' blacklisted serialnumber
            result = MsgBox("Product activation for FastWrite failed. Please contact the makers of FastWrite for further information.\n\n(Error = E_ACTIVATESN_BLACKLISTED_SN)", vbOKOnly + vbExclamation, "Error")
        Case eWeb_ReturnValues.E_ACTIVATESN_INVALID_ORDER ' invalid order
            result = MsgBox("Product activation for FastWrite failed. Please contact the makers of FastWrite for further information.\n\n(Error = E_ACTIVATESN_INVALID_ORDER)", vbOKOnly + vbExclamation, "Error")
        Case eWeb_ReturnValues.E_ACTIVATESN_LIMIT_MET ' activation limit met
            result = MsgBox("Product activation for FastWrite failed. FastWrite has already been registered on another computer. Please contact the makers of FastWrite for further information.\n\n(Error = E_ACTIVATESN_LIMIT_MET", vbOKOnly + vbExclamation, "Error")
        Case eWeb_ReturnValues.E_ACTIVATESN_NOT_UNIQUE ' non-unique serial number
            result = MsgBox("Product activation for FastWrite failed. FastWrite has already been registered on another computer. Please contact the makers of FastWrite for further information.\n\n(Error = E_ACTIVATESN_NOT_UNIQUE", vbOKOnly + vbExclamation, "Error")
    End Select
   End If
   
   'Remember to release memory after purchase.
   Set resultData = Nothing
   Set esellerate = Nothing
   
   If result = 1 Then
     BuyFastWrite = True
   Else
     BuyFastWrite = False
   End If
End Function


Public Sub ActivateFastWrite(serial As String)
   Dim esellerate As eWebSDK
   Dim activationResult As eWeb_ReturnValues
   
   Set esellerate = CreateObject("eWebSDK") 'Create the eSellerate control
   
   activationResult = esellerate.eWeb_ActivateSerialNumber("PUB8214398618", "ACT256708285", serial, False)
   ' Parameter 1, PUB8214398618, is the Publisher ID.
   ' Parameter 2, ACT256708285, is the Activation ID used for product activation.
   ' Parameter 3, serial, is the serial number to be activated.
   ' Parameter 4, False, means not to ask user for connection if one doesn't exist.
   
   If activationResult >= 0 Then
       SaveSetting "FastWrite", vbNull, "SerialNumber", serial
       frmMain.mnuRegisterOK
   Else
       Call MsgBox("Product activation for FastWrite failed.", vbOKOnly + vbExclamation, "Error")
   End If
   
   'Remember to release memory after purchase.
   Set esellerate = Nothing
   
End Sub

Public Sub Unregister()
Dim serialNumber As String
On Error Resume Next
   serialNumber = GetSetting("FastWrite", vbNull, "SerialNumber", "nonreg")
   If serialNumber = "nonreg" Then Exit Sub
   
   Call DeleteSetting("FastWrite", vbNull, "SerialNumber")
End Sub


