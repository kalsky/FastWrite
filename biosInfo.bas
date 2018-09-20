Attribute VB_Name = "biosInfo"
Public Function GetBiosInfo() As String
Dim retVal As String
On Error Resume Next

Set WMIService = GetObject("winmgmts:\\.\root\cimv2")
Set Items = WMIService.ExecQuery("Select * from Win32_BIOS", , 48)
For Each SubItems In Items
    
    retVal = retVal & "Manufacturer: " & SubItems.Manufacturer & ","
    retVal = retVal & "SMBIOSBIOSVersion: " & SubItems.SMBIOSBIOSVersion & ","
    retVal = retVal & "Version: " & SubItems.Version & ","
Next
Debug.Print retVal
GetBiosInfo = retVal
End Function

Public Function GetMACAddress() As String
Dim obj, objs

Set objs = GetObject("winmgmts:").ExecQuery( _
 "SELECT MACAddress FROM Win32_NetworkAdapter " & _
 "WHERE ((MACAddress Is Not NULL) " & _
 "AND (Manufacturer <> 'Microsoft'))")
 
For Each obj In objs
 GetMACAddress = obj.MACAddress
 Exit For
Next obj
End Function

