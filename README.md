<div align="center">

## DUN \(Dial\-Up Networking\) Statistics


</div>

### Description

Dial up networking (or RAS) statistics' availability differs depending on the windows platform. In windows 95 and 98 you can access the statistics via the Dyn_Data section of the registry. For windows NT you have to use one of the performance monitoring techniques, and for windows 2000 you probably can use the performance monitoring techniques also, but you can use the new RAS methods.

If you think that is code can be voted then please rate this... http://munim.cjb.net
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Munim\.VIP](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/munim-vip.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/munim-vip-dun-dial-up-networking-statistics__1-33006/archive/master.zip)

### API Declarations

```
Public Type VBRasStats95
  BytesXmited As Long
  BytesRcved As Long
  FramesXmited As Long
  FramesRcved As Long
  CrcErr As Long
  TimeoutErr As Long
  AlignmentErr As Long
  HardwareOverrunErr As Long
  FramingErr As Long
  BufferOverrunErr As Long
  Runts As Long
  TotalBytesXmited As Long
  TotalBytesRcved As Long
  ConnectSpeed As Long
End Type
Public Const HKEY_DYN_DATA = &H80000006
Public Const KEY_READ = &H20019
Public Declare Function RegQueryValue _
     Lib "advapi32.dll" Alias "RegQueryValueExA" _
     (ByVal hKey As Long, ByVal lpValueName As String, _
     ByVal lpReserved As Long, lpType As Long, _
     lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey _
     Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx _
     Lib "advapi32.dll" Alias "RegOpenKeyExA" _
     (ByVal hKey As Long, ByVal lpSubKey As String, _
     ByVal ulOptions As Long, ByVal samDesired As Long, _
     phkResult As Long) As Long
```


### Source Code

```
Function VBGetRasStats95(clsVBRasStats As VBRasStats95) As Long
  Dim hKey As Long, rtn As Long, lngLen As Long, lResult As Long
  On Error GoTo StatErrorHandler
  lResult = RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StatData", _
                     0&, KEY_READ, hKey)
  With clsVBRasStats
   lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\Buffer", _
              0&, ByVal 0&, .BufferOverrunErr, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\BytesRecvd", _
              0&, ByVal 0&, .BytesRcved, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\BytesXmit", _
              0&, ByVal 0&, .BytesXmited, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\ConnectSpeed", _
              0&, ByVal 0&, .ConnectSpeed, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\CRC", _
              0&, ByVal 0&, .CrcErr, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\Alignment", _
              0&, ByVal 0&, .AlignmentErr, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\FramesRecvd", _
              0&, ByVal 0&, .FramesRcved, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\FramesXmit", _
              0&, ByVal 0&, .FramesXmited, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\Framing", _
              0&, ByVal 0&, .FramingErr, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\Overrun ", _
              0&, ByVal 0&, .HardwareOverrunErr, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\Runts", _
              0&, ByVal 0&, .Runts, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\Timeout", _
              0&, ByVal 0&, .TimeoutErr, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\TotalBytesRecvd", _
              0&, ByVal 0&, .TotalBytesRcved, lngLen)
   lResult = lResult Or rtn: lngLen = 4
   rtn = RegQueryValue(hKey, "Dial-Up Adapter\TotalBytesXmit", _
              0&, ByVal 0&, .TotalBytesXmited, lngLen)
   lResult = lResult Or rtn
  End With
StatErrorHandler:
  RegCloseKey hKey
  VBGetRasStats95 = lResult
End Function
```

