<div align="center">

## Get Windows, System, User and Temp Directories


</div>

### Description

Functions to get the Windows Directory, System Directory, Temp Directory, and User Directory.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[syntax\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/syntax.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/syntax-get-windows-system-user-and-temp-directories__1-55605/archive/master.zip)

### API Declarations

```
Public Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
```


### Source Code

```
'Get the windows directory
Public Function sWindowsDirectory() as String
 Dim sOut As String
 sOut = Space(260)
 GetWindowsDirectory sOut, 260
 sOut = Left(sOut, InStr(sOut, Chr(0)) - 1)
 sWindowsDirectory = sOut
End Function
'Get the system directory
Public Function sSystemDirectory() as String
 Dim sOut As String
 sOut = Space(260)
 GetSystemDirectory sOut, 260
 sOut = Left(sOut, InStr(sOut, Chr(0)) - 1)
 sSystemDirectory = sOut
End Function
'Get the temp directory
Public Function sTempDirectory() as String
 Dim sOut As String
 sOut = Space(260)
 ExpandEnvironmentStrings "%TEMP%", sOut, 260
 sOut = Left(sOut, InStr(sOut, Chr(0)) - 1)
 sTempDirectory = sOut
End Function
'Get the user directory
Public Function sUserDirectory() as String
 Dim sOut As String
 sOut = Space(260)
 ExpandEnvironmentStrings "%USERPROFILE%", sOut, 260
 sOut = Left(sOut, InStr(sOut, Chr(0)) - 1)
 sUserDirectory = sOut
End Function
```

