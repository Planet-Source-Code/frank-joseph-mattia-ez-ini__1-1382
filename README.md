<div align="center">

## EZ \- \.ini


</div>

### Description

Access .ini files in the blink of an eye. Use one line of your input to quickly retrive .ini values. With the same one line of code write to your .ini file. If you have any improvements on this code, E-Mail me at "karatebob@hotmail.com".
 
### More Info
 
When you call this in your code, this is the syntax you will need to use.

Dim X As String

X = mfncGetFromIni(SectionHeader, VariableName, FileName)

Text1.Text = X

Returns the value of a string of an .ini file.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Frank Joseph Mattia](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/frank-joseph-mattia.md)
**Level**          |Unknown
**User Rating**    |4.3 (166 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/frank-joseph-mattia-ez-ini__1-1382/archive/master.zip)

### API Declarations

```
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
```


### Source Code

```
Function mfncGetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
  '**********************************************************************************************
  ' DESCRIPTION:Reads from an *.INI file strFileName (full path & file name)
  ' RETURNS:The string stored in [strSectionHeader], line beginning
  ' strVariableName=
'**********************************************************************************************
  ' Initialise variable
  Dim strReturn As String
  ' Blank the return string
  strReturn = String(255, Chr(0))
  'Get requested information, trimming the returned
  ' string
  mfncGetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
Function mfncWriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
  '*****************************************************************************************************
  ' DESCRIPTION:Writes to an *.INI file called strFileName (full  path & file name)
  ' RETURNS:Integer indicating failure (0) or success (other)  to write
    '*****************************************************************************************************
  mfncWriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function
```

