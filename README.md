<div align="center">

## Convert VB6 to VB5 every time a project loads\! \*\*\* Updated version\! \*\*\*


</div>

### Description

This program converts the VB6 project file to VB5, before opening VB to open the file. If the file is already VB5 compatible it will leave it alone. Compile the file to the same folder

where Vb5.exe is and then hold down SHIFT

while right clicking on a .vbp file. Choose Open with... then click on other and choose the file that you compiled. Now whenever you open a .vbp file it will convert it then allow VB to open it. NOTE: if there are more tags that are not VB5 compatible, please tell me. The old version did not work with project files that had a space in the path, but this version is compatible. If you like this app, please rate it.
 
### More Info
 
Call from windows

When you compile it the .exe file needs to be in the same folder as Vb5.exe and this only works when the file is opened from Windows. If you open the file in VB itself, the conversion will not take place.

A call to VB

None that I know of


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/convert-vb6-to-vb5-every-time-a-project-loads-updated-version__1-9411/archive/master.zip)





### Source Code

```
' Make a project with only a module and put this
' in it:
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Function GetShortPath(strFileName As String) As String
 Dim lngRes As Long, strPath As String
 strPath = String$(165, 0)
 lngRes = GetShortPathName(strFileName, strPath, 164)
 GetShortPath = Left$(strPath, lngRes)
End Function
Public Function GetPathAndFileName(ByVal PathAndFileName, ByRef FileName As String) As String
 Dim lPos As Long
 Dim lLastPos As Long
 lPos = InStr(1, PathAndFileName, "\")
 While lPos <> 0
 lLastPos = lPos
 lPos = InStr(lLastPos + 1, PathAndFileName, "\")
 Wend
 GetPathAndFileName = Left(PathAndFileName, lLastPos - 1)
 FileName = Mid(PathAndFileName, lLastPos + 1)
End Function
Sub Main()
 On Error Resume Next
 Dim property As String
 Dim newfile As String
 Open Command For Input As #1
 Do Until EOF(1)
 Line Input #1, property
 If property = "Retained=0" Then
 Else
 If property = "Retained=1" Then
  Else
  If property = "DebugStartupOption=0" Then
  Else
  If property = "DebugStartupOption=1" Then
   Else
   newfile = newfile & property & vbCrLf
  End If
  End If
 End If
 End If
 Loop
 Close #1
 Open Command For Output As #1
 Print #1, newfile
 Close #1
 Dim RetVal
 Dim Path As String
 Dim File As String
 Dim ShortPath
 Dim apppath, cmdline
 If Len(App.Path) <> 2 Then 'if path is not root, add a "\"
 apppath = App.Path & "\"
 Else
 apppath = App.Path
 End If
 Path = GetPathAndFileName(Command, File)
 ShortPath = GetShortPath(Path)
 cmdline = apppath & "Vb5.exe " & ShortPath & "\" & File
 RetVal = Shell(cmdline, vbNormalFocus)
 End
End Sub
```

