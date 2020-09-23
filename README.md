<div align="center">

## Launch IE in \*Edit\* Mode


</div>

### Description

This is a VB Script to Launch IE in Edit Mode, that is it functions kinda like FrontPage, supports command-line arguments it also can be done in VB, check it out so simple!
 
### More Info
 
Option Explicit

Main

Public Sub Main ()

Dim Args , FileName , ie

Set Args = WScript.Arguments

If Args.Count > 0 Then

FileName = Args(0)

Else

FileName = "about:blank"

End If

Set Args = Nothing

Set ie=CreateObject("InternetExplorer.Application")

ie.Visible=true

ie.Navigate FileName

ie.Document.designMode = "On"

set ie=Nothing

End Sub


<span>             |<span>
---                |---
**Submitted On**   |2004-12-24 06:47:02
**By**             |[M\. J\. Highlander](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/m-j-highlander.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0, VB Script
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Launch\_IE\_18329312242004\.zip](https://github.com/Planet-Source-Code/m-j-highlander-launch-ie-in-edit-mode__1-57889/archive/master.zip)








