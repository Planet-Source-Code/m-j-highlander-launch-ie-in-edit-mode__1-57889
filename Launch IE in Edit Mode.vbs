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
