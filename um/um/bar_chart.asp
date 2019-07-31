<HTML>
<TITLE> PerlScript Test </TITLE>
<%@language=VBscript%>
<%
Dim oWshShell
Set oWshShell = CreateObject("WScript.Shell")
oWshShell.Run "c:\one.pl" 'Launch Notepad
'response.write(oWshShell)
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
FSO.CopyFile "c:\a.txt", "c:\pdasql\"
%>

</HTML>

