<%
' This function takes a filename and returns the appropriate image for
' that file type based on it's extension.  If you pass it "dir", it assumes
' that the corresponding item is a directory and shows the folder icon.
Function ShowImageForType(strName)
	Dim strTemp

	' Set our working string to the one passed in
	strTemp = strName
	
	' If it's not a directory, get the extension and set it to strTemp
	' If it is a directory, then we already have the correct value
	If strTemp <> "dir" Then
		strTemp = LCase(Right(strTemp, Len(strTemp) - InStrRev(strTemp, ".", -1, 1)))
	End If
	
	' Debugging line used to perfect that above string parser
	'Response.Write strTemp
	
	' Set the part of the image file name that's unique to the type of file
	' to it's correct value and set this to strTemp. (yet another use of it!)
	Select Case strTemp
		Case "asp"
			strTemp = "asp"
		Case "dir"
			strTemp = "dir"
		Case "htm", "html"
			strTemp = "htm"
		Case "gif", "jpg"
			strTemp = "img"
		Case "txt"
			strTemp = "txt"
		Case Else
			strTemp = "misc"
	End Select

	' All our logic is done... build the IMG Tag for display to the browser
	' Place it into... where else... strTemp!

	' My images are all GIFs and all start with "dir_" for my own sanity.
	' They end with one of the values set in the select statement above.

	strTemp = "<img src=""images/dir_" & strTemp & ".gif"" width=""16"" height=""16"" border=""0"">"

	' Set return value and exit function
	ShowImageForType = strTemp
End Function
'That's it for functions on this one!
%>


<%' Now to the Runtime code:
Dim strPath   'Path of directory to show
Dim objFSO    'FileSystemObject variable
Dim objFolder 'Folder variable
Dim objItem   'Variable used to loop through the contents of the folder

' A recordset object variable and some selected constants from adovbs.inc.
' I use these for the sorting code.
Dim rstFiles
Const adVarChar = 200
Const adInteger = 3
Const adDate = 7


' You could just as easily read this from some sort of input, but I don't
' need you guys and gals roaming around our server so I've hard coded it to
' a directory I set up to illustrate the sample.
' NOTE: As currently implemented, this needs to end with the /
strPath = "./"

' Create our FSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

' Get a handle on our folder
Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))

' Show a little description line and the title row of our table
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#0099FF">
  <tr>
    <td bgcolor="#0099FF">
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="3" color="#FFFFFF"><b>Operations 
        Reports</b> </font></div>
    </td>
  </tr>
</table>
<br />

<br />

<table border="5" bordercolor="#0099FF" cellspacing="0" cellpadding="2" align="center">
  <tr bgcolor="#000000"> 
    <td><font color="#FFFFFF"><b>File Name:</b></font></td>
		
    <td><font color="#FFFFFF"><b>File Size (bytes):</b></font></td>
		
    <td><font color="#FFFFFF"><b>Date Created:</b></font></td>
		
    <td><font color="#FFFFFF"><b>File Type:</b></font></td>
	</tr>
<%

' First I deal with any subdirectories.  I just display them and when you
' click you go to them via plain HTTP.  You might want to loop them back
' through this file once you've set it up to take a path as input.  It seems
' like the logical thing to do to me at least!
For Each objItem In objFolder.SubFolders
	' Deal with the stupid VTI's that keep giving our visitors 404's
	If InStr(1, objItem, "_vti", 1) = 0 Then
	%>
	
  <tr bgcolor="#FFFFFF"> 
    <td align="left" >&nbsp;<a href="<%= strPath & objItem.Name %>"><%= objItem.Name %></a></td>
    <td align="right"><%= objItem.Size %></td>
		
    <td align="left" ><%= objItem.DateCreated %></td>
		
    <td align="left" ><%= objItem.Type %></td>
	</tr>
	<%
	End If
Next 'objItem

' Now that I've done the SubFolders, do the files!

' In order to be able to sort them easily and still close the FSO relatively
' quickly I'm going to make use of an ADO Recordset object with no attached
' datasource.  While it does have a slightly greater overhead then an array
' or dictionary object, it gives me named access to the fields and has built
' in sorting functionality.
Set rstFiles = Server.CreateObject("ADODB.Recordset")
rstFiles.Fields.Append "name", adVarChar, 255
rstFiles.Fields.Append "size", adInteger
rstFiles.Fields.Append "date", adDate
rstFiles.Fields.Append "type", adVarChar, 255
rstFiles.Open

For Each objItem In objFolder.Files
	rstFiles.AddNew
	rstFiles.Fields("name").Value = objItem.Name
	rstFiles.Fields("size").Value = objItem.Size
	rstFiles.Fields("date").Value = objItem.DateCreated
	rstFiles.Fields("type").Value = objItem.Type
Next 'objItem

' All done!  Kill off our File System Object variables.
Set objItem = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

' Now we can sort our data and display it:

' Sort ascending by size and secondarily descending by date
' (by date is mainly for illustration since all our files
'  are different sizes)
rstFiles.Sort = "size ASC, date DESC"

rstFiles.MoveFirst

Do While Not rstFiles.EOF
	If InStr(1, rstFiles.Fields("name").Value, "html", 1) > 0 Then
	%>
	
  <tr bgcolor="#FFFFFF"> 
    <td align="left" >&nbsp;<a href="<%= strPath & rstFiles.Fields("name").Value %>"><%=left(rstFiles.Fields("name").Value,InStr(1, rstFiles.Fields("name").Value, ".", 1)) %></a></td>
    <td align="right"><%= rstFiles.Fields("size").Value %></td>
		
    <td align="left" ><%= rstFiles.Fields("date").Value %></td>
		
    <td align="left" ><%= rstFiles.Fields("type").Value %></td>
	</tr>
	<%
	end if
	rstFiles.MoveNext
Loop

' Close our ADO Recordset object
rstFiles.Close
Set rstFiles = Nothing

'Close the table
%>
</table>

