<%
Dim strPath,parentPath, path   'Path of directory to show
Dim objFSO    'FileSystemObject variable
Dim objFolder 'Folder variable
Dim objItem   'Variable used to loop through the contents of the folder
Dim clientname,clientdir

clientname = Request("clientname")
clientdir = Request("clientdir")

if Instr(clientdir, "/eri_th/finance") then 
	Rootpath = Session("rootpath")
else
	Rootpath = "/eri_th/finance/" & trim(clientdir) &"/"
	Session("rootpath") = "/eri_th/finance/" & trim(clientdir) &"/"
end if

prevdir	= left(clientdir,instrrev(clientdir,"/",len(clientdir)-1))

subval = request("sub")

if trim(subval) = "true" then 
	currentdir = replace(clientdir,rootpath,"")
else
	Rootpath = session("rootpath")
end if

if trim(prevdir) = trim(rootpath) then 
	subval = "false"
end if
%>
<html>
<head>
<title>AVAILABLE FILES FOR <%=ucase(clientname)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		
</head>
<script>
function openwin(url,mwidth,mheight){
cwin = window.open(url,"childwin","status=no, menubar=no,HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<body>
<%
' A recordset object variable and some selected constants from adovbs.inc.
' I use these for the sorting code.
Dim rstFiles
Const adVarChar = 200
Const adInteger = 3
Const adDate = 7

if Instr(clientdir, rootpath) then
	strPath = clientdir
else 
	strpath = rootpath
end if

' Create our FSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

' Get a handle on our folder
Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))

' Show a little description line and the title row of our table
%>
</strong>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#6699cc"><span class="standardheader"><font size="2">gEnergyOne 
      Data File Access System</font></span></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"><div align="center"><strong>AVAILABLE FILES FOR <%=ucase(clientname)%><%if trim(Request("sub"))="true" then%> : <%=ucase(currentdir)%><%end if%></strong></div></td>
	  </tr>
   <tr><td></td></tr><tr><td bgcolor="#FFFFFF"><div align="center" 
        <% if trim(Request("sub"))="true" then %>
        <strong><a href="./gfile.asp?clientname=<%=clientname%>&clientdir=<%=prevdir%>&sub=<%=subval%>">Back 
        To Previous Folder</a></strong> 
        <%end if%>
      </div></td></tr> 

</table>
<br>
<table width="95%" border="3" align="center" cellpadding="2" cellspacing="0" bordercolor="#6699cc">
	<tr bgcolor="#CCCCCC">
		<td><font color="#FFFFFF"><b>Directory Name:</b></font></td>
		<td><font color="#FFFFFF"><b>Date Created:</b></font></td>
	</tr>

<%
Dim filecount
filecount = 0
For Each objItem In objFolder.SubFolders
	If InStr(1, objItem, "_vti", 1) = 0 Then
	%>
	<tr bgcolor="#FFFFFF">
		<td align="left" ><a href="./gfile.asp?clientname=<%=clientname%>&clientdir=<%=strpath & objItem.Name & "/"%>&sub=true&prevdir=<%=strpath%>"><%= objItem.Name %></a></td>
		<td align="left" ><%= objItem.DateCreated %></td>
	</tr>
	<%
	End If
Next 'objItem
%>
</table>
<br>
<%

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
	filecount = filecount + 1	
Next 'objItem

' All done!  Kill off our File System Object variables.
Set objItem = Nothing
Set objFolder = Nothing
Set objFSO = Nothing

' Now we can sort our data and display it:
%>
<table width="95%" border="3" align="center" cellpadding="2" cellspacing="0" bordercolor="#6699cc">
	<tr bgcolor="#CCCCCC">
		<td><font color="#FFFFFF"><b>File Name:</b></font></td>
		<td><font color="#FFFFFF"><b>File Size (bytes):</b></font></td>
		<td><font color="#FFFFFF"><b>Date Created:</b></font></td>
	</tr>
<%
' Sort ascending by size and secondarily descending by date
' (by date is mainly for illustration since all our files
'  are different sizes)
rstFiles.Sort = "size ASC, date DESC"
if filecount <> 0 then 
rstFiles.MoveFirst
Do While Not rstFiles.EOF
	%>
	<tr bgcolor="#FFFFFF">
		<td align="left" ><a href="<%= strPath & rstFiles.Fields("name").Value %>" target="_blank"><%= rstFiles.Fields("name").Value %></a></td>
		<td align="right"><%= rstFiles.Fields("size").Value %></td>
		<td align="left" ><%= rstFiles.Fields("date").Value %></td>
	</tr>
	<%
	rstFiles.MoveNext
Loop

' Close our ADO Recordset object
rstFiles.Close
else
	%>
	<tr bgcolor="#FFFFFF">
		<td align="center" colspan = 3>NO FILES ARE CURRENTLY AVAILABLE</td>
	</tr>
	<%
end if
Set rstFiles = Nothing
'Close the table
%>
</table>

</body>
</html>