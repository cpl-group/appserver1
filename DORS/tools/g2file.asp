<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim strPath,parentPath, path   'Path of directory to show
Dim objFSO    'FileSystemObject variable
Dim objFolder 'Folder variable
Dim objItem   'Variable used to loop through the contents of the folder
Dim clientname,clientdir
clientdir = trim(request("clientdir"))
clientname = trim(request("clientname"))

pid = getkeyvalue("pid")
Set cnn = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")
cnn.Open application("cnnstr_security")
sql = "select clientdir, clientname from (select pid, value as clientdir from g2_pidprefs where type='dfasdir') a, (select pid, value as clientname from g2_pidprefs where type='dfaslabel') b where a.pid = b.pid and b.pid = '" & pid & "'"
rst.open sql, cnn

if not rst.eof then 
	basename = trim(rst("clientname"))
	basedir = trim(rst("clientdir"))
	
else
	Response.write "Data File Access System is not properly configured."
	response.end
end if
if clientdir ="" then 
	clientname = basename
	clientdir = basedir
end if


if Instr(clientdir, "/dors/filebin") then 
	Rootpath = Session("rootpath")
else
	Rootpath = "/dors/filebin/" & trim(clientdir) &"/"
	Session("rootpath") = "/dors/filebin/" & trim(clientdir) &"/"
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
<link rel="Stylesheet" href="/GENERGY2/styles.css" type="text/css">		
</head>
<script>
function openwin(url,mwidth,mheight){
cwin = window.open(url,"childwin","status=no, menubar=no,HEIGHT="+mheight+", WIDTH="+mwidth)
}
try{top.applabel("Document Center: File Bin");}catch(exception){}
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
Set rstFolders = Server.CreateObject("ADODB.Recordset")

rstFolders.Fields.Append "name", adVarChar, 255
rstFolders.Fields.Append "date", adDate
rstFolders.Open

For Each objItem In objFolder.SubFolders
	rstFolders.AddNew
	rstFolders.Fields("name").Value = objItem.Name
	rstFolders.Fields("date").Value = objItem.DateCreated
Next 'objItem


' Show a little description line and the title row of our table
%>
</strong>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"><strong>
        <%if trim(Request("sub"))="true" then%>
        &nbsp;&nbsp;&nbsp;<a href="/dors/Tools/g2file.asp?clientname=<%=clientname%>&clientdir=<%=rootpath%>&sub=false">ROOT</a>/ 
        <%	
	for each x in  split(currentdir,"/")
	prevdir	= left(clientdir,instrrev(clientdir,x,len(clientdir)-1)-1) & x & "/"
	if trim(prevdir) <> trim(clientdir) then 
%>
        <a href="/dors/Tools/g2file.asp?clientname=<%=clientname%>&clientdir=<%=prevdir%>&sub=<%=subval%>"><%=ucase(x)%></a>/ 
        <% else %>
        <%=ucase(x)%> 
        <%
	end if 
next %>
        <%end if%>
        </strong></td>
  </tr>
  <tr>
    <td></td>
  </tr>
</table>
<br>
<% if not rstfolders.EOF then %>
<table width="95%" border="3" align="center" cellpadding="2" cellspacing="0" bordercolor="#6699cc">
	<tr bgcolor="#CCCCCC">
		<td><font color="#FFFFFF"><b>Directory Name:</b></font></td>
		<td><font color="#FFFFFF"><b>Date Created:</b></font></td>
	</tr>

<%
Dim filecount
filecount = 0
rstFolders.MoveFirst
Do while not rstFolders.EOF
	%>
	<tr bgcolor="#FFFFFF">
    <td align="left" valign="middle" ><img src="/GENERGYONEV2/ClientConsole/xmltree/Images/folderclosed.gif" width="18" height="16">&nbsp;<a href="/dors/Tools/g2file.asp?clientname=<%=clientname%>&clientdir=<%=strpath & trim(rstFolders("name")) & "/"%>&sub=true&prevdir=<%=strpath%>"><%=rstFolders("name") %></a></td>
		<td align="left" ><%=rstFolders("date")%></td>
	</tr>
	<%
	rstFolders.MoveNext
Loop
%>
</table>
<br>
<%
end if 
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
rstFiles.Sort = "size ASC, date DESC"
if filecount <> 0 then 
rstFiles.MoveFirst
%>
<table width="95%" border="3" align="center" cellpadding="2" cellspacing="0" bordercolor="#6699cc">
	<tr bgcolor="#CCCCCC">
		<td><font color="#FFFFFF"><b>Document Name:</b></font></td>
		<td><font color="#FFFFFF"><b>Document Size (bytes):</b></font></td>
		<td><font color="#FFFFFF"><b>Date Created:</b></font></td>
	</tr>
<%
' Sort ascending by size and secondarily descending by date
' (by date is mainly for illustration since all our files
'  are different sizes)
Do While Not rstFiles.EOF
if instr(rstFiles("name"), ".dwf")<>0 then 
	url = "/dors/tools/cadview.asp?cad=" & strPath & rstFiles.Fields("name").Value 
	icon = "/dors/images/cad.gif"	
else
	url = strPath & rstFiles.Fields("name").Value 
	icon = "/dors/images/notes.gif"	
end if 

	%>
	<tr bgcolor="#FFFFFF">
		<td align="left" ><img src="<%=icon%>">&nbsp;<a href="<%=url%>" target="_blank"><%= rstFiles.Fields("name").Value %></a></td>
		<td align="right"><%= rstFiles.Fields("size").Value %></td>
		<td align="left" ><%= rstFiles.Fields("date").Value %></td>
	</tr>
	<%
	rstFiles.MoveNext
Loop

' Close our ADO Recordset object
rstFiles.Close
%>
</table>
<%
end if
Set rstFiles = Nothing
'Close the table
%>

</body>
</html>
