<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

%>
<html>
<head>
<title>Portfolio View</title>
<script>
function utilityEdit(uid)
{	document.location = 'utilityEdit.asp?uid='+uid
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css"> 
</head>
<body>
<FORM>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td colspan="2"><span class="standardheader"><b>&nbsp;Utility View</b></span></td></tr></table>
<%
rst1.Open "SELECT * FROM tblutility", cnn1
if not rst1.EOF then%>
	<table width="100%" border="0" cellpadding="3" cellspacing="1">
	<tr bgcolor="#cccccc">
		<td><span class="standard"><b>Utility Display</b></span></td>
		<td><span class="standard"><b>Utility</b></span></td>
		<td><span class="standard"><b>Suffix</b></span></td>
	</tr>

	<%do until rst1.EOF%>
	<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="utilityEdit('<%=rst1("utilityid")%>');">
		<td><span class="standard"><%=rst1("utilitydisplay")%></span></td>
		<td><span class="standard"><%=rst1("utility")%></span></td>
		<td><span class="standard"><%=rst1("utilitysuffix")%></span></td>
	</tr>
	
	<%rst1.movenext
	loop%>
</table>
<input type="button" value="Add Utility" onclick="utilityEdit('');">
<%
else
	Response.Write "<BR>There are no utilities set up."
end if
rst1.close
%>
</FORM>
</body>
</html>
