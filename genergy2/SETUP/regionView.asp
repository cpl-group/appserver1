<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Rate Setup")) then '"Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

%>
<html>
<head>
<title>Region View</title>
<script>
function regionEdit(rid)
{	document.location = 'regionedit.asp?rid='+rid
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<FORM>
<table width="100%" cellpadding="3" cellspacing="0" border="0" bgcolor="#FFFFFF">
<tr>
  <td bgcolor="#3399cc"><span class="standardheader">&nbsp;Rate Setup</span></td>
</tr>
</table>
<%
rst1.Open "SELECT * FROM regions", cnn1
if not rst1.EOF then%>
	<table width="100%" border="0" cellpadding="3" cellspacing="1">
	<tr bgcolor="#dddddd">
		<td width="30%"><span class="standard"><b>Region Name</b></span></td>
		<td width="70%"><span class="standard"><b>Region Code</b></span></td>
	</tr>

	<%do until rst1.EOF%>
	<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="regionEdit('<%=rst1("id")%>');">
		<td><span class="standard"><%=rst1("city")%></span></td>
		<td><span class="standard"><%=rst1("city_code")%> </span></td>
	</tr>
	
	<%rst1.movenext
	loop%>
	<tr><td colspan="2"><input type="button" value="Add Region" onclick="regionEdit('');"><input type="button" value="Copy Region" onclick="regionEdit('copy');"></td>
</table>
<%
else
	Response.Write "<BR>There are no regions set up for this Portfolio<br><input type=""button"" value=""Add Region"" onclick=""regionEdit('');"">"
end if
rst1.close
%>
</FORM>
</body>
</html>
