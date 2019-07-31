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

dim uid
uid = secureRequest("uid")

dim utility, utilityDisplay, utilitySuffix
if trim(uid)<>"" then
	rst1.Open "SELECT * FROM tblutility WHERE utilityid='"&uid&"'", cnn1
	if not rst1.EOF then
		utility = rst1("utility")
		utilityDisplay = rst1("utilityDisplay")
		utilitySuffix = rst1("utilitySuffix")
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Utility Edit</title>
</head>
<link rel="Stylesheet" href="setup.css" type="text/css">
<body>
<form name="form2" method="post" action="UtilitySave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#3399cc">
	<td colspan="2"><span class="standardheader">
		<%if trim(uid)<>"" then%>
			Update Utility
		<%else%>
			Add New Utility
		<%end if%>
	</span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Utility</span></td> 
	<td><input type="text" name="utility" value="<%=utility%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Utility Display</span></td>
	<td><input type="text" name="utilityDisplay" value="<%=utilityDisplay%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Utility Suffix</span></td>
	<td><input type="text" name="utilitySuffix" value="<%=utilitySuffix%>"></td>
</tr>
<tr bgcolor="#cccccc"> 
	<td><span class="standard">&nbsp;</span></td>
	
	<td>
		<%if trim(uid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard">
		<%end if%>
	</td>
</tr>
</table>
<input type="hidden" name="uid" value="<%=uid%>">
</form>
</body>
</html>
