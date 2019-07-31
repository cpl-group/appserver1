<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, uid, scode, cid, action
pid = secureRequest("pid")
dim cnn1, rst1, strsql, cmd, sql
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
uid = secureRequest("uid")
cid = secureRequest("cid")
scode = secureRequest("scode")
action = secureRequest("action")
cnn1.open getConnect(pid,0,"billing")
cmd.activeconnection = cnn1
if action="Add" then
	cmd.commandtext = "INSERT INTO servicecode (utility, code) VALUES ("&uid&",'"&scode&"')"
elseif action="Edit" then
	cmd.commandtext = "UPDATE servicecode SET code='"&scode&"' WHERE id="&cid
else
	scode=""
end if
if cmd.commandtext<>"" then cmd.execute
%>
<html>
<head>
	<title>OUC Billing</title>
<link rel="Stylesheet" href="../setup.css" type="text/css">
</head>
<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="servicecode.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">OUC Service Codes</span></td>
</tr>
</table>
<table cellpadding="3" cellspacing="0" align="center">
<tr><td>Utility</td>
	<td>Service Code</td></tr>
<tr><td>
<select name="uid" onchange="submit()">
	<optgroup label="Utilities"></optgroup><%
	rst1.open "SELECT utilityid, UtilityDisplay, isnull(code,'') as c FROM tblutility u LEFT JOIN servicecode s ON s.utility=u.utilityid ORDER BY UtilityDisplay", cnn1
	if not rst1.eof then 
		if uid="" then uid = cint(rst1("utilityid"))
	else
		uid = 0
	end if
	do until rst1.eof%>
		<option value="<%=rst1("utilityid")%>" <%if cint(rst1("utilityid"))=cint(uid) then response.write "SELECTED"%>><%=rst1("utilitydisplay")%> <%if rst1("c")<>"" then%> [<%=rst1("c")%>]<%end if%></option><%
		rst1.movenext
	loop
	rst1.close%>
</select>
</td><td>
<%
rst1.open "SELECT * FROM servicecode WHERE utility="&uid, cnn1
if not rst1.eof then 
	scode = rst1("code")
	cid = rst1("id")
end if
rst1.close
%>
<input name="scode" value="<%=scode%>" type="text" size="8">
</td><td>
<%if trim(scode)="" then%>
	<input type="submit" name="action" value="Add">
<%else%>
	<input type="submit" name="action" value="Edit">
<%end if%>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="cid" value="<%=cid%>">
</td></tr></table>
<div id="processing" style="display:none;width: 90; height: 33; border: 1px solid; left: 105px; top: 75px; position:absolute; background-color: #F5F5DC; text-align: center; vertical-align: middle;">Processing request</div>
</form>
</body>
</html>