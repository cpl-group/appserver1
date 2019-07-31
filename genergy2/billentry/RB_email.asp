<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, building, action, bperiod, byear, bname, mail_sent, utilityid, uname,cnn2
pid = secureRequest("pid")
building = secureRequest("building")
action = secureRequest("action")
byear = secureRequest("byear")
bperiod = secureRequest("bperiod")
utilityid = secureRequest("utilityid")
'response.write "pid:" & pid
'response.write "building:" & building
'response.write "byear:" & byear
'response.write "bperiod:" & bperiod
'response.write "utilityid:" & utilityid
'response.end

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set cnn2 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,building,"billing")
cnn2.open getConnect(pid, building, "dbCore")

rst1.open "SELECT strt FROM buildings WHERE bldgnum='"&building&"'", cnn1
if not rst1.eof then bname = rst1("strt")
rst1.close
rst1.open "SELECT utilitydisplay FROM tblutility WHERE utilityid='"&utilityid&"'", cnn2
if not rst1.eof then uname = rst1("utilitydisplay")
rst1.close
if action="Email" then
	dim Mail, cdoConfig, Fields,body
	Set cdoConfig = Server.CreateObject("CDO.Configuration")  
	Set Fields = cdoConfig.Fields
	With Fields  
		.Item(cdoSMTPServer) = "2012dc"  
		.Update  
	End With  
	Set Mail = Server.CreateObject("CDO.Message")
	With Mail
		Set .Configuration = cdoConfig
		.To = "rb@cplems.com"
		.From = "filestore@cplems.com"
		 Body= "The "&uname&" utility bill has been entered for "&bname&" bill period: "&bperiod&", bill year: "&byear&". "&vbcr&vbcr&vbcr&"Sent by"&vbcr&getXmlUserName()
		.Subject = Body
		.TextBody = Body
		.Send
	end With
	set Mail = nothing
'	dim email, body
'	set email = server.createObject("CDONTS.NewMail")
'	email.To= "r&b@genergy.com"
'	email.From= "gsa@genergy.com"
'	Body= "The "&uname&" utility bill has been entered for "&bname&" bill period: "&bperiod&", bill year: "&byear&". "&vbcr&vbcr&vbcr&"Sent by"&vbcr&getXmlUserName()
'	email.Subject = Body
'	email.Body = Body
'	email.Bodyformat=0
'	email.Mailformat=0
'	email.Send 
	mail_sent = true
end if
'TK: 04/28/2006
on error resume next
set rst1 = nothing
set email = nothing
If cnn1.State = 1 then
	cnn1.Close 
End If
If cnn2.State = 1 then
	cnn2.Close 
End If
Set cnn1 = nothing
Set cnn2 = nothing
'#TK: 04/28/2006
%>
<html>
<head>
	<title>R&amp;B Email</title>
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
</head>
<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="RB_email.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#6699cc">
  <td nowrap><span class="standardheader">R&amp;B Email - <%=left(bname,14)%> (<%=uname%>)</span></td>
</tr>
</table>
<%if mail_sent then%>
<table cellpadding="3" cellspacing="0" align="center" width="100%">
<tr><td align="center">Eamil Sent to group <nobr>r&amp;b@genergy.com</nobr></td></tr>
<tr><td  align="center"><input type="button" name="action" value="Close" onClick="window.close()"></td></tr>
<tr>

<%else%>
<table cellpadding="3" cellspacing="0" align="center">
<tr><td>Bill Year</td>
	<td>Bill Period</td></tr>
<tr><td><input name="byear" value="<%'=scode%>" type="text" size="8"></td>
	<td><input name="bperiod" value="<%'=scode%>" type="text" size="8"></td>
</tr>
<tr>
	<td colspan="2"><input type="submit" name="action" value="Email"></td>
</tr>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="building" value="<%=building%>">
<input type="hidden" name="utilityid" value="<%=utilityid%>">
</table>
<%end if%>
</form>
</body>
</html>