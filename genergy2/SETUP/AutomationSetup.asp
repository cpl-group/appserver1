<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim id, bldg, automated, webpost, email, print, ftp, action, bldgName, ftp_address, ftp_user, ftp_password, process_lag, detailed, datafile, prior_bill
bldg = request("bldg")
id = request("id")
automated = trim(request("automated"))
webpost = trim(request("webpost"))
email = trim(request("email"))
print = trim(request("print"))
detailed = trim(request("detailed"))
ftp = trim(request("ftp"))
prior_bill = trim(secureRequest("prior_bill"))
ftp_address = left(secureRequest("ftp_address"),50)
ftp_user = left(secureRequest("ftp_user"),20)
ftp_password = left(secureRequest("ftp_password"),20)
process_lag = request("process_lag")
datafile = request("datafile")
if id="" then id = 0
if automated="" then automated = 0
if webpost="" then webpost = 0
if email="" then email = 0
if print="" then print = 0
if detailed="" then detailed = 0
if datafile="" then datafile = 0
if ftp="" then ftp = 0
if prior_bill="" then prior_bill = 0
if process_lag="" or not(isnumeric(process_lag)) then process_lag = 3

action = trim(request("action"))

dim cnn1, cnnMainModule, strsql, rst1, cmd
set cnnMainModule = server.createobject("ADODB.connection")
cnnMainModule.open getConnect(0,bldg,"billing")
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

''submit section
if action="Update" then 'subtract from fee
	if id<>0 then
		cmd.CommandText = "UPDATE Automation_Setup SET automated="&automated&", webpost="&webpost&", email="&email&", [print]="&print&", ftp="&ftp&", ftp_address='"&ftp_address&"', ftp_user='"&ftp_user&"', ftp_password='"&ftp_password&"', bill_day='"&process_lag&"', detailed="&detailed&", datafile="&datafile&", prior_bill="&prior_bill&" WHERE id="&id
	else
		cmd.CommandText = "INSERT INTO Automation_Setup (bldgnum, automated, webpost, email, [print], ftp, ftp_address, ftp_user, ftp_password, bill_day, detailed, datafile, prior_bill) VALUES ('"&bldg&"', "&automated&", "&webpost&", "&email&", "&print&", "&ftp&", '"&ftp_address&"', '"&ftp_user&"', '"&ftp_password&"', '"&process_lag&"', "&detailed&", "&datafile&", "&prior_bill&")"
	end if
end if
if cmd.CommandText<>"" then
	cmd.activeconnection = cnn1
	'response.write cmd.CommandText
	'response.end
	cmd.execute()
	%>
	<script>
	window.close();
	</script>
	<%
	response.end
end if

strsql = "SELECT bldgname, b.strt, a.* FROM buildings b, Automation_setup a WHERE b.bldgnum=a.bldgnum and b.bldgnum='"&bldg&"'"
rst1.open strsql, cnn1
if not rst1.eof then
	bldgName = rst1("strt")
	id = rst1("id")
	if rst1("automated") then automated = 1 else automated = 0
	if rst1("webpost") then webpost = 1 else webpost = 0
	if rst1("email") then email = 1 else email = 0
	if rst1("print") then print = 1 else print = 0
	if rst1("detailed") then detailed = 1 else detailed = 0
	if rst1("datafile") then datafile = 1 else datafile = 0
	if rst1("ftp") then ftp = 1 else ftp = 0
	if rst1("prior_bill") then prior_bill = 1 else prior_bill = 0
	ftp_address = rst1("ftp_address")
	ftp_user = rst1("ftp_user")
	ftp_password = rst1("ftp_password")
	process_lag = rst1("bill_day")
end if
rst1.close
%>	
<link rel="Stylesheet" href="setup.css" type="text/css">
<title>Bill Process Setup for <%=bldgname%></title>
<script>
function closeWinda(){
window.close()
}

function checkDisable(){
	frm = document.forms[0];
	if(frm.automated[0].checked){
		frm.process_lag.disabled=false
		frm.webpost.disabled=false;
		frm.email.disabled=false;
		frm.print.disabled=false;
		frm.detailed.disabled=false;
		frm.datafile.disabled=false;
		frm.ftp.disabled=false;
		if(!frm.ftp.checked){
			frm.ftp_address.value="";
			frm.ftp_user.value="";
			frm.ftp_password.value="";
			document.all['ftp_section'].style.display="none";
		}else{
			document.all['ftp_section'].style.display="inline";
		}
	}else{
		frm.process_lag.disabled=true
		frm.webpost.disabled=true;
		frm.email.disabled=true;
		frm.print.disabled=true;
		frm.detailed.disabled=true;
		frm.datafile.disabled=true;
		frm.ftp.disabled=true;
		document.all['ftp_section'].style.display="none";
		
		frm.webpost.checked=false;
		frm.email.checked=false;
		frm.print.checked=false;
		frm.detailed.checked=false;
		frm.datafile.checked=false;
		frm.ftp.checked=false;
		frm.ftp_address.value="";
		frm.ftp_user.value="";
		frm.ftp_password.value="";
	}
}
</script>
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form action="AutomationSetup.asp" method="get">
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#3399cc">
		<td><font color='white'>Bill Processor Setup for <%=bldgname%></font></td>
	</tr>
</table>
<table cellpadding="2" cellspacing="0">
<tr><td colspan="2" style="border-top-style: solid; border-top-color: Gray; border-top-width: 1px;"><strong>Automation Setup</strong></td></tr>
<tr><td><input type="radio" name="automated" onclick="checkDisable()" value="1" <%if automated=1 then%>checked<%end if%>></td><td>Automated&nbsp;Processing</td></tr>
<tr><td></td>
	<td>
	<table cellpadding="0" cellspacing="0">
		<tr><td><input type="text" name="process_lag" size="1" maxlength="2" value="<%=process_lag%>">&nbsp;</td><td># of days after end of bill period to process</td></tr>
		<tr><td valign="top" align="center"><input type="checkbox" name="webpost" value="1" <%if webpost=1 then%>checked<%end if%> <%if automated=0 then%>disabled<%end if%>></td><td>Post tenant invoices</td></tr>
		<tr><td valign="top" align="center"><input type="checkbox" name="email" value="1" <%if email=1 then%>checked<%end if%> <%if automated=0 then%>disabled<%end if%>></td><td>Email tenant invoices to billing contacts</td></tr>
		<tr><td valign="top" align="center"><input type="checkbox" name="print" value="1" <%if print=1 then%>checked<%end if%> <%if automated=0 then%>disabled<%end if%>></td><td>Print tenant invoices and bill summary</td></tr>
		<tr><td valign="top" align="center"><input type="checkbox" name="detailed" value="1" <%if detailed=1 then%>checked<%end if%> <%if automated=0 then%>disabled<%end if%>></td><td>Use detailed tenant invoicing</td></tr>
		<tr><td valign="top" align="center"><input type="checkbox" name="datafile" value="1" <%if datafile=1 then%>checked<%end if%> <%if automated=0 then%>disabled<%end if%>></td><td>Process accounting data file</td></tr>
		<tr><td valign="top" align="center"><input type="checkbox" name="ftp" value="1" onclick="checkDisable()" <%if ftp=1 then%>checked<%end if%> <%if automated=0 then%>disabled<%end if%>></td><td>FTP</td></tr>
		<tr><td></td><td>
			<table cellpadding="0" cellspacing="0" id="ftp_section" style="display:<%if ftp=1 then%>inline<%else%>none<%end if%>">
				<tr><td>Address</td><td><input type="text" name="ftp_address" size="16" maxlength="50" value="<%=ftp_address%>"></td></tr>
				<tr><td>User Name</td><td><input type="text" name="ftp_user" size="16" maxlength="16" value="<%=ftp_user%>"></td></tr>
				<tr><td>Password</td><td><input type="text" name="ftp_password" size="16" maxlength="16" value="<%=ftp_password%>"></td></tr>
			</table>
		</td></tr>
	</table>
</td></tr>
<tr><td><input type="radio" name="automated" onclick="checkDisable()" value="0"  <%if automated=0 then%>checked<%end if%>></td><td>Manual Processing</td></tr>
<tr><td colspan="2" style="border-top-style: solid; border-top-color: Gray; border-top-width: 1px;"><strong>Rate Processing Setup</strong></td></tr>
<tr><td valign="top" align="center"><input type="checkbox" name="prior_bill" value="1" <%if prior_bill=1 then%>checked<%end if%>></td><td>Use Previous Period Utility Bill</td></tr>
<tr><td align="center" colspan="2" style="border-top-style: solid; border-top-color: Gray; border-top-width: 1px;"><input name="action" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" type="submit" value="Update">&nbsp;<input name="close"  style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" type="button" value="Cancel" onclick="javascript:closeWinda();"></td></tr>
</table>
<input type="hidden" name="id" value="<%=id%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
</form>
</body>