<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, tripcode, billperiod, billyear, action, output, file
pid = secureRequest("pid")
dim cnn1, rst1, strsql, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
tripcode = trim(request("tripcode"))
billperiod = trim(request("billperiod"))
billyear = trim(request("billyear"))
action = trim(request("action"))
cnn1.open getConnect(pid,0,"billing")
if action="Process Bills" then
	cmd.CommandText = "sp_OUC_RunInvoice"
	cmd.CommandType = adCmdStoredProc
	Set cmd.ActiveConnection = cnn1
	Set prm = cmd.CreateParameter("trip", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("pid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("output", adVarChar, adParamOutput, 2000)
	cmd.Parameters.Append prm
	cmd.Parameters("trip") = tripcode
	cmd.Parameters("by") = billyear
	cmd.Parameters("bp") = billperiod
	cmd.Parameters("pid") = pid
	cmd.execute()
	output = cmd.Parameters("output")
elseif action="Accounting File" then
	cmd.CommandText = "sp_oucbcp"
	cmd.CommandType = adCmdStoredProc
	Set cmd.ActiveConnection = cnn1
	Set prm = cmd.CreateParameter("trip", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("pid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("output", adVarChar, adParamOutput, 2000)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("file", adVarChar, adParamOutput, 500)
	cmd.Parameters.Append prm
	cmd.Parameters("trip") = tripcode
	cmd.Parameters("by") = billyear
	cmd.Parameters("bp") = billperiod
	cmd.Parameters("pid") = pid
	cmd.execute()
	output = cmd.Parameters("output")
	file = cmd.Parameters("file")
end if
%>
<html>
<head>
	<title>OUC Billing</title>
<link rel="Stylesheet" href="../setup.css" type="text/css">
</head>
<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="oucBilling.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">OUC Bill Processing</span></td>
</tr>
</table>
<table cellpadding="3" cellspacing="0" align="center">
<tr><td>Trip Code</td>
	<td>Bill Year</td>
	<td>Bill Period</td></tr>
<tr><td>
<select name="tripcode" onchange="submit()">
	<optgroup label="Trip Codes"></optgroup><%
	rst1.open "SELECT distinct tripcode FROM super_tripcodes WHERE pid="&pid, cnn1
	if not rst1.eof then 
		if tripcode="" then tripcode = cint(rst1("tripcode"))
	else
		tripcode = 0
	end if
	do until rst1.eof%>
		<option value="<%=rst1("tripcode")%>" <%if cint(rst1("tripcode"))=cint(tripcode) then response.write "SELECTED"%>><%=rst1("tripcode")%></option><%
		rst1.movenext
	loop
	rst1.close%>
</select>
</td><td>
<select name="billyear" onchange="submit()">
	<optgroup label="Billing Dates"></optgroup><%
	rst1.open "SELECT distinct billyear FROM super_tripcodes st, billyrperiod byp WHERE st.bldgnum=byp.bldgnum and tripcode="&tripcode&" ORDER BY billyear desc", cnn1
	if not rst1.eof then
		if billyear="" then billyear = cint(rst1("billyear"))
	else
		billyear = 0
	end if
	do until rst1.eof%>
		<option value="<%=rst1("billyear")%>" <%if rst1("billyear")=billyear then response.write "SELECTED"%>><%=rst1("billyear")%></option><%
		rst1.movenext
	loop
	rst1.close%>
</select>
</td><td>
<!-- 
<select name="billperiod">
	<optgroup label="Billing Dates"></optgroup><%
	'rst1.open "SELECT distinct tripcode, billperiod, datestart, dateend FROM super_tripcodes st, "&makeIPUnion("billyrperiod","")&" byp WHERE st.bldgnum=byp.bldgnum and billyear="&billyear&" and tripcode="&tripcode, application("cnnstr_supermod")
	'do until rst1.eof%>
		<option value="<%'=rst1("billperiod")%>" <%'if rst1("billperiod")=billperiod then response.write "SELECT"%>><%'=Left(monthname(month(rst1("datestart"))),3)&" "&day(rst1("datestart"))&" - "&Left(monthname(month(rst1("dateend"))),3)&" "&day(rst1("dateend"))%></option><%
	'	rst1.movenext
	'loop
	'rst1.close%>
</select> -->
<select name="billperiod">
	<optgroup label="Bill Period"></optgroup><%
	rst1.open "SELECT distinct tripcode, billperiod FROM super_tripcodes st, billyrperiod byp WHERE st.bldgnum=byp.bldgnum and billyear="&billyear&" and tripcode="&tripcode, cnn1
	do until rst1.eof%>
		<option value="<%=rst1("billperiod")%>" <%if rst1("billperiod")=billperiod then response.write "SELECT"%>><%=rst1("billperiod")%></option><%
		rst1.movenext
	loop
	rst1.close%>
</select>
</td></tr><tr><td align="center" colspan="3">
<input type="submit" name="action" value="Process Bills" onclick="document.all.processing.style.display='inline'">
<input type="submit" name="action" value="Accounting File" onclick="document.all.processing.style.display='inline'">
<input type="hidden" name="pid" value="<%=pid%>">
</td></tr><tr><td colspan="3">
<div style="overflow-y: auto; height: 50px;">
<%
dim bldgerr, errnum, i
if action="Process Bills" then
	if output <> "" then
		response.write "<font color=""red"">"
		output = split(output,"|")
		for each bldgerr in output
			errnum = split(bldgerr,";")
			if ubound(errnum)>1 then
				response.write errnum(0)&"<br>"
				for i=1 to ubound(errnum)
					select case trim(errnum(i))
					case "-1"
						response.write "<li>Readings for this building are not entered.</li>"
					case "-2"
						response.write "<li>Finacial numbers have not been entered.</li>"
					case "0"
						response.write "<li>Invoice was produced with an error.</li>"
					end select
				next
				response.write "<br>"
			end if
		next
		response.write "</font>"
	else
		response.write "Successful bill processing of trip code "&tripcode&" bill period "&billperiod&", "&billyear
	end if
elseif action="Accounting File" then%>
	<!-- Successful building of data file for trip code <%=tripcode%> bill period <%=billperiod%>, <%=billyear%><br> -->
	<a href="<%=file%>">Click to download accounting file</a><br>
	<%if output<>"" then response.write "File excludes:<li>"&replace(output,"|","</li><li>")&"</li>"
end if
%>
</div>
</td></tr></table>
<div id="processing" style="display:none;width: 90; height: 33; border: 1px solid; left: 105px; top: 75px; position:absolute; background-color: #F5F5DC; text-align: center; vertical-align: middle;">Processing request</div>
</form>
</body>
</html>