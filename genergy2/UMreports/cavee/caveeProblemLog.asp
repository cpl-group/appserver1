<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim rst1, cnn1, cmd1
set rst1 = server.createobject("ADODB.recordset")
set cnn1 = server.createobject("ADODB.connection")
set cmd1 = server.createobject("ADODB.command")

dim bldg, meterid, pid, action, close, note
meterid = trim(request("meterid"))
bldg = trim(request("bldg"))
pid = trim(request("pid"))
action = trim(request("action"))
cnn1.open getConnect(pid,bldg,"billing")

if trim(request("close")) = "1" then close = 1 else close = 0
note = trim(request("note"))
if action="Update" then
	cmd1.activeconnection = getLocalConnect(bldg)
	'response.write "UPDATE caveeProblemLog SET close="&close&", note='"&replace(note,"'","''")&"' WHERE meterid="&meterid&" and closed=0"
	'response.end
	cmd1.commandtext = "UPDATE Cavee_Problem_Log SET closed="&close&", user_note='"&replace(note,"'","''")&"' WHERE meterid="&meterid&" and closed=0"
	cmd1.execute
	meterid = ""
	bldg = ""
	pid = ""
end if
cmd1.ActiveConnection = getConnect(pid,bldg,"billing")
%>
<html>
<head><title>Cavee Problem Log</title></head>
<!-- <link rel="Stylesheet" href="/genergy2/setup.css" type="text/css"> -->
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<body>
<form method="get">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#6699cc">
	<td height="26">
		<table border=0 cellpadding="0" cellspacing="0" width="100%">
		<tr><td class="standardheader">Cavee Problem Log</td>
		</tr>
		</table>
	</td>
	<%if meterid<>"" then%><td align="right"><button id="qmark2" onclick="document.location='caveeProblemLog.asp'" 								style="cursor:hand;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">View All Open Cavee Tickets</button>
							   				 <button id="qmark2" onclick="document.location='/genergy2/setup/meteredit.asp?meterid=<%=meterid%>&pid=<%=pid%>&bldg=<%=bldg%>'" style="cursor:hand;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">Meter Setup</button></td><%end if%>
</tr>
<tr><td>
<%if meterid="" then%>
	<table bgcolor="#cccccc" cellpadding="3" cellspacing="1" border="0" width="800">
		<tr bgcolor="#3399cc" class="standardheader">
			<td>Building</td>
			<td>Meter</td>
			<td>Tenant</td>
			<td>Date</td>
			<td>Problem</td>
		</tr>
	<%
	rst1.open "SELECT isnull(description,'?') as d, * FROM Cavee_Problem_Log cpl, meters m, buildings b, tblleasesutilityprices lup, tblleases l WHERE lup.leaseutilityid=m.leaseutilityid and l.billingid=lup.billingid and m.meterid=cpl.meterid and b.bldgnum=m.bldgnum and closed=0 ORDER BY date desc", cnn1
	do until rst1.eof%>
		<tr hlc="lightgreen" nlc="white" style="background-color:white;cursor:hand" 
			onMouseOver="this.style.backgroundColor='lightgreen'" 
			onMouseOut="this.style.backgroundColor='white'"
			onClick="document.location='caveeProblemLog.asp?meterid=<%=rst1("meterid")%>&bldg=<%=rst1("bldgnum")%>&pid=<%=rst1("portfolioid")%>'">
			<td><%=rst1("strt")%></td>
			<td><%=rst1("meternum")%></td>
			<td><%=rst1("billingname")%></td>
			<td><%=rst1("date")%></td>
			<td><%=rst1("d")%></td>
		</tr>
		<%rst1.movenext
	loop
	rst1.close%>
	</table>
<%else
	dim count
'	response.write "SELECT isnull(description,'?') as d, * FROM Cavee_Problem_Log cpl, meters m, buildings b WHERE m.meterid="&meterid&" and m.meterid=cpl.meterid and b.bldgnum=m.bldgnum ORDER BY date desc"
'	response.end
	rst1.open "SELECT isnull(description,'?') as d, isnull(user_note,'-') as unote, * FROM Cavee_Problem_Log cpl, meters m, buildings b WHERE m.meterid="&meterid&" and m.meterid=cpl.meterid and b.bldgnum=m.bldgnum ORDER BY date desc", getLocalConnect(bldg)
	if not rst1.eof then%>
		<strong>Problem report for meter <em><%=rst1("meternum")%></em> of <%=rst1("strt")%>.</strong><br>
		<table cellpadding="3" cellspacing="0" border="0">
		<tr><td>Date</td>
			<td><%=rst1("date")%></td>
		</tr>
		<tr><td>Description</td>
			<td><%=rst1("Description")%></td>
		</tr>
		<tr><td>Close Ticket</td>
			<td><input name="close" value="1" type="checkbox"></td>
		</tr>
		<tr><td valign="top">Note</td>
			<td><textarea name="note" cols="30" rows="4"><%=rst1("user_note")%></textarea></td>
		</tr>
		<tr><td></td>
			<td><input name="action" value="Update" type="submit"></td>
		</tr>
		</table>
		<%
		rst1.movenext
	end if
	if rst1.eof then%>
		<strong>No Problem History</strong><br>
	<%else%>
	<strong>History</strong><br>
	<table bgcolor="#cccccc" cellpadding="3" cellspacing="1" border="0" width="600">
	<tr bgcolor="#3399cc" class="standardheader">
	<td>Date</td>
	<td>Description</td>
	<td>Note</td>
	</tr>
	<%
	do until rst1.eof%>
		<tr><td><%=rst1("date")%></td>
			<td><%=rst1("d")%></td>
			<td><%=rst1("unote")%></td>
		</tr>
		<%rst1.movenext
	loop
	%>
	</table>
	<%end if%>
<%end if%>
</td></tr></table>
<input type="hidden" name="meterid" value="<%=meterid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="pid" value="<%=pid%>">
</form>
</body>
</html>