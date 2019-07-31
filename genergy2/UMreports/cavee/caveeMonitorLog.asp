<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim rst1, cnn1
set rst1 = server.createobject("ADODB.recordset")
set cnn1 = server.createobject("ADODB.connection")

dim bldg, meterid, tdate, lid, tid, pid
meterid = trim(request("meterid"))
bldg = trim(request("bldg"))
tdate = trim(request("date"))
lid = trim(request("lid"))
tid = trim(request("tid"))
pid = trim(request("pid"))
cnn1.open getConnect(pid,bldg,"billing")
%>
<html>
<head><title>Cavee Monitor Log</title></head>
<!-- <link rel="Stylesheet" href="/genergy2/setup.css" type="text/css"> -->
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<body>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#6699cc">
	<td height="26">
		<table border=0 cellpadding="0" cellspacing="0" width="100%">
		<tr><td class="standardheader">Cavee Monitor</td>
		</tr>
		</table>
	</td>
	<td align="right"><%if meterid<>"" and tdate<>"" then%><button id="qmark2" onclick="document.location='caveeMonitorLog.asp?meterid=<%=meterid%>&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>'" 								style="cursor:hand;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">View All Days</button><%end if%>
							   				 <button id="qmark2" onclick="document.location='/genergy2/setup/meteredit.asp?meterid=<%=meterid%>&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>'" style="cursor:hand;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">Meter Setup</button></td>
</tr>
<tr><td>
<%if tdate<>"" then
	rst1.open "SELECT cm.meterid, convert(datetime, left(date,11)) as [date], [date] as totaltime, cm.note, b.strt, l.billingname, m.meternum, billyear, cm.billperiod, cm.failed_test, cm.org_reading, cm.est_reading FROM cavee_monitor cm, meters m, buildings b, tblleasesutilityprices lup, tblleases l WHERE cm.meterid="&meterid&" and convert(datetime, left(date,11))='"&tdate&"' and m.meterid=cm.meterid and m.leaseutilityid=lup.leaseutilityid and lup.billingid=l.billingid and b.bldgnum=m.bldgnum  ORDER BY [date], cm.note", getLocalConnect(bldg)
	'response.write "SELECT cm.meterid, convert(datetime, left(date,11)) as [date], [date] as totaltime, cm.note, b.strt, l.billingname, m.meternum FROM cavee_monitor cm, meters m, buildings b, tblleasesutilityprices lup, tblleases l WHERE cm.meterid="&meterid&" and convert(datetime, left(date,11))='"&tdate&"' and m.meterid=cm.meterid and m.leaseutilityid=lup.leaseutilityid and lup.billingid=l.billingid and b.bldgnum=m.bldgnum  ORDER BY [date] desc, cm.note"
	if not rst1.eof then%>
		<strong>Meter <em><%=rst1("meternum")%></em> of <%=rst1("billingname")%> in <%=rst1("strt")%> for <%=tdate%></strong>
		<table bgcolor="#cccccc" cellpadding="3" cellspacing="1" border="0" width="600">
			<tr bgcolor="#3399cc" class="standardheader">
				<td>Date</td>
				<td>Bill&nbsp;Period</td>
				<td>Org.&nbsp;Reading</td>
				<td>Est.&nbsp;Reading</td>
				<td>Note</td>
			</tr>
		<%do until rst1.eof%>
			<tr bgcolor="white">
				<td><%=rst1("totaltime")%></td>
				<td><%=rst1("billyear")%>, <%=rst1("billperiod")%></td>
				<td><%=rst1("org_reading")%></td>
				<td><%=rst1("est_reading")%></td>
				<td><%=rst1("note")%></td>
			</tr>
			<%rst1.movenext
		loop
		rst1.close%>
		</table><%
	end if
elseif meterid<>"" then
	rst1.open "SELECT cm.meterid, convert(datetime, left(date,11)) as [date], cm.note, b.strt, l.billingname, m.meternum FROM cavee_monitor cm, meters m, buildings b, tblleasesutilityprices lup, tblleases l WHERE cm.meterid="&meterid&" and m.meterid=cm.meterid and m.leaseutilityid=lup.leaseutilityid and lup.billingid=l.billingid and b.bldgnum=m.bldgnum  GROUP BY convert(datetime, left(date,11)), cm.note, cm.meterid, b.strt, l.billingname, m.meternum ORDER BY [date] desc, cm.note", getLocalConnect(bldg)
	if not rst1.eof then%>
		<strong>Meter <em><%=rst1("meternum")%></em> of <%=rst1("billingname")%> in <%=rst1("strt")%></strong>
		<table bgcolor="#cccccc" cellpadding="3" cellspacing="1" border="0" width="600">
			<tr bgcolor="#3399cc" class="standardheader">
				<td>Date</td>
				<td>Problem</td>
			</tr>
		<%
		'response.write "SELECT meterid, convert(datetime, left(date,11)) as [date], note FROM cavee_monitor WHERE meterid="&meterid&" GROUP BY convert(datetime, left(date,11)), note, meterid ORDER BY [date], note"
		'response.end
		do until rst1.eof%>
			<tr hlc="lightgreen" nlc="white" style="background-color:white;cursor:hand" 
				onMouseOver="this.style.backgroundColor='lightgreen'" 
				onMouseOut="this.style.backgroundColor='white'"
				onClick="document.location='caveeMonitorLog.asp?meterid=<%=meterid%>&bldg=<%=bldg%>&date=<%=rst1("date")%>&tid=<%=tid%>&lid=<%=lid%>'">
				<td><%=rst1("date")%></td>
				<td><%=rst1("note")%></td>
			</tr>
			<%rst1.movenext
		loop
		rst1.close%>
		</table>
	<%end if%>
<%end if%>
</td></tr></table>
</body>
</html>