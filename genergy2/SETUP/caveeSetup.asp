<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim i, pid, bldg, tid, lid, meterid, action, Meternum, cavee, minchck, sumcheck, spikecheck, high_lowcheck, dataFrequency, interval, cavee_monitor, spike_adu, rundays(7), tempweekday
pid = request("pid")
bldg = request("bldg")
tid = request("tid")
lid = request("lid")
meterid = request("meterid")
if request("cavee")="1" then cavee=1 else cavee=0
if request("interval") = "" then interval = -1 else interval = cint(trim(request("interval")))
if request("cavee_monitor")="1" then cavee_monitor=1 else cavee_monitor=0
if request("spike_adu")="1" then spike_adu=1 else spike_adu=0
if request("minchck")="1" then minchck=1 else minchck=0
for i = 1 to 7
	if request(left(weekdayname(i),3))="1" then rundays(i) = 1 else rundays(i) = 0
next


spikecheck = request("spikecheck")
if isNumeric(spikecheck) then	
	if cint(spikecheck)=1 then	spikecheck=1 else	spikecheck = 0
else 
	spikecheck=0
end if

high_lowcheck =request("high_lowcheck")
if isNumeric(high_lowcheck) then
	if cint(high_lowcheck)=1 then high_lowcheck=1  else high_lowcheck = 0
else 
	high_lowcheck=0
end if

action = request("action")
if request("datafrequency") = "" then dataFrequency = -1 else dataFrequency = cint(trim(request("datafrequency")))

'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

if trim(action)<>"" then
	dim strsqldays
	for i = 1 to 7
		strsqldays = strsqldays & ", ["&i&"] = "&rundays(i)
	next
	strsql = "UPDATE meters SET cavee="&cavee&", [interval]="&interval&" WHERE meterid="&meterid
	cnn1.execute strsql
	strsql = "update cavee_setup set cavee_on="&cavee&", spikecheck="&spikecheck&", high_lowcheck="&high_lowcheck&", cavee_monitor="&cavee_monitor&", spike_adu="&spike_adu&", [15minchck]="&minchck&" "&strsqldays&" WHERE meterid="&meterid
	cnn1.execute strsql
end if

strsql = "SELECT * FROM Cavee_Setup WHERE meterid="&meterid
rst1.Open strsql, cnn1
if rst1.eof then 
	strsql = "INSERT Cavee_Setup (meterid) values ("&meterid&")"
	cnn1.execute strsql
end if
rst1.close

if trim(meterid)<>"" then
	dim someSql
	someSql = "SELECT m.data_frequency, m.interval, cs.cavee_on, m.meternum, isnull(cs.spikecheck,0) as spikecheck, isnull(cs.high_lowcheck,0) as high_lowcheck, isnull(cs.[15minchck],0) as [15minchck], isnull(cs.cavee_on,0) as cavee_on, isnull(cs.cavee_monitor,0) as cavee_monitor, isnull(cs.spike_adu,0) as spike_adu, [1] as sun, [2] as mon, [3] as tue, [4] as wed, [5] as thu, [6] as fri, [7] as sat FROM Meters m LEFT OUTER JOIN Cavee_Setup cs ON m.MeterId = cs.MeterId WHERE m.meterid="&meterid
	rst1.Open someSql, cnn1
	if not rst1.EOF then
		Meternum = rst1("Meternum")
		if rst1("cavee_on")="True" then cavee=1 else cavee=0
		interval = cint(trim(rst1("interval")))
		dataFrequency = cint(trim(rst1("data_frequency")))
		if rst1("spikecheck")="True" then spikecheck=1 else spikecheck=0
		if rst1("high_lowcheck")="True" then high_lowcheck=1 else high_lowcheck=0
		if rst1("15minchck")="True" then minchck=1 else minchck=0
		if rst1("cavee_monitor")="True" then cavee_monitor=1 else cavee_monitor=0
		if rst1("spike_adu")="True" then spike_adu=1 else spike_adu=0
		for i = 1 to 7
			if rst1(left(weekdayname(i),3))="True" then rundays(i) = 1 else rundays(i) = 0
		next
	end if
	rst1.close
end if

dim bldgname, portfolioname, billingname
if trim(bldg)<>"" then
  rst1.open "SELECT b.bldgname, p.name, billingname FROM buildings b, portfolio p, tblleases l WHERE l.billingid='"&tid&"' and b.portfolioid=p.id and b.bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
		billingname = rst1("billingname")
    
	end if
	rst1.close
end if
%>
<html>
<head>
<title>CAVEE Setup</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#dddddd">
<form method="get" action="caveeSetup.asp">
<table border="0" cellpadding="3" cellspacing="0" width="100%">
	<tr bgcolor="#3399cc"> 
		<td colspan="4"><span class="standardheader">CAVEE Setup | </span><span style="font-weight:normal;"> <%=billingname%></span></td>
	</tr>
	<tr bgcolor="#dddddd"> 
		<td style="border-top:1px solid #cccccc;"><b>Meter <%=meternum%> (ID# <%=meterid%>)</b></td>
	</tr>
	<tr bgcolor="#dddddd"> 
		<td>
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr><td align="right" width="5%">Active:</td>
					<td><input type="Checkbox" name="cavee" value="1" <%if cavee=1 then%>CHECKED<%end if%>></td></tr>
			<tr><td align="right" width="5%">Run&nbsp;Every&nbsp;15&nbsp;minutes:</td>
					<td><input type="checkbox" value="1" name="minchck" <%if minchck=1 then%> CHECKED <%end if%>></td></tr>
			<tr><td align="right" width="5%">Cavee Monitor:</td>
					<td><input type="checkbox" value="1" name="cavee_monitor" <%if cavee_monitor=1 then%> CHECKED <%end if%>></td></tr>
			<tr bgcolor="#dddddd"><td colspan="2"><hr size="1"></td></tr>
			<tr><td valign="top" align="right">Days to Run Cavee:</td>
					<td><%for i = 1 to 7
								tempweekday = left(weekdayname(i),3)
							%>
							<input type="checkbox" value="1" name="<%=tempweekday%>" <%if rundays(i)=1 then response.write "CHECKED"%>><%=tempweekday%>&nbsp; &nbsp; 
							<%next%>
					</td>
			</tr>
			<tr bgcolor="#dddddd"><td colspan="2"><hr size="1"></td></tr>
				<tr><td valign="top" align="right" nowrap>Meter Interval:</td>
						<td><table cellspacing="0" cellpadding="0" border="0">
								<tr><td><input type="radio" name="interval" value=1 <%if cint(interval)=1 then%>CHECKED<%end if%>></td>
										<td>1 minute</td></tr>
								<tr><td><input type="radio" name="interval" value=15 <%if cint(interval)=15 then%>CHECKED<%end if%>></td>
										<td>15 minute</td></tr>
								<tr><td><input type="radio" name="interval" value=60 <%if cint(interval)=60 then%>CHECKED<%end if%>></td>
										<td>60 minute</td></tr>
								</table>
						</td>
				</tr>
				<tr bgcolor="#dddddd"><td colspan="2"><hr size="1"></td></tr>
				<tr><td valign="top" align="right">Checks:</td>
						<td><table cellspacing="0" cellpadding="0" border="0">
								<tr><td><input type="checkbox" value="1" name="spikecheck" <%if spikecheck=1 then%> CHECKED <%end if%>></td>
										<td>Spike Check</td></tr>
								<tr><td><input type="checkbox" value="1" name="high_lowcheck" <%if high_lowcheck=1 then%> CHECKED <%end if%>></td>
										<td>High/Low Check</td></tr>
								<tr><td><input type="checkbox" value="1" name="spike_adu" <%if spike_adu=1 then%> CHECKED <%end if%>></td>
										<td>Spike/Adu</td></tr>
								</table>
						</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
&nbsp;<%if not(isBuildingOff(bldg)) then%>&nbsp;<input type="submit" name="action" value="Save"><%end if%>&nbsp;<input type="button" value="Close" onclick="window.close()">
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
<input type="hidden" name="meterid" value="<%=meterid%>">
</form>
</body>
</html>
