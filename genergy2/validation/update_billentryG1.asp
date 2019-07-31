<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'12/17/2007 N.Ambo made changes to lines 191-193 to include criteria for when "extusage" is not 0; te value for current usage should
'refelect the "used" value from the query unless extusage is 0, in which case it will use the "usedtotal" value
'3/10/2008 N.Ambo made changes to add functionality for showing the full meter reading history; added button "Show All History"
function GetNumericParam(key)
	GetNumericParam = request(key)
	if trim(GetNumericParam)="" or not(isnumeric(GetNumericParam)) then GetNumericParam = 0
	GetNumericParam = cdbl(GetNumericParam)
end function

dim byear, bperiod, meterid, tnum, tname, building, isposted, pid, utilityid, yscroll,viewtype, status
meterid = request("meterid")
building = request("building")
byear = request("byear")
bperiod = request("bperiod")
utilityid = request("utilityid")
tname = request("tname")
tnum = request("tnumber")
tid  = request("tid")
isposted = request("posted")
pid = request("pid")
yscroll = request("yscroll")
viewtype 	= request("t")
status = request("showstatus") 'Added by N.Ambo 3/10/2008 - determine whether full meter reading hsitroy should be shown or not

dim rst1, cnn1, strsql, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(building)

dim rst2 
set rst2 = server.createobject("ADODB.recordset")

dim startdate, enddate, datasource, calcintpeak, coindemand, ratename
rst1.open "SELECT rt.type, b.DateStart, b.DateEnd, datasource, lup.calcintpeak, lup.coincident, lup.coincident_peak FROM billyrperiod b, meters m, tblleasesutilityprices lup, dbo.ratetypes rt WHERE lup.ratetenant=rt.id and lup.leaseutilityid=m.leaseutilityid and b.bldgnum=m.bldgnum and billyear="&byear&" and billperiod="&bperiod&" and m.meterid="&meterid, cnn1
'response.write "SELECT rt.type, b.DateStart, b.DateEnd, datasource, lup.calcintpeak, lup.coincident, lup.coincident_peak FROM billyrperiod b, meters m, tblleasesutilityprices lup, dbo.ratetypes rt WHERE lup.ratetenant=rt.id and lup.leaseutilityid=m.leaseutilityid and b.bldgnum=m.bldgnum and billyear="&byear&" and billperiod="&bperiod&" and m.meterid="&meterid
'response.end
if not rst1.eof then
	if rst1("calcintpeak") then calcintpeak = true else calcintpeak = false end if
	if (rst1("coincident") or rst1("coincident_peak")) then coindemand = true else coindemand = false end if
	startdate = rst1("DateStart")
	enddate = rst1("DateEnd")
	datasource = rst1("datasource")
	ratename = rst1("type")
end if
rst1.close
if request("action")="Save" then
'	dim Prev, Current, Used, PrevOff, CurrentOff, UsedOff, PrevInt, CurrentInt, UsedInt
'	Prev		= request("Prev")
'	Current		= request("Current")
'	Used		= request("Used")
'	PrevOff		= request("PrevOff")
'	CurrentOff	= request("CurrentOff")
'	UsedOff		= request("UsedOff")
'	PrevInt		= request("PrevInt")
'	CurrentInt	= request("CurrentInt")
'	UsedInt		= request("UsedInt")
'	if not(isnumeric(Prev)) or Trim(Prev) = ""  then Prev = 0
'	if not(isnumeric(Current)) or Trim(Current) = ""  then Current = 0
'	if not(isnumeric(Used)) or Trim(Used) = ""  then Used = 0
'	if not(isnumeric(PrevOff)) or Trim(PrevOff) = ""  then PrevOff = 0
'	if not(isnumeric(CurrentOff)) or Trim(CurrentOff) = ""  then CurrentOff = 0
'	if not(isnumeric(UsedOff)) or Trim(UsedOff) = ""  then UsedOff = 0
'	if not(isnumeric(PrevInt)) or Trim(PrevInt) = ""  then PrevInt = 0
'	if not(isnumeric(CurrentInt)) or Trim(CurrentInt) = ""  then CurrentInt = 0
'	if not(isnumeric(UsedInt)) or Trim(UsedInt) = ""  then UsedInt = 0
	if coindemand then
		dim coin_demand_save, coin_datepeak_save
		if not isnumeric(trim(request("coin_demand"))) then coin_demand_save = 0 else coin_demand_save = trim( request("coin_demand") ) end if
		if not isdate(trim(request("coin_datepeak"))) then coin_datepeak_save = 0 else coin_datepeak_save = trim(request("coin_datepeak")) end if
		' update coincident demand
		dim sqlcoin, rstcoin
		set rstcoin = server.createobject("ADODB.recordset")
		sqlcoin = "select leaseutilityid from coincidentdemand where leaseutilityid = '" & request("luid") & "' and billyear=" & byear & " and billperiod="&bperiod
		rstcoin.open sqlcoin, cnn1
		if rstcoin.eof then
			dim tempsql
			tempsql = "insert into coincidentdemand(leaseutilityid, billyear, billperiod, demand, datepeak) values (" & request("luid") & ", " & byear & ", " & bperiod & ", " & coin_demand_save & ", '" & coin_datepeak_save & "')"
			'response.write tempsql
			'response.end
			rst1.open tempsql, cnn1
			'rst1.close
		else
			dim tempsqlupdate
			tempsqlupdate = "update coincidentdemand set demand = "& coin_demand_save & ", datepeak = '"&coin_datepeak_save&"' where leaseutilityid = " & request("luid") & " and billyear = "&byear&" and billperiod = "&bperiod
			rst1.open tempsqlupdate, cnn1
			'rst1.close
		end if
		rstcoin.close
		set rstcoin = nothing
	end if 'the previous only saves info the coin_demand table
	dim cest, pdest
	if request("cest")="1" then cest = 1 else cest = 0
	if request("pdest")="1" then pdest = 1 else pdest = 0
	'stored proc scripting
	cnn1.CursorLocation = adUseClient
	'specify stored procedure to run
	cmd.CommandText = "sp_validation_v4"
	cmd.CommandType = adCmdStoredProc

	Set prm = cmd.CreateParameter("meterid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adSmallInt, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("kwh", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("kw", adDouble, adParamInput)
	cmd.Parameters.Append prm
	set prm = cmd.createParameter("kw_off", adDouble, adParamInput)
	cmd.parameters.append prm
	set prm = cmd.createParameter("kw_int", adDouble, adParamInput)
	cmd.parameters.append prm
	Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("on", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("off", adDouble, adParamInput)
	cmd.Parameters.Append prm
	set prm = cmd.createParameter("int", adDouble, adParamInput)
	cmd.parameters.append prm
	Set prm = cmd.CreateParameter("diff", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("note", adVarChar, adParamInput, 250)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("prevkwh", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("pdnote", adVarChar, adParamInput, 250)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("est_usage", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("est_demand", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("prevkwho", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("kwho", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("diffo", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("prevkwhi", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("kwhi", adDouble, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("diffi", adDouble, adParamInput)
	cmd.Parameters.Append prm
	'set connection
	Set cmd.ActiveConnection = cnn1
	'set input params
	cmd.Parameters("meterid")	= meterid
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.Parameters("kwh")		= GetNumericParam("Current")
	cmd.Parameters("kw")		= GetNumericParam("Demand")
	cmd.Parameters("kw_off")	= GetNumericParam("demand_off")
	cmd.Parameters("kw_int")	= GetNumericParam("demand_int")
	cmd.Parameters("user")		= getKeyValue("user")
	cmd.Parameters("on")		= GetNumericParam("OnPeak")
	cmd.Parameters("off")		= GetNumericParam("OffPeak")
	cmd.parameters("int")		= GetNumericParam("intPeak")
	cmd.Parameters("diff")		= GetNumericParam("Used")
	cmd.Parameters("note")		= left(request("note"),249)
	cmd.Parameters("prevkwh")	= GetNumericParam("Prev")
	cmd.Parameters("pdnote")	= left(request("pdnote"),249)
	cmd.Parameters("est_usage")	= cest
	cmd.Parameters("est_demand")= pdest
	cmd.Parameters("prevkwho")	= GetNumericParam("PrevOff")
	cmd.Parameters("kwho")		= GetNumericParam("CurrentOff")
	cmd.Parameters("diffo")		= GetNumericParam("UsedOff")
	cmd.Parameters("prevkwhi")	= GetNumericParam("PrevInt")
	cmd.Parameters("kwhi")		= GetNumericParam("CurrentInt")
	cmd.Parameters("diffi")		= GetNumericParam("UsedInt")
	
	'response.write "exec sp_validation_v2 " & cmd.Parameters("meterid") & ", " & cmd.Parameters("by") & ", " & cmd.Parameters("bp") & ", " & _
	'	cmd.Parameters("kwh") & ", " & cmd.Parameters("kw") & ", " & cmd.Parameters("kw_off")	& ", " & cmd.Parameters("kw_int") & ", " & _
	'	cmd.Parameters("user") & ", " & cmd.Parameters("on") & ", " & cmd.Parameters("off") & ", " & cmd.parameters("int") & ", " & cmd.Parameters("diff") & ", '" & _
	'	cmd.Parameters("note") & "', " & cmd.Parameters("prevkwh") & ", '" & cmd.Parameters("pdnote") & "'," & cmd.Parameters("est_usage") & ", " & cmd.Parameters("est_demand") & ", " & _
	'	cmd.Parameters("est_usage") & ", " & cmd.Parameters("est_demand") & ", " & cmd.Parameters("prevkwho") & ", " & cmd.Parameters("kwho") & ", " & cmd.Parameters("diffo") & ", " & cmd.Parameters("prevkwhi") & ", " & cmd.Parameters("kwhi") & ", " & cmd.Parameters("diffi")
	'response.end
	cmd.execute
	
	' 03/25/2008 : Added by Tarun , Customer Notes and Internal Notes added to the consumption table
	dim strUpdateNotes, mlcustnotes, mlintnotes
	
	mlcustnotes = left(request("mlcnote"),249)
	mlintnotes = left(request("mlinote"),249)
	if trim(cstr(byear)) <> "" and trim(cstr(bperiod)) <> "" then 
		strUpdateNotes = "Update Consumption Set MLCustomerNotes ='" & mlcustnotes & "', MLInternalNotes = '" & mlintnotes & "' " & _
						 " WHERE meterid = " & meterid & " and billyear =" & byear & " and billperiod = " & bperiod 
						  
		rst2.open strUpdateNotes, cnn1	
	end if
	
	
	%>
	<script>
		var mscroll = opener.mainscrollpoint();
		var yscroll = opener.scrollpoint(0)
		var yscroll2 = opener.scrollpoint(1)
		var showscroll = opener.displaypoint('abovevar')
		var showscroll2 = opener.displaypoint('belowvar')
		opener.document.location.href = "validation_select.asp?t=<%=viewtype%>&pid=<%=pid%>&building=<%=building%>&byear=<%=byear%>&utilityid=<%=utilityid%>&bperiod=<%=bperiod%>&yscroll="+yscroll+"&yscroll2="+yscroll2+"&showscroll="+showscroll+"&showscroll2="+showscroll2+"&mscroll="+mscroll;
		window.close();
	</script>
	<%response.end
end if

'3/10/2008 N.Ambo added to show full meter read history if button "Show All History" is selected
if status = "All" then
	strsql = "SELECT isnull(c.prev,0)+isnull(c.PreviousInt,0)+isnull(c.Previousoff,0) as prevtotal, isnull([current],0)+isnull(currentint,0)+isnull(currentoff,0) as currtotal, isnull(used,0)+isnull(usedint,0)+isnull(usedoff,0) as usedtotal, m.MeterNum, c.BillYear, c.BillPeriod, c.OnPeak, ISNULL(c.IntPeak, 0) AS intpeak, c.OffPeak, c.Prev, PreviousOff, PreviousInt, c.[Current], c.CurrentOff, c.CurrentInt, c.Used, UsedOff, UsedInt, pd.Demand, c.UserNote AS note, m.DateLastRead, pd.UserNote AS pdnote, pd.demand_int, pd.demand_off, tlup.Coincident AS coindemand,cd.Demand AS coin_demand, cd.DatePeak AS coin_datepeak, tlup.leaseutilityid as luid, c.estimated as cest, pd.estimated as pdest, extusg, IsNull(c.MLCustomerNotes, '') as MLCustomerNotes, IsNull(c.MLInternalNotes, '') as MLInternalNotes  FROM Consumption c INNER JOIN Meters m ON m.MeterId = c.MeterId INNER JOIN PeakDemand pd ON pd.MeterId = c.MeterId AND c.BillYear = pd.BillYear AND c.BillPeriod = pd.BillPeriod left OUTER JOIN validation v ON v.meterid = c.MeterId AND v.billperiod = c.BillPeriod AND v.billyear = c.BillYear join tblleasesutilityprices tlup on tlup.leaseutilityid=m.leaseutilityid left outer join coincidentdemand cd on cd.leaseutilityid=tlup.leaseutilityid and cd.billyear=c.billyear and cd.billperiod=c.billperiod WHERE c.MeterId="& MeterId &" and (tlup.leaseutilityid = m.leaseutilityid) ORDER BY c.BillYear DESC, c.BillPeriod DESC"
else
	strsql = "SELECT isnull(c.prev,0)+isnull(c.PreviousInt,0)+isnull(c.Previousoff,0) as prevtotal, isnull([current],0)+isnull(currentint,0)+isnull(currentoff,0) as currtotal, isnull(used,0)+isnull(usedint,0)+isnull(usedoff,0) as usedtotal, m.MeterNum, c.BillYear, c.BillPeriod, c.OnPeak, ISNULL(c.IntPeak, 0) AS intpeak, c.OffPeak, c.Prev, PreviousOff, PreviousInt, c.[Current], c.CurrentOff, c.CurrentInt, c.Used, UsedOff, UsedInt, pd.Demand, c.UserNote AS note, m.DateLastRead, pd.UserNote AS pdnote, pd.demand_int, pd.demand_off, tlup.Coincident AS coindemand,cd.Demand AS coin_demand, cd.DatePeak AS coin_datepeak, tlup.leaseutilityid as luid, c.estimated as cest, pd.estimated as pdest, extusg, IsNull(c.MLCustomerNotes, '') as MLCustomerNotes, IsNull(c.MLInternalNotes, '') as MLInternalNotes  FROM Consumption c INNER JOIN Meters m ON m.MeterId = c.MeterId INNER JOIN PeakDemand pd ON pd.MeterId = c.MeterId AND c.BillYear = pd.BillYear AND c.BillPeriod = pd.BillPeriod left OUTER JOIN validation v ON v.meterid = c.MeterId AND v.billperiod = c.BillPeriod AND v.billyear = c.BillYear join tblleasesutilityprices tlup on tlup.leaseutilityid=m.leaseutilityid left outer join coincidentdemand cd on cd.leaseutilityid=tlup.leaseutilityid and cd.billyear=c.billyear and cd.billperiod=c.billperiod WHERE c.MeterId="& MeterId &" and ((c.BillYear="& byear &" and c.BillPeriod<="& bperiod &") or (c.BillYear="& byear-1 &" and c.BillPeriod>="& bperiod &")) AND (tlup.leaseutilityid = m.leaseutilityid) ORDER BY c.BillYear DESC, c.BillPeriod DESC"
end if

'response.write strsql
'response.end
rst1.open strsql, cnn1

%>
<html>
<head><title>Bill Validation</title>
<script>
function makesubmit()
{	var frm = document.forms['form1']
	if(frm.Prev.value!=frm.PrevOrig.value)
	{	if(frm.note.value=='')
		{	alert('You must enter a note when changing the previous KWH value');
			return(0);
		}
	}
	
	frm.submit();
}

function ExtTotal(part){
	var frm=document.forms[0];
	try{
		document.all[part+'Total'].innerHTML=ParseFloatExt(frm[part].value)+ParseFloatExt(frm[part+'Off'].value)+ParseFloatExt(frm[part+'Int'].value);
	}catch(exception){}
}

function ParseFloatExt(num){
	var output;
	output = (parseFloat(num));
	if(isNaN(output)) output=0;
	return(output);
}
</script>
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #dddddd; }

INPUT {
	text-align : right;
}

.bordercell { border-left:1px solid #ffffff; border-bottom:1px solid #ffffff; border-top:1px solid #ffffff; }
.bordercolumn { border-left:1px solid #eeeeee; }
</style>
</head>
<body leftmargin="0" topmargin="0">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
	<tr>
		<td bgcolor="#6699cc"><span class="standardheader">
			<%=tnum%> &nbsp;<%=tname%>&nbsp;&nbsp;Meter: <%=rst1("meternum")%>, Date Last Read: <%=breakDate(rst1("datelastread")) %> Rate: <%=ratename%></span>
		</td>
		<td bgcolor="#6699cc" align="right">&nbsp;&nbsp;
			<%
			dim rsLink, luid, tid
			set rsLink = server.createobject("adodb.recordset")
			rsLink.open "select leaseutilityid from meters where meterid = '"&request("meterid")&"'", getLocalConnect(request("building"))
			if not rsLink.eof then luid = rsLink("leaseutilityid")
			rsLink.close
			set rsLink = nothing
			
			dim meterSetupLink, manualURL
			meterSetupLink = "window.open('/genergy2/setup/contentfrm.asp?action=meteredit&pid="&request("pid")&"&bldg="&request("building")&_
				"&tid="&tid&"&lid="&luid&"&meterid="&request("meterid")&"','MeterSetup','width=900,height=525,resizable=yes,toolbar=no,scrollbars=yes')"	%>
		<%
			
			if trim(datasource) = ""  OR ISNULL(datasource) then		
			
				manualURL = "window.opener.document.location = '/genergy2/manualentry/entry_select.asp?pid="&pid&"&building="&building&"&byear="&byear'&"&bperiod="&bperiod				
				if isPosted<>"True" then manualURL = manualURL & "&bperiod="&bperiod
				manualURL = manualURL&"&meterid="&meterid&"';window.close();"
				%>
				<a class="standardheader" href="#" onClick="<%=manualURL%>">Manual Entry</a>&nbsp;
				<%						
			else
				dim LMPURL, pulseURL
				LMPURL = "window.open('/genergy2/eri_th/lmp/lmp.asp?meterid="&meterid&"&bldg="&building&"&utility="&utilityid&_
					"&indiWindow=true','LMPPopup','width=750,height=600,toolbar=no')"
				
				pulseURL = "/genergy2/UMreports/meterPulseReport.asp?meterid="&meterid&"&startdate="&startdate&"&enddate="&enddate&"&bldg="&building
				%>
														
				<%
			end if%>
		</td>
	</tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="3">
	<tr bgcolor="#dddddd">
		<td class="bordercell">Year</td>
		<td class="bordercell">Period</td>
		<%if not rst1("extusg") then%>
			<td class="bordercell">On&nbsp;Peak</td>
			<%if calcintPeak or rst1("extusg")then%><td class="bordercell">Int&nbsp;Peak</td><%end if%>
			<td class="bordercell">Off&nbsp;Peak</td>
		<%end if%>
		<td class="bordercell">Prev Reading</td>
		<td class="bordercell">Current Reading</td>
		<td class="bordercell">Current Usage</td>
		<td class="bordercell">Demand<%if calcintPeak or rst1("extusg") then%>&nbsp;On<%end if%></td>
		<%if calcintPeak or rst1("extusg") then%>
			<td class="bordercell">Demand&nbsp;Int</td>
			<td class="bordercell">Demand&nbsp;Off</td>
		<%end if%>
		<%if coindemand then%>
			<td class="bordercell">Coin. Demand <font color="red" style="bold">*</font></td>
			<td class="bordercell">Coin. Date of Peak <font color="red" style="bold">*</font></td>
		<%end if%>
		<td class="bordercell">Estimated Usage</td>
		<td class="bordercell">Estimated Demand</td>
		<td class="bordercell"></td>
		<td class="bordercell"</td>
		<td class="bordercell"></td>
		<td class="bordercell"></td>
		<td class="bordercell"</td>
	</tr>
	<%
	dim currentbp
	if not rst1.EOF then
		'response.write "luid:" & rst1("luid")
		currentbp = trim(rst1("billperiod"))
		%>
		<form name="form1" method="post">
		<tr valign="top" bgcolor="#eeeeee">
			<td><%=rst1("BillYear")%><input type="hidden" name="meterid" value="<%=meterid%>"><input type="hidden" name="byear" value="<%=byear%>"><input type="hidden" name="bperiod" value="<%=bperiod%>"></td>
			<td><%=rst1("BillPeriod")%></td>
			<%if isposted<>"True" and not(isbuildingOff(building)) then%>
				<%if not rst1("extusg") then%>
					<td><input type="text" name="OnPeak" value="<%=rst1("OnPeak")%>" size="5"></td>
					<%if calcintPeak or rst1("extusg") then%>
						<td><input type="text" name="IntPeak" value="<%=rst1("intPeak")%>" size="5"></td>
					<%else%>
						<input type="hidden" name="IntPeak" value=0>	
					<%end if%>
					<td><input type="text" name="OffPeak" value="<%=rst1("OffPeak")%>" size="5"></td>
				<%end if%>
				<td><table cellspacing="0" cellpadding="0">
						<tr>
						<%if rst1("extusg") then response.write "<td>On</td>"%>
						<td><input type="hidden" name="PrevOrig" value="<%=rst1("Prev")%>"><input type="text" name="Prev" value="<%=rst1("Prev")%>" size="5" <%if rst1("extusg") then%>onKeyUp="ExtTotal('Used')"<%end if%>></td>
						<%if rst1("extusg") then%>
						<tr><td>Off</td><td><input type="text" name="PrevOff" value="<%=rst1("PreviousOff")%>" size="5" onKeyUp="ExtTotal('Used')"></td></tr>
						<tr><td>Int</td><td><input type="text" name="PrevInt" value="<%=rst1("PreviousInt")%>" size="5" onKeyUp="ExtTotal('Used')"></td></tr>
						<tr><td></td><td id="PrevTotal">&nbsp;</td></tr>
						<%end if%>
				</table></td>
				<td><table cellspacing="0" cellpadding="0">
						<tr>
						<%if rst1("extusg") then response.write "<td>On</td>"%>
						<td><input type="text" name="Current" value="<%=rst1("Current")%>" size="5" onKeyUp="Used.value=this.value-Prev.value; <%if rst1("extusg") then%>ExtTotal('Used')<%end if%>"></td></tr>
						<%if rst1("extusg") then%>
						<tr><td>Off</td><td><input type="text" name="CurrentOff" value="<%=rst1("CurrentOff")%>" size="5" onKeyUp="UsedOff.value=this.value-PrevOff.value; ExtTotal('Used')"></td></tr>
						<tr><td>Int</td><td><input type="text" name="CurrentInt" value="<%=rst1("CurrentInt")%>" size="5" onKeyUp="UsedInt.value=this.value-PrevInt.value; ExtTotal('Used')"></td></tr>
						<tr><td></td><td id="CurrentTotal">&nbsp;</td></tr>
						<%end if%>
				</table></td>
				<td><table cellspacing="0" cellpadding="0">
						<tr>
						<%if rst1("extusg") then response.write "<td>On</td>"%>
						<td><input type="text" name="Used" readonly value="<%=rst1("Used")%>" size="5" <%if rst1("extusg") then%>onKeyUp="ExtTotal('Used')"<%end if%>></td></tr>
						<%if rst1("extusg") then%>
						<tr><td>Off</td><td><input type="text" name="UsedOff" readonly value="<%=rst1("UsedOff")%>" size="5" onKeyUp="ExtTotal('Used')"></td></tr>
						<tr><td>Int</td><td><input type="text" name="UsedInt" readonly value="<%=rst1("UsedInt")%>" size="5" onKeyUp="ExtTotal('Used')"></td></tr>
						<tr><td>Total&nbsp;</td><td align="right" id="UsedTotal">&nbsp;</td></tr>
						<%end if%>
				</table></td>
				<td><input type="text" name="demand" value="<%=rst1("Demand")%>" size="5"></td>
				<%if calcintPeak or rst1("extusg") then%>
					<td><input type="text" name="demand_int" value="<%=rst1("Demand_int")%>" size="5"></td>
					<td><input type="text" name="demand_off" value="<%=rst1("Demand_off")%>" size="5"></td>
				<%else%>
					<input type="hidden" name="demand_int" value=0>
					<input type="hidden" name="demand_off" value=0>
				<%end if%>
				<%if coindemand then%>
					<td><input type="text" name="coin_demand" value="<%=rst1("coin_demand")%>" size="5"></td>
					<td><textarea cols="15" rows="" name="coin_datepeak"><%=rst1("coin_datepeak")%></textarea></td>
				<%end if%>
				<td><input type="checkbox" name="cest" value="1" <%if rst1("cest")="True" then response.write "CHECKED"%>></td>
				<td><input type="checkbox" name="pdest" value="1" <%if rst1("pdest")="True" then response.write "CHECKED"%>></td>
				<%origLink%>
				<%=rst1("note")%>
				<%=rst1("pdnote")%>
				<%=rst1("MLCustomernotes")%>
				<%=rst1("MLInternalnotes")%>				
				</tr>
				<tr bgcolor="#eeeeee">
				<td colspan="<%if calcintPeak or rst1("extusg") then%>18<%else%>15<%end if%>" style="border-bottom:1px solid #cccccc;">
				<%if not(isBuildingOff(building)) then%>
					<input type="button" onClick="makesubmit()" name="Action2" value="Save" size="15" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
					<input type="hidden" name="Action" value="Save">
					<%if coindemand then%>
						<font color="red" style="bold">*</font>  :   This value will be changed for the tenant that this meter is in; therefore all 		meters for this tenant will reflect this change.
					<%end if%>
				<%end if%>
				</td>
			<%else%>
				<%if not rst1("extusg") then%>
					<td align="right"><%=rst1("OnPeak")%></td>
					<%if calcintPeak or rst1("extusg") then%><td align="right"><%=rst1("intPeak")%></td><%end if%>
					<td align="right"><%=rst1("OffPeak")%></td>
				<%end if%>
				<td align="right" class="bordercolumn"><%if not rst1("extusg") then%><%=rst1("Prev")%><%else%><%=rst1("PrevTotal")%><%end if%>&nbsp;</td>
				<td align="right" class="bordercolumn"><%if not rst1("extusg") then%><%=rst1("Current")%><%else%><%=rst1("CurrTotal")%><%end if%>&nbsp;</td>
				<td align="right" class="bordercolumn"><%if not rst1("extusg") then%><%=rst1("Used")%><%else%><%=rst1("UsedTotal")%><%end if%>&nbsp;</td>
				<td align="right"><%=rst1("Demand")%></td>
				<%if calcintPeak or rst1("extusg") then%>
					<td align="right"><%=rst1("Demand_int")%></td>
					<td align="right"><%=rst1("Demand_off")%></td>
				<%end if%>
				<%if coindemand then%>
					<td align="right"><%=rst1("coin_demand")%></td>
					<td align="right"><%=rst1("coin_datepeak")%></td>
				<%end if%>
				<td ><%=rst1("cest")%>&nbsp;</td>
				<td ><%=rst1("pdest")%>&nbsp;</td>
				<%origLink%>
				<td><%=rst1("note")%></td>
				<td><%=rst1("pdnote")%></td>
				<td><%=rst1("MLCustomernotes")%></td>
				<td><%=rst1("MLInternalnotes")%></td>				
			<%end if%>
	
			<input type="hidden" name="building" value="<%=building%>">
			<input type="hidden" name="pid" value="<%=pid%>">
			<input type="hidden" name="utilityid" value="<%=utilityid%>">
			<input type="hidden" name="luid" value="<%=rst1("luid")%>">
		</tr>
		</form>
		<%rst1.movenext
	end if%>
	<%
	dim PrevTotal, CurrentTotal, UsedTotal
	do until rst1.EOF
		if trim(rst1("billperiod"))<>currentbp then
			currentbp = trim(rst1("billperiod"))%>
			<tr valign="top" bgcolor="#ffffff">
				<td class="bordercolumn"><%=rst1("BillYear")%>&nbsp;</td>
				<td class="bordercolumn"><%=rst1("BillPeriod")%>&nbsp;</td>
				<%if not rst1("extusg") then%>
					<td align="right" class="bordercolumn"><%=rst1("OnPeak")%>&nbsp;</td>
					<%if calcintpeak or rst1("extusg") then %>
						<td align="right" class="bordercolumn"><%=rst1("IntPeak")%>&nbsp;</td>
					<%end if%>
					<td align="right" class="bordercolumn"><%=rst1("OffPeak")%>&nbsp;</td>
				<%end if%>
				<td align="right" class="bordercolumn"><%if not rst1("extusg") then%><%=rst1("Prev")%><%else%><%=rst1("PrevTotal")%><%end if%>&nbsp;</td>
				<td align="right" class="bordercolumn"><%if not rst1("extusg") then%><%=rst1("Current")%><%else%><%=rst1("CurrTotal")%><%end if%>&nbsp;</td>
				<td align="right" class="bordercolumn"><%if not rst1("extusg") then%><%=rst1("Used")%><%else%><%=rst1("UsedTotal")%><%end if%>&nbsp;</td>
				<td align="right" class="bordercolumn"><%=rst1("Demand")%>&nbsp;</td>
				<%if calcintpeak or rst1("extusg") then%>
					<td align="right" class="bordercolumn"><%=rst1("Demand_int")%>&nbsp;</td>
					<td align="right" class="bordercolumn"><%=rst1("Demand_off")%>&nbsp;</td>
				<%end if%> 
				<%if coindemand then%>
					<td align="right" class="bordercolumn"><%=rst1("coin_demand")%>&nbsp;</td>
					<td align="right" class="bordercolumn"><%=rst1("coin_datepeak")%>&nbsp;</td>
				<%end if%>
				<td class="bordercolumn"><%=rst1("cest")%>&nbsp;</td>
				<td class="bordercolumn"><%=rst1("pdest")%>&nbsp;</td>
				<%origLink%>
				<td class="bordercolumn"><%=rst1("note")%>&nbsp;</td>
				<td class="bordercolumn"><%=rst1("pdnote")%>&nbsp;</td>
				<td class="bordercolumn"><%=rst1("MLCustomernotes")%>&nbsp;</td>
				<td class="bordercolumn"><%=rst1("MLInternalnotes")%>&nbsp;</td>				
			</tr>
		<%end if
		rst1.movenext
	loop
	%>
</table>
<table>
<tr>&nbsp;
<td colspan="3" align="right"><input type="button" class="standard" style="cursor:hand;background-color:#dddddd;border:1px outset #ffffff;color:336699;" value="Show All History" onclick="open('/genergy2/validation/update_billentryG1.asp?meterid=<%=meterid%>&byear=<%=byear%>&bperiod=<%=bperiod%>&tname=<%=tname%>&tnumber=&building=<%=building%>&pid=<%=pid%>&utilityid=<%=utilityid%>&showstatus=All&posted=True', 'update_billentry','left=8,top=8,scrollbars=yes,width=770, height=380, status=no')" ID="Button1" NAME="Button1"></td>
</tr>
</table>
</body>
<script>
ExtTotal('Used');
</script>
</html>
<%
Function breakDate(strng)
	dim RegularExpressionObject, ReplacedString, RetStr
	Set RegularExpressionObject = New RegExp
	
	With RegularExpressionObject
	.Pattern = " "
	.IgnoreCase = True
	.Global = False
	End With
	if not isnull(rst1("datelastread")) then 
		ReplacedString = RegularExpressionObject.Replace(rst1("datelastread"), " <br>")
		RetStr = ReplacedString
		Set RegularExpressionObject = nothing
		
		breakDate = RetStr
	end if
End Function
%>
<%sub origLink()%>
	<td><a href="#" onClick="open('original_meterreadings.asp?meterid=<%=meterid%>&bperiod=<%=rst1("BillPeriod")%>&byear=<%=rst1("BillYear")%>&building=<%=building%>','','width=210,height=230')">see&nbsp;orig.&nbsp;read</a></td>
<%end sub%>
