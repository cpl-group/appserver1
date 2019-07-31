<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim bldg, pid, bldgName, action, ypid, pypid, utype, bperiod, byear, utilityid, startdate, enddate

bldg = request("bldg")
ypid = securerequest("ypid")
pypid = securerequest("pypid")
luid = securerequest("luid")
action = securerequest("action")
pid = securerequest("pid")
startdate = securerequest("startdate")
enddate = securerequest("enddate")

dim cnn1, cnnMainModule, strsql, rst1, cmd, luid, prm
'set cnnMainModule = server.createobject("ADODB.connection")
'cnnMainModule.open application("cnnstr_SuperMod")
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

''submit section
cmd.activeConnection = cnn1
if action="Add" then
	cmd.CommandText = "INSERT INTO Billyrperiod_partial (ypid, lid, datestart, dateend) Values ("&ypid&", "&luid&", '"&startdate&"', '"&enddate&"')"
elseif action="Edit" then
	cmd.CommandText = "UPDATE Billyrperiod_partial SET datestart='"&startdate&"', dateend='"&enddate&"' WHERE id="&pypid
elseif action="Delete" then
	cmd.CommandText = "Delete Billyrperiod_partial WHERE id="&pypid
end if
'response.write cmd.CommandText
'response.end
if cmd.CommandText<>"" then 
	cmd.execute
	'cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "sp_master_pulse_dump_partial"
	'cmd.commandTimeout = 1800
	Set prm = cmd.CreateParameter("luid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("ypid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	cmd.parameters("luid") = luid
	cmd.parameters("ypid") = ypid
	cmd.execute
	%>
	<script>
	opener.location.reload();
	window.close();
	</script>
	<%
	response.end
end if

strsql = "select b.bldgname, bp.utility, u.utilitydisplay, datestart, dateend, billperiod, billyear FROM buildings b, billyrperiod bp,tblutility u WHERE u.utilityid=bp.utility and ypid="&ypid&" and b.bldgnum='"&bldg&"'"
rst1.open strsql, cnn1
if not rst1.eof then
	bldgName = rst1("bldgname")
	startdate = rst1("datestart")
	enddate = rst1("dateend")
	utilityid = rst1("utility")
	byear = rst1("billyear")
	bperiod = rst1("billperiod")
	utype = rst1("utilitydisplay")
end if
rst1.close

if pypid <> "" then
	strsql = "SELECT lid, datestart, dateend FROM Billyrperiod_partial WHERE id="&pypid
	rst1.open strsql, cnn1
	if not rst1.eof then
		startdate = rst1("datestart")
		enddate = rst1("dateend")
		luid = rst1("lid")
	end if
	rst1.close
end if


%>	
<link rel="Stylesheet" href="setup.css" type="text/css">
<title>Partial Bill Setup</title>
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 onload="window.focus()">
<form action="PartialBill.asp" method="post">
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#3399cc"><td><font color='white'>Partial Bill Setup: <strong><%=bldgname%>&nbsp;<%=bperiod%>,&nbsp;<%=byear%>&nbsp;(<%=utype%>)</strong></font></td></tr>
</table>
<table width="100%" cellspacing="0" cellpadding="2">
<tr><td>Start Date</td><td><input type="Text" name="startdate" size="10" value="<%=startdate%>"></td></tr>
<tr><td>End Date</td><td><input type="Text" name="enddate" size="10" value="<%=enddate%>"></td></tr>
<tr><td valign="top">Lease&nbsp;Utility</td><td>
	<%if luid="" then
		strsql = "SELECT lup.LeaseUtilityId, TenantNum, left(BillingName,23) as BillingName FROM tblleasesutilityprices lup, tblleases l WHERE l.leaseexpired=0 and lup.billingid=l.billingid and leaseutilityid not in (SELECT lid FROM Billyrperiod_partial WHERE ypid="&ypid&") and l.bldgnum='"&bldg&"' and lup.utility="&utilityid&" ORDER BY billingname, tenantnum"
		rst1.open strsql, cnn1
		if not rst1.eof then
			%><select name="luid"><%
			do until rst1.eof
				%><option value="<%=rst1("leaseutilityid")%>"><%=rst1("billingname")%> (<%=rst1("tenantnum")%>)</option><%
				rst1.movenext
			loop
			%></select><%
		else
			%>All tenants for this bill period have partial bills specified.<%
			luid="0"
		end if
		rst1.close
	else
		strsql = "SELECT lup.LeaseUtilityId, TenantNum, left(BillingName,23) as BillingName FROM tblleasesutilityprices lup, tblleases l WHERE l.leaseexpired=0 and lup.billingid=l.billingid and lup.leaseutilityid="&luid
		rst1.open strsql, cnn1
		if not rst1.eof then
		%><%=rst1("billingname")%> (<%=rst1("tenantnum")%>)<%
		end if
		rst1.close
	end if%>
</td></tr>
<tr><td colspan="2" align="center"><%if luid<>"0" then%><input name="action" type="submit" value="<%if pypid="" then%>Add<%else%>Edit<%end if%>">&nbsp;<%if pypid<>"" then%><input name="action" type="submit" value="Delete">&nbsp;<%end if%><%end if%><input name="" type="button" value="Cancel" onclick="window.close()"></td></tr>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="ypid" value="<%=ypid%>">
<%if luid<>"" then%><input type="hidden" name="luid" value="<%=luid%>"><%end if%>
<input type="hidden" name="pypid" value="<%=pypid%>">
</form>
</body>
