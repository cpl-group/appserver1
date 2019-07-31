<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim building, byear, bperiod, prevYear, prevPeriod, checking
building = request.querystring("building")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
checking = trim(request.querystring("checking"))
prevYear = byear
prevPeriod = bperiod-1
if prevPeriod<1 then 
	prevYear = byear-1
	prevPeriod = 12
end if

dim rst1, cnn1, strsql, cnn2
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy1")

dim CurrentOutput, prevOutput
CurrentOutput = ""
prevOutput = ""

if checking="Tenant" then
	rst1.open "SELECT Distinct l.BillingName as onlyCurrent FROM tblMetersByPeriod mbp INNER JOIN tblLeasesUtilityPrices lup ON lup.LeaseUtilityId=mbp.LeaseUtilityId INNER JOIN tblLeases l ON l.billingid=lup.billingid WHERE mbp.bldgnum='"&building&"' and billperiod="&bperiod&" and billyear="&byear&" and mbp.Leaseutilityid not in (SELECT Distinct LeaseUtilityId from tblMetersByPeriod WHERE bldgnum='"&building&"' and billperiod="&prevPeriod&" and billyear="&prevYear&")", cnn1
	do until rst1.eof
		CurrentOutput = CurrentOutput & rst1("onlyCurrent")&"<BR>"
		rst1.movenext
	loop
	rst1.close
	rst1.open "SELECT Distinct l.BillingName as onlyPrev FROM tblMetersByPeriod mbp INNER JOIN tblLeasesUtilityPrices lup ON lup.LeaseUtilityId=mbp.LeaseUtilityId INNER JOIN tblLeases l ON l.billingid=lup.billingid WHERE mbp.bldgnum='"&building&"' and billperiod="&prevPeriod&" and billyear="&prevYear&" and mbp.Leaseutilityid not in (SELECT Distinct LeaseUtilityId from tblMetersByPeriod WHERE bldgnum='"&building&"' and billperiod="&bperiod&" and billyear="&byear&")", cnn1
	do until rst1.eof
		prevOutput = prevOutput & rst1("onlyPrev")&"<BR>"
		rst1.movenext
	loop
	rst1.close
else
	rst1.open "SELECT c.meternum as onlyCurrent FROM tblmetersbyperiod c WHERE bldgnum='"&building&"' and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and c.meternum not in (SELECT c.meternum FROM tblmetersbyperiod c WHERE bldgnum='"&building&"' and c.BillYear="&prevYear&" and c.BillPeriod="&prevPeriod&")", cnn1
	do until rst1.eof
		CurrentOutput = CurrentOutput & rst1("onlyCurrent")&"<BR>"
		rst1.movenext
	loop
	rst1.close
	rst1.open "SELECT c.meternum as onlyPrev FROM tblmetersbyperiod c WHERE bldgnum='"&building&"' and c.BillYear="&prevYear&" and c.BillPeriod="&prevPeriod&" and c.meternum not in (SELECT c.meternum FROM tblmetersbyperiod c WHERE bldgnum='"&building&"' and c.BillYear="&byear&" and c.BillPeriod="&bperiod&")", cnn1
	do until rst1.eof
		prevOutput = prevOutput & rst1("onlyPrev")&"<BR>"
		rst1.movenext
	loop
	rst1.close
end if

%>
<html>
<head>
<title><%=checking%> Differences</title>
</head>
<body>
<span style="font-family:arial, helvetica, san-serif; font-size:12px;">
<%
if trim(currentOutput)<>"" then
	response.write "<b>"&checking&"s not in Previous Period</b><br>"
	response.write currentOutput
end if
if trim(prevOutput)<>"" then
	response.write "<b>"&checking&"s not in Current Period</b><br>"
	response.write prevOutput
end if
if trim(currentOutput)="" and trim(prevOutput)="" then
	response.write "<b>No differences between current and last period "&lcase(checking)&"s.</b>"
end if
%>
</span>
</body>
</html>
