<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim building, byear, bperiod, prevYear, prevPeriod, checking, utilityid
building = request.querystring("building")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
utilityid = request("utilityid")
checking = trim(request.querystring("checking"))
prevYear = byear
prevPeriod = bperiod-1
if prevPeriod<1 then 
	prevYear = byear-1
	prevPeriod = 12
end if
'response.write("tml.building:"+building)
'response.end
dim rst1, cnn1, strsql, cnn2
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(building)

dim CurrentOutput, prevOutput
CurrentOutput = ""
prevOutput = ""

if checking="Tenant" then
	rst1.open "SELECT Distinct l.BillingName as onlyCurrent, l.tenantnum FROM tblMetersByPeriod mbp INNER JOIN tblbillbyperiod b ON b.id=mbp.bill_id INNER JOIN tblLeasesUtilityPrices lup ON lup.LeaseUtilityId=mbp.LeaseUtilityId INNER JOIN tblLeases l ON l.billingid=lup.billingid WHERE mbp.bldgnum='"&building&"' and b.billperiod="&bperiod&" and b.billyear="&byear&" and mbp.Leaseutilityid not in (SELECT Distinct LeaseUtilityId from tblMetersByPeriod WHERE bldgnum='"&building&"' and billperiod="&prevPeriod&" and billyear="&prevYear&") and b.reject=0 and lup.utility="&utilityid, cnn1
	do until rst1.eof
		CurrentOutput = CurrentOutput & rst1("onlyCurrent")&"<BR>"
		rst1.movenext
	loop
	rst1.close
	rst1.open "SELECT Distinct l.BillingName as onlyPrev FROM tblMetersByPeriod mbp INNER JOIN tblbillbyperiod b ON b.id=mbp.bill_id INNER JOIN tblLeasesUtilityPrices lup ON lup.LeaseUtilityId=mbp.LeaseUtilityId INNER JOIN tblLeases l ON l.billingid=lup.billingid WHERE mbp.bldgnum='"&building&"' and mbp.billperiod="&prevPeriod&" and mbp.billyear="&prevYear&" and mbp.Leaseutilityid not in (SELECT Distinct LeaseUtilityId from tblMetersByPeriod WHERE bldgnum='"&building&"' and billperiod="&bperiod&" and billyear="&byear&") and b.reject=0 and lup.utility="&utilityid, cnn1
	do until rst1.eof
		prevOutput = prevOutput & rst1("onlyPrev")&"<BR>"
		rst1.movenext
	loop
	rst1.close
else
	rst1.open "SELECT c.meternum as onlyCurrent FROM tblmetersbyperiod c INNER JOIN tblbillbyperiod b ON b.id=c.bill_id INNER JOIN tblLeasesUtilityPrices lup ON lup.LeaseUtilityId=c.LeaseUtilityId WHERE c.bldgnum='"&building&"' and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and c.meternum not in (SELECT c.meternum FROM tblmetersbyperiod c WHERE bldgnum='"&building&"' and c.BillYear="&prevYear&" and c.BillPeriod="&prevPeriod&") and b.reject=0 and lup.utility="&utilityid, cnn1
	do until rst1.eof
		CurrentOutput = CurrentOutput & rst1("onlyCurrent")&"<BR>"
		rst1.movenext
	loop
	rst1.close
	rst1.open "SELECT c.meternum as onlyPrev FROM tblmetersbyperiod c INNER JOIN tblbillbyperiod b ON b.id=c.bill_id INNER JOIN tblLeasesUtilityPrices lup ON lup.LeaseUtilityId=c.LeaseUtilityId WHERE c.bldgnum='"&building&"' and c.BillYear="&prevYear&" and c.BillPeriod="&prevPeriod&" and c.meternum not in (SELECT c.meternum FROM tblmetersbyperiod c WHERE bldgnum='"&building&"' and c.BillYear="&byear&" and c.BillPeriod="&bperiod&") and b.reject=0 and lup.utility="&utilityid, cnn1
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
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
</head>
<body>
<div style="margin:10px;">
<%
if trim(currentOutput)<>"" then
	response.write "<b>"&checking&"s not in previous period</b><br>"
	response.write currentOutput
end if
if trim(prevOutput)<>"" then
	response.write "<b>"&checking&"s not in current period</b><br>"
	response.write prevOutput
end if
if trim(currentOutput)="" and trim(prevOutput)="" then
	response.write "<b>No differences between current and last period "&lcase(checking)&"s.</b>"
end if
%>
</div>
</body>
</html>
