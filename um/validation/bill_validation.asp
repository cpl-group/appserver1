<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#include file="checksession.asp"-->
<%
dim super, btabcolor, stabcolor, flagColor, byear, bperiod, building, flagColorHilight, procPage, pid, yscroll
building = request.querystring("building")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
pid = request.querystring("pid")
yscroll = request("yscroll")

dim rst1, cnn1, strsql, cnn2
set cnn1 = server.createobject("ADODB.connection")
set cnn2 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy1")
cnn2.open application("cnnstr_main")

rst1.open "SELECT Supervisor from employees WHERE substring(username,7,90)='"&session("login")&"'", cnn2
if not rst1.eof then
	if rst1("Supervisor")="True" then
		super=true
	else
		super=false
	end if
end if
rst1.close
session("isSuper") = super
if super then
	'Supervisor
	stabcolor="#0099FF"
	btabcolor="#CCCCCC"
	flagColor = "#99FF66"
	flagColorHilight = "#CCFF99"
	procPage = "super_process.asp"
	strsql = "SELECT distinct m.meterid, m.meternum, v.revdate, c.validate, c.svalidate, bbp.posted, m.bldgnum,c.meterid, c.currentkwh, isNull(c.kwhused,0) as kwhused, isNull(pd.demand,0) as demand, l.tenantnum, l.billingname, bbp.totalamt, bbp.adminfee, bbp.sqft, pd.demand, v.biller, v.org_kwh, v.org_kw, "&_
	"case when bbp.sqft=0 then 0 else(bbp.demand/bbp.sqft)end as wsqft, "&_
	"isNuLL((SELECT avg(kwhused) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3))),0) as avgKWH, "&_
	"isNuLL((SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3))),0) as avgKW, "&_
	"isNuLL((SELECT avg(totalamt) FROM tblbillbyperiod bbp2 WHERE bbp2.leaseutilityid=bbp.leaseutilityid and ((bbp2.billyear="&byear-1&" and bbp2.billperiod>="&bperiod&"+9)or(bbp2.billyear="&byear&" and bbp2.billperiod<"&bperiod&" and bbp2.billperiod>="&bperiod&"-3))),0) as avgAmt, "&_
	"isNull(case when (SELECT avg(kwhused) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((c.kwhused - (SELECT avg(kwhused) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3))))/(SELECT avg(kwhused) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3)))*100) end, 0) as kwhvarience, "&_
	"isNull(case when (SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((pd.demand - (SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3))))/(SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)))*100) end, 0) as kwvarience, "&_
	"isNull(case when (SELECT avg(totalamt) FROM tblbillbyperiod bbp2 WHERE bbp2.leaseutilityid=bbp.leaseutilityid and ((bbp2.billyear="&byear-1&" and bbp2.billperiod>="&bperiod&"+9)or(bbp2.billyear="&byear&" and bbp2.billperiod<"&bperiod&" and bbp2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((bbp.totalamt - (SELECT avg(totalamt) FROM tblbillbyperiod bbp2 WHERE bbp2.leaseutilityid=bbp.leaseutilityid and ((bbp2.billyear="&byear-1&" and bbp2.billperiod>="&bperiod&"+9)or(bbp2.billyear="&byear&" and bbp2.billperiod<"&bperiod&" and bbp2.billperiod>="&bperiod&"-3))))/(SELECT avg(totalamt) FROM tblbillbyperiod bbp2 WHERE bbp2.leaseutilityid=bbp.leaseutilityid and ((bbp2.billyear="&byear-1&" and bbp2.billperiod>="&bperiod&"+9)or(bbp2.billyear="&byear&" and bbp2.billperiod<"&bperiod&" and bbp2.billperiod>="&bperiod&"-3)))*100) end, 0) as Amtvarience "&_
	"FROM consumption c  "&_
	"INNER JOIN meters m ON m.Meterid=c.Meterid  "&_
	"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
	"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
	"INNER JOIN tblleases l ON lup.billingid=l.billingid "&_
	"INNER JOIN tblbillbyperiod bbp on lup.leaseutilityid=bbp.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod "&_
	"LEFT JOIN validation v on m.Meterid=v.Meterid and c.billyear=v.billyear and c.billperiod=v.billperiod "&_
	"WHERE c.billyear="&byear&" and c.billperiod="&bperiod&" and m.bldgnum='"&building&"' and m.pp<>1 and Online='1' and leaseexpired=0 "&_
	"ORDER BY l.billingname, m.meternum, v.revdate desc"
else 'and BillPeriod not in (SELECT billperiod FROM tblbillbyperiod WHERE bldgnum='"&building&"' and billyear ="&byear&" and posted=1 group by billyear, billperiod)
	'Biller
	stabcolor="#CCCCCC"
	btabcolor="#0099FF"
	flagColor = "#FF0000"
	flagColorHilight = "#FF3333"
	procPage = "biller_process.asp"
	strsql = "SELECT Distinct m.meternum, c.validate, bbp.posted, c.svalidate, m.bldgnum,c.meterid, c.currentkwh, isNull(c.kwhused,0) as kwhused, isNull(pd.demand,0) as demand, l.tenantnum, isnull(l.billingname,'') as billingname, "&_
	"isNuLL((SELECT avg(kwhused) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3))),0) as avgKWH, "&_
	"isNuLL((SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3))),0) as avgKW, "&_
	"isNull(case when (SELECT avg(kwhused) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((c.kwhused - (SELECT avg(kwhused) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3))))/(SELECT avg(kwhused) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3)))*100) end, 0) as kwhvarience, "&_
	"isNull(case when (SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((pd.demand - (SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3))))/(SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)))*100) end, 0) as kwvarience "&_
	"FROM consumption c  "&_
	"INNER JOIN meters m ON m.Meterid=c.Meterid  "&_
	"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
	"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
	"INNER JOIN tblleases l ON lup.billingid=l.billingid "&_
	"LEFT JOIN tblbillbyperiod bbp on lup.leaseutilityid=bbp.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod "&_
	"WHERE c.billyear="&byear&" and c.billperiod="&bperiod&" and m.bldgnum='"&building&"' and m.pp<>1 and Online='1' and leaseexpired=0 "&_
	"ORDER BY l.billingname, m.meternum"
end if

rst1.open strsql, cnn1
'response.write strsql
'response.end
dim displaydate, perioddates ' get start and end dates of current period for header display
set perioddates = server.createobject("ADODB.recordset")
perioddates.open "SELECT distinct DateStart, DateEnd FROM tblbillbyperiod WHERE bldgnum='"&building&"' and billperiod="&bperiod&" and billyear="&byear, cnn1
if not perioddates.EOF then
	displaydate = " ("&month(perioddates("DateStart"))&"/"&day(perioddates("DateStart"))&" - "&month(perioddates("DateEnd"))&"/"&day(perioddates("DateEnd"))&")"
end if
perioddates.close

dim previousMeterid, isposted
isposted = false
if not rst1.eof then if rst1("posted")="True" then isposted = true
%>
<html>
<head><title>Bill Validation</title></head>
<script>
var checkboxf = 0
function updatemeter(meterid, byear, bperiod, tnumber, tname)
{	if(checkboxf==0)
	{	var newwin = open('update_billentry.asp?meterid='+meterid+'&byear='+byear+'&bperiod='+bperiod+'&tname='+tname+'&tnumber='+tnumber+'&building=<%=building%>&posted=<%=isposted%>', 'update_billentry','scrollbars=yes,width=900, height=380, status=no');
		newwin.focus();
	}
}

function scrollpoint()
{ return(document.all['meterlist'].scrollTop);
}

function movemeterlist(y)
{ document.all['meterlist'].scrollTop = y;
}

function nullfunction()
{
}
</script>
<body bgcolor="#FFFFFF" LINK="#000099" vlink="#000099" alink="#000099" onload="movemeterlist(<%=yscroll%>)">
<%
dim numoftenants, numofmeters, numoftenantsPrev, numofmetersPrev 'need to fill these variables next section (assume are zero)
numoftenants =0
numofmeters =0
numoftenantsPrev =0
numofmetersPrev =0
dim rst2, strsql2, prevbyear, prevbperiod
set rst2 = server.createobject("ADODB.Recordset")
prevbyear = byear
prevbperiod = bperiod-1
if prevbperiod<1 then 
	prevbperiod = 12
	prevbyear = prevbyear-1
end if
strsql2 = "SELECT (SELECT count(Distinct LeaseUtilityID) FROM tblMetersByPeriod WHERE bldgnum='"&building&"' and billperiod="&bperiod&" and billyear="&byear&") as tenants, (SELECT count(*) FROM tblmetersbyperiod c WHERE c.bldgnum='"&building&"' and c.BillYear="&byear&" and c.BillPeriod="&bperiod&") as meters"
rst2.open strsql2, cnn1
if not rst2.EOF then
	numofmeters = rst2("meters")
	numoftenants = rst2("tenants")
end if
rst2.close

strsql2 = "SELECT (SELECT count(Distinct LeaseUtilityID) FROM tblMetersByPeriod WHERE bldgnum='"&building&"' and billperiod="&prevbperiod&" and billyear="&prevbyear&") as tenants, (SELECT count(*) FROM tblmetersbyperiod c WHERE c.bldgnum='"&building&"' and c.BillYear="&prevbyear&" and c.BillPeriod="&prevbperiod&") as meters"
rst2.open strsql2, cnn1
if not rst2.EOF then
	numofmetersPrev = rst2("meters")
	numoftenantsPrev = rst2("tenants")
end if
rst2.close

dim prevbuildingAvgKW, prevbuildingAvgKWH, prevbuildingAvgBillAmt, avgBuildingCostKW, avgBuildingCostKWH, avgBuildingFuelAdj'now get these header fields as well
strsql2 = "SELECT isNull(avg(ub.fuelAdj),0) as avgfueladj, isNull(avg(TotalKW),0) as avgTotalKW, isNull(avg(TotalKWH),0) as avgTotalKWH, isNull(avg(TotalBillAmt),0) as avgTotalBillAmt, isNull(avg(CostKW),0) as avgCostKW, isNull(avg(CostKWH),0) as avgCostKWH FROM tblBillByPeriod bbp INNER JOIN utilitybill ub ON ub.ypid=bbp.ypid Where ((billyear="&byear&" and billperiod<"&bperiod&" and billperiod>="&bperiod&"-3)or(billyear="&byear-1&" and billperiod>="&bperiod&"+9)) and bbp.bldgnum='"& building &"'"
rst2.open strsql2, cnn1
'response.write strsql2
'response.end
if not rst2.EOF then
	prevbuildingAvgKW = rst2("avgTotalKW")
	prevbuildingAvgKWH = rst2("avgTotalKWH")
	prevbuildingAvgBillAmt = rst2("avgTotalBillAmt")
	avgBuildingCostKW = rst2("avgCostKW")
	avgBuildingCostKWH = rst2("avgCostKWH")
	avgBuildingFuelAdj = rst2("avgfueladj")
end if
rst2.close

dim currentkw, currentkwh, currentcostkw, currentcostkwh, currentfueladj, buildingname, currentBillAmt 'totals (and building name for top)
'dim aveKWH'averages
strsql2 = "select (select distinct strt from buildings where bldgnum='"&building&"') as building, FuelAdj, TotalKWH, AvgKWH, TotalKW, CostKWH, CostKW, TotalBillAmt FROM utilitybill where ypid in (select ypid FROM billyrperiod where bldgnum='"&building&"' and Billyear="&byear&" and BillPeriod="&bperiod&")"
rst2.open strsql2, cnn1
if not rst2.EOF then
	buildingname = rst2("building")
	currentkw = rst2("TotalKW")
	currentkwh = rst2("TotalKWH")
	currentcostkw = rst2("CostKW")
	currentcostkwh = rst2("CostKWH")
	currentBillAmt = rst2("TotalBillAmt")
	currentfueladj = rst2("FuelAdj")
end if
rst2.close
'rst.open strsql2
dim bKWflag, bKWHflag, bCostKWflag, bCostKWHflag, tenantsFlag, metersFlag'the four building variance flags get set if building variance below is to high
if prevbuildingAvgKW<>0 then if abs(currentkw-prevbuildingAvgKW)/prevbuildingAvgKW>.2 then bKWflag = " bgcolor="""&flagcolor&""""
if prevbuildingAvgKWH<>0 then if abs(currentkwh-prevbuildingAvgKWH)/prevbuildingAvgKWH>.2 then bKWHflag = " bgcolor="""&flagcolor&""""
if avgBuildingCostKW<>0 then if abs(currentcostkw-avgBuildingCostKW)/avgBuildingCostKW>.2 then bCostKWflag = " bgcolor="""&flagcolor&""""
if avgBuildingCostKWH<>0 then if abs(currentcostkwh-avgBuildingCostKWH)/avgBuildingCostKWH>.2 then bCostKWHflag = " bgcolor="""&flagcolor&""""
if numoftenants<>numoftenantsPrev then tenantsFlag = " bgcolor="""&flagcolor&""""
if numofmeters<>numofmetersPrev then metersFlag = " bgcolor="""&flagcolor&""""
strsql2 = "select count(meternum) as mcount from tblmetersbyperiod where meterid not in (SELECT c.meterid FROM consumption c, peakdemand p WHERE p.meterid=c.meterid and c.billperiod=p.billperiod and c.billyear=p.billyear and c.billyear="&byear&" and c.billperiod="&bperiod&") and BldgNum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
rst2.open strsql2, cnn1
if cint(rst2("mcount"))>0 then metersFlag = " bgcolor="""&flagcolor&""""
rst2.close


%>
<table width="100%" border="0" bgcolor="#FFFFFF"><tr><td bgcolor="#3399CC" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">
<%=buildingname%>&nbsp;<%=byear%>,&nbsp;Period&nbsp;<%=bperiod%><%=displaydate%>
</font></b></td></tr></table>
<center>
<table border="0" width="900"><tr><td>
	<table width="650" border="0" cellspacing="0" cellpadding="3">
	<tr style="background-color: #0099FF; font-family: Arial, Helvetica, sans-serif;font-size:13">
		<td width="2%">&nbsp;</td>
		<td align="right">KW</td>
		<td align="right">KWH</td>
		<td align="right">Cost KW</td>
		<td align="right">Cost KWH</td>
		<td align="right">Fuel Adjustment</td>
		<td align="right">Total Bill Amount</td>
	</tr>
	<tr style="background-color: #CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12">
		<td style="background-color: #0099FF;">Average</td>
		<td align="right"<%=bKWflag%>><%=formatnumber(prevbuildingAvgKW)%></td>
		<td align="right"<%=bKWHflag%>><%=formatnumber(prevbuildingAvgKWH,0)%></td>
		<td align="right"<%=bCostKWflag%>><%=formatcurrency(avgBuildingCostKW)%></td>
		<td align="right"<%=bCostKWHflag%>><%=formatcurrency(avgBuildingCostKWH)%></td>
		<td align="right"><%=formatcurrency(avgBuildingFuelAdj,6)%></td>
		<td align="right"><%=formatcurrency(prevbuildingAvgBillAmt)%></td></tr>
	<tr style="background-color: #CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12">
		<td style="background-color: #0099FF;">Current</td>
		<td align="right"<%=bKWflag%>><%=formatnumber(currentkw)%></td>
		<td align="right"<%=bKWHflag%>><%=formatnumber(currentkwh,0)%></td>
		<td align="right"<%=bCostKWflag%>><%=formatcurrency(currentcostkw)%></td>
		<td align="right"<%=bCostKWHflag%>><%=formatcurrency(currentcostkwh)%></td>
		<td align="right"><%=formatcurrency(currentfueladj,6)%></td>
		<td align="right"><%=formatcurrency(currentBillAmt)%></td></tr>
	</table>
</td><td>
	<table width="250" border="0" cellspacing="0" cellpadding="3">
	<tr style="background-color: #0099FF; font-family: Arial, Helvetica, sans-serif;font-size:13">
		<td width="1%"></td>
		<td align="right">Tenants&nbsp;Billed</td>
		<td align="right">Meters&nbsp;Billed</td></tr>
	<tr style="background-color: #CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12">
		<td style="background-color: #0099FF;">This&nbsp;Period</td>
		<td align="right"<%=tenantsFlag%>><a href="javascript:nullfunction()" onclick="window.open('tenantmeterlist.asp?building=<%=building%>&bperiod=<%=bperiod%>&byear=<%=byear%>&checking=Tenant', '', 'toolbar=no,width=250,height=200, resizable=no,scrollbars=yes')"><%=numoftenants%></a></td>
		<td align="right"<%=metersFlag%>><a href="javascript:nullfunction()" onclick="window.open('tenantmeterlist.asp?building=<%=building%>&bperiod=<%=bperiod%>&byear=<%=byear%>&checking=Meter', '', 'toolbar=no,width=250,height=200, resizable=no,scrollbars=yes')"><%=numofmeters%></a></td></tr>
	<tr style="background-color: #CCCCCC; font-family: Arial, Helvetica, sans-serif;font-size:12">
		<td style="background-color: #0099FF;">Last&nbsp;Period</td>
		<td align="right"<%=tenantsFlag%>><%=numoftenantsPrev%></td>
		<td align="right"<%=metersFlag%>><%=numofmetersPrev%></td></tr>
	</table>
</td></tr></table>
<form name="form1" method="post" action="<%=procPage%>">
<table border="0" cellspacing="0" cellpadding="0"><tr><td>
<table border="0" cellspacing="0" cellpadding="0" style="font-family: Arial, Helvetica, sans-serif;font-size:13;color:white">
<tr>
<%if not super then%><td width="100" bgcolor="<%=btabcolor%>" align="center">&nbsp;<b>Biller&nbsp;Validation</b>&nbsp;</td><%else%>
	<td width="100" bgcolor="<%=stabcolor%>" align="center">&nbsp;<b>Supervisor&nbsp;Validation</b>&nbsp;</td>
<%end if%>
</tr></table>
<div style="border: 2px solid #0099FF;width:900;background-color: #CCCCCC;">
<table width="100%" border="0" cellspacing="0" cellpadding="3">
<tr style="background-color: #0099FF; font-family: Arial, Helvetica, sans-serif;font-size:13">
<%if not isposted then%><td width="40">Accept</td><%end if%>
<td width="70">Tenant #</td>
<td width="120">Tenant Name</td>
<td width="60">Meter #</td>
<td width="60">Average KWH</td>
<td width="60">Current KWH Usage</td>
<td width="60">Variance KWH</td>
<td width="60">Average KW</td>
<td width="60">Current KW</td>
<td>Variance KW</td>
<%if super then%>
<td width="60">Bill Amount</td>
<td width="60">Average Amount</td>
<td width="60">Variance Amount</td>
<td>Updated KWH/KW</td>
<%end if%>
</tr></table>
<div  id="meterlist" style="overflow:auto;height:200">
<table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
previousMeterid = ""
do until rst1.eof
	if previousMeterid<>trim(rst1("meterid")) then
		dim rowColor, ischecked, tenantname, hilight, kwvartemp, kwhvartemp
		kwvartemp = rst1("kwvarience")
		if not isnumeric(trim(kwvartemp)) then kwvartemp = 0
		kwhvartemp = rst1("kwhvarience")
		if not isnumeric(trim(kwhvartemp)) then kwhvartemp = 0
		if formatnumber(kwvartemp,2)>20 or formatnumber(kwhvartemp,2)>20 or (clng(rst1("avgKWH")) = 0) then
			rowColor = flagColor
			hilight = flagColorHilight
			ischecked = ""
		else
			rowColor = "#CCCCCC"
			hilight = "#DDDDDD"
			ischecked = " CHECKED"
		end if
		if rst1("validate")="True" and Not(super) then ischecked = " CHECKED"
		if rst1("svalidate")="True" and super then ischecked = " CHECKED"
		tenantname = rst1("billingname")
		if len(tenantname)>10 then tenantname = left(tenantname,10)&"..."
'response.write "billingname: "&strsql&"||"
'response.end		
		response.write "<tr style=""background-color: "& rowColor &"; font-family: Arial, Helvetica, sans-serif;font-size:12;cursor:hand"" valign=""top"" onclick=""updatemeter("&rst1("meterid")&", "&byear&", "&bperiod&", '"&rst1("tenantnum")&"', '["&server.urlencode(rst1("billingname"))&"]')"" onMouseOver=""this.style.backgroundColor='"& hilight &"'"" onMouseOut=""this.style.backgroundColor='"& rowColor &"'"">"
		if not isposted then response.write "<td width=""40""><input type=""checkbox"" value="""&rst1("meterid")&""" name=""meters"" onMouseOver=""checkboxf=1"" onMouseOut=""checkboxf=0"" style=""cursor:auto"""& ischecked &"></td>"
		response.write "<td width=""70"">"&rst1("tenantnum")&"</td>"
		response.write "<td width=""120""><nobr>"&tenantname&"</nobr></td>"
		response.write "<td width=""60"">"&rst1("meternum")&"</td>"
		response.write "<td width=""60"">"&formatnumber(rst1("avgKWH"),0)&"</td>"
		response.write "<td width=""60"">"&formatnumber(rst1("KWHused"),0)&"</td>"
		response.write "<td width=""60"">"&formatnumber(kwhvartemp,0)&"%</td>"
		response.write "<td width=""60"">"&formatnumber(rst1("AvgKW"),2)&"</td>"
		response.write "<td width=""60"">"&formatnumber(rst1("demand"),2)&"</td>"
		response.write "<td>"&formatnumber(kwvartemp,2)&"%</td>"
		if super then
			response.write "<td width=""60"">"&formatcurrency(rst1("totalamt"),2)&"</td>"
			response.write "<td width=""60"">"&formatcurrency(rst1("avgAmt"),2)&"</td>"
			response.write "<td width=""60"">"&formatnumber(rst1("Amtvarience"),2)&"%</td>"
			response.write "<td>"
			if rst1("validate")="True" then response.write "Accepted<br>"
			if trim(rst1("biller"))<>"" then response.write rst1("biller")&"<br>"&rst1("org_kwh")&"kwh/ "&rst1("org_kw")&"kw"
			response.write "&nbsp;</td>"
		end if
		response.write "</tr>"&vbNewLine
		previousMeterid=trim(rst1("meterid"))
	end if
	rst1.movenext
loop
'response.write strsql
%>
</tr></table>
</div>
<%if not isposted then%>
<table width="100%" border="0" cellspacing="0" cellpadding="7">
<tr style="font-family: Arial, Helvetica, sans-serif;font-size:13" valign="top">
	<td width="95%" style="background-color: #CCCCCC;" align="right"><%if super then%>post bills <input name="post" value="1" type="checkbox"<%if isposted then response.write " CHECKED"%>><%end if%></td>
	<td width="5%" style="background-color: #0000FF;cursor:hand" onclick="document.forms['form1'].yscroll.value=scrollpoint();document.forms['form1'].submit();"><b>Accept</b></td>
</table>
<%end if%>
</div>
</td></tr></table>
<table cellpadding="3">
<tr style="font-family: Arial, Helvetica, sans-serif;font-size:13"><tr style="font-family: Arial, Helvetica, sans-serif;font-size:13"><td><table width="8" height="8" cellpadding="0" cellspacing="0" border="0" bgcolor="<%=flagColor%>"><tr><td></td></tr></table></td><td>Signifies a variance of over 20% (except in case of meter numbers)</td></tr>
</table>
<input type="hidden" name="byear" value="<%=byear%>">
<input type="hidden" name="bperiod" value="<%=bperiod%>">
<input type="hidden" name="building" value="<%=building%>">
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="buildingname" value="<%=buildingname%>">
<input type="hidden" name="yscroll" value="0">
</form>
</center>
</body>
</html>

