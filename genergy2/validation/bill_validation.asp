<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim super, btabcolor, stabcolor, flagColor, byear, bperiod, building, flagColorHilight, procPage, pid, utilityid, yscroll
building = request.querystring("building")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
pid = request.querystring("pid")
utilityid = request("utilityid")
yscroll = request("yscroll")
if trim(yscroll)="" then yscroll = 100

dim rst1, cnn1, strsql, cnn2
set cnn1 = server.createobject("ADODB.connection")
set cnn2 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(building)
cnn2.open application("cnnstr_main")

rst1.open "SELECT Supervisor from employees WHERE substring(username,7,90)='"&getXmlUserName()&"'", cnn2
if not rst1.eof then
	if rst1("Supervisor")="True" then
		super=true
	else
		super=false
	end if
end if
rst1.close
session("isSuper") = super

dim usage, demand, UBtable
select case cint(utilityid)
case 1
	usage = "Mlbs/hr"
	demand = "Mlbs"
  UBtable = "utilitybill_steam"
case 2
	usage = "KWH"
	demand = "KW"
  UBtable = "utilitybill"
case 3
	usage = "CCF"
	demand = "-"
  UBtable = "utilitybill_coldwater"
case 4
	usage = "CF"
	demand = "-"
  UBtable = "utilitybill_gas"
case 6
	usage = "Ton/hr"
	demand = "Tons"
  UBtable = "utilitybill_chilledwater"
case else
	usage = "?"
	demand = "?"
end select

if super then
	'Supervisor
	stabcolor="#0099FF"
	btabcolor="#CCCCCC"
	flagColor = "#009900"
	flagColorHilight = "#009900"
	procPage = "super_process.asp"
	strsql = "SELECT distinct m.meterid, m.meternum, m.variance, v.revdate, c.validate, c.svalidate, bbp.posted, m.bldgnum,c.meterid, c.[current], isNull(c.used,0) as kwhused, isNull(pd.demand,0) as demand, l.tenantnum, l.billingname, isNull(bbp.totalamt,0) as totalamt, bbp.adminfee, bbp.sqft, pd.demand, v.biller, v.org_kwh, v.org_kw, "&_
	"case when bbp.sqft=0 then 0 else(bbp.demand/bbp.sqft)end as wsqft, "&_
	"isNuLL((SELECT avg(used) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3))),0) as avgKWH, "&_
	"isNuLL((SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3))),0) as avgKW, "&_
	"isNuLL((SELECT avg(totalamt) FROM tblbillbyperiod bbp2 WHERE bbp2.leaseutilityid=bbp.leaseutilityid and ((bbp2.billyear="&byear-1&" and bbp2.billperiod>="&bperiod&"+9)or(bbp2.billyear="&byear&" and bbp2.billperiod<"&bperiod&" and bbp2.billperiod>="&bperiod&"-3))),0) as avgAmt, "&_
	"isNull(case when (SELECT avg(used) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((c.used - (SELECT avg(used) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3))))/(SELECT avg(used) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3)))*100) end, 0) as kwhvarience, "&_
	"isNull(case when (SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((pd.demand - (SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3))))/(SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)))*100) end, 0) as kwvarience, "&_
	"isNull(case when (SELECT avg(totalamt) FROM tblbillbyperiod bbp2 WHERE bbp2.leaseutilityid=bbp.leaseutilityid and ((bbp2.billyear="&byear-1&" and bbp2.billperiod>="&bperiod&"+9)or(bbp2.billyear="&byear&" and bbp2.billperiod<"&bperiod&" and bbp2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((bbp.totalamt - (SELECT avg(totalamt) FROM tblbillbyperiod bbp2 WHERE bbp2.leaseutilityid=bbp.leaseutilityid and ((bbp2.billyear="&byear-1&" and bbp2.billperiod>="&bperiod&"+9)or(bbp2.billyear="&byear&" and bbp2.billperiod<"&bperiod&" and bbp2.billperiod>="&bperiod&"-3))))/(SELECT avg(totalamt) FROM tblbillbyperiod bbp2 WHERE bbp2.leaseutilityid=bbp.leaseutilityid and ((bbp2.billyear="&byear-1&" and bbp2.billperiod>="&bperiod&"+9)or(bbp2.billyear="&byear&" and bbp2.billperiod<"&bperiod&" and bbp2.billperiod>="&bperiod&"-3)))*100) end, 0) as Amtvarience "&_
	"FROM consumption c  "&_
	"INNER JOIN meters m ON m.Meterid=c.Meterid  "&_
	"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
	"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
	"INNER JOIN tblleases l ON lup.billingid=l.billingid "&_
	"LEFT JOIN tblbillbyperiod bbp on lup.leaseutilityid=bbp.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod "&_
	"LEFT JOIN validation v on m.Meterid=v.Meterid and c.billyear=v.billyear and c.billperiod=v.billperiod "&_
	"WHERE c.billyear="&byear&" and c.billperiod="&bperiod&" and m.bldgnum='"&building&"' and lup.utility="&utilityid&" and m.lmp<>1 and Online='1' and leaseexpired=0 "&_
	"ORDER BY l.billingname, m.meternum, v.revdate desc"
else 'and BillPeriod not in (SELECT billperiod FROM tblbillbyperiod WHERE bldgnum='"&building&"' and billyear ="&byear&" and posted=1 group by billyear, billperiod)
	'Biller
	stabcolor="#CCCCCC"
	btabcolor="#6699cc"
	flagColor = "#cc0000"
	flagColorHilight = "#ff0000"
	procPage = "biller_process.asp"
	strsql = "SELECT Distinct m.meternum, m.variance, c.validate, bbp.posted, c.svalidate, m.bldgnum,c.meterid, c.[current], isNull(c.used,0) as kwhused, isNull(pd.demand,0) as demand, l.tenantnum, isnull(l.billingname,'') as billingname, "&_
	"isNuLL((SELECT avg(used) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3))),0) as avgKWH, "&_
	"isNuLL((SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3))),0) as avgKW, "&_
	"isNull(case when (SELECT avg(used) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((c.used - (SELECT avg(used) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3))))/(SELECT avg(used) FROM consumption c2 WHERE c2.meterid=c.meterid and ((c2.billyear="&byear&"-1 and c2.billperiod>="&bperiod&"+9)or(c2.billyear="&byear&" and c2.billperiod<"&bperiod&" and c2.billperiod>="&bperiod&"-3)))*100) end, 0) as kwhvarience, "&_
	"isNull(case when (SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)))=0 then '0' else abs((pd.demand - (SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3))))/(SELECT avg(demand) FROM peakdemand d2 WHERE d2.meterid=c.meterid and ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)))*100) end, 0) as kwvarience "&_
	"FROM consumption c "&_
	"INNER JOIN meters m ON m.Meterid=c.Meterid "&_
	"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
	"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
	"INNER JOIN tblleases l on lup.billingid=l.billingid "&_
	"LEFT JOIN tblbillbyperiod bbp on m.leaseutilityid=bbp.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod "&_
	"WHERE c.billyear="&byear&" and c.billperiod="&bperiod&" and m.bldgnum='"&building&"' and lup.utility="&utilityid&" and m.lmp<>1 and Online='1' and leaseexpired=0 "&_
	"ORDER BY l.billingname, m.meternum"
end if

'response.write strsql
'response.end
rst1.open strsql, cnn1
dim displaydate, perioddates ' get start and end dates of current period for header display
set perioddates = server.createobject("ADODB.recordset")
perioddates.open "SELECT distinct DateStart, DateEnd FROM tblbillbyperiod WHERE bldgnum='"&building&"' and billperiod="&bperiod&" and utility="&utilityid&" and billyear="&byear, cnn1
if not perioddates.EOF then
	displaydate = " ("&month(perioddates("DateStart"))&"/"&day(perioddates("DateStart"))&" - "&month(perioddates("DateEnd"))&"/"&day(perioddates("DateEnd"))&")"
end if
perioddates.close

dim previousMeterid, isposted, needAcceptButton
needAcceptButton = false
%>
<html>
<head><title>Bill Validation</title>
<script>
var checkboxf = 0
function updatemeter(meterid, byear, bperiod, tnumber, tname, posted)
{	if(checkboxf==0)
	{	var newwin = open('update_billentry.asp?meterid='+meterid+'&byear='+byear+'&bperiod='+bperiod+'&tname='+tname+'&tnumber='+tnumber+'&building=<%=building%>&pid=<%=pid%>&utilityid=<%=utilityid%>&posted='+posted, 'update_billentry','left=8,top=8,scrollbars=yes,width=770, height=380, status=no');
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
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #dddddd; }
</style>
</head>
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
strsql2 = "SELECT (SELECT count(Distinct m.LeaseUtilityID) FROM tblMetersByPeriod m, tblleasesutilityprices l WHERE l.leaseutilityid=m.leaseutilityid and m.bldgnum='"&building&"' and billperiod="&bperiod&" and billyear="&byear&" and l.utility="&utilityid&") as tenants, (SELECT count(*) FROM tblmetersbyperiod c, tblleasesutilityprices l WHERE l.leaseutilityid=c.leaseutilityid and c.bldgnum='"&building&"' and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and l.utility="&utilityid&") as meters"
rst2.open strsql2, cnn1
if not rst2.EOF then
	numofmeters = rst2("meters")
	numoftenants = rst2("tenants")
end if
rst2.close

strsql2 = "SELECT (SELECT count(Distinct m.LeaseUtilityID) FROM tblMetersByPeriod m, tblleasesutilityprices l WHERE l.leaseutilityid=m.leaseutilityid and m.bldgnum='"&building&"' and billperiod="&prevbperiod&" and billyear="&prevbyear&" and l.utility="&utilityid&") as tenants, (SELECT count(*) FROM tblmetersbyperiod c, tblleasesutilityprices l WHERE l.leaseutilityid=c.leaseutilityid and c.bldgnum='"&building&"' and c.BillYear="&prevbyear&" and c.BillPeriod="&prevbperiod&" and l.utility="&utilityid&") as meters"
rst2.open strsql2, cnn1
if not rst2.EOF then
	numofmetersPrev = rst2("meters")
	numoftenantsPrev = rst2("tenants")
end if
rst2.close

dim prevbuildingAvgKW, prevbuildingAvgKWH, prevbuildingAvgBillAmt, avgBuildingCostKW, avgBuildingCostKWH, avgBuildingFuelAdj'now get these header fields as well
if utilityid=2 then
  strsql2 = "SELECT isNull(avg(ub.fuelAdj),0) as avgfueladj, isNull(avg(TotalKW),0) as avgTotalKW, isNull(avg(TotalKWH),0) as avgTotalKWH, isNull(avg(TotalBillAmt),0) as avgTotalBillAmt, isNull(avg(CostKW),0) as avgCostKW, isNull(avg(CostKWH),0) as avgCostKWH FROM tblBillByPeriod bbp INNER JOIN "&UBtable&" ub ON ub.ypid=bbp.ypid Where ((billyear="&byear&" and billperiod<"&bperiod&" and billperiod>="&bperiod&"-3)or(billyear="&byear-1&" and billperiod>="&bperiod&"+9)) and bbp.utility="&utilityid&" and bbp.bldgnum='"& building &"'"
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
end if

dim currentkw, currentkwh, currentcostkw, currentcostkwh, currentfueladj, buildingname, currentBillAmt 'totals (and building name for top)
'dim aveKWH'averages
if utilityid=2 then
  strsql2 = "select (select distinct strt from buildings where bldgnum='"&building&"') as building, FuelAdj, sum(TotalKWH) as TotalKWH, sum(CostKWH)/sum(TotalKWH) as AvgKWH, sum(TotalKW) as TotalKW, sum(CostKWH) as CostKWH, sum(CostKW) as CostKW, sum(TotalBillAmt) as TotalBillAmt  FROM "&UBtable&" where ypid in (select ypid FROM billyrperiod where bldgnum='"&building&"' and Billyear="&byear&" and BillPeriod="&bperiod&") GROUP BY FuelAdj"
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
end if

dim bKWflag, bKWHflag, bCostKWflag, bCostKWHflag, tenantsFlag, metersFlag'the four building variance flags get set if building variance below is to high
if prevbuildingAvgKW<>0 then if abs(currentkw-prevbuildingAvgKW)/prevbuildingAvgKW>.2 then bKWflag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
if prevbuildingAvgKWH<>0 then if abs(currentkwh-prevbuildingAvgKWH)/prevbuildingAvgKWH>.2 then bKWHflag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
if avgBuildingCostKW<>0 then if abs(currentcostkw-avgBuildingCostKW)/avgBuildingCostKW>.2 then bCostKWflag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
if avgBuildingCostKWH<>0 then if abs(currentcostkwh-avgBuildingCostKWH)/avgBuildingCostKWH>.2 then bCostKWHflag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
if numoftenants<>numoftenantsPrev then tenantsFlag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""
if numofmeters<>numofmetersPrev then metersFlag = " style=""color:"&flagcolor&";border-bottom:1px solid #dddddd;"""

%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td bgcolor="#6699cc"><span class="standardheader"><%=buildingname%> &nbsp;&nbsp;Period&nbsp;<%=bperiod%><%=displaydate%>,&nbsp;<%=byear%></span></td>
</tr>
</table>

<table border=0 cellpadding="3" cellspacing="1" width="100%">
<tr bgcolor="#dddddd">
  <td width="8%">&nbsp;</td>
  <td width="10%" align="right"><%=demand%></td>
  <td width="10%" align="right"><%=usage%></td>
  <td width="10%" align="right">Cost <%=demand%></td>
  <td width="10%" align="right">Cost <%=usage%></td>
  <td width="10%" align="right">Fuel Adjustment</td>
  <td width="12%" align="right">Total Bill Amount</td>
  <td width="10%">&nbsp;</td>
  <td width="10%" align="right">Tenants&nbsp;Billed</td>
  <td width="10%" align="right">Meters&nbsp;Billed</td>
</tr>
<tr>
  <td align="right" class="tblunderline">Average:</td>
  <td class="tblunderline" align="right"<%=bKWflag%>><%=formatnumber(prevbuildingAvgKW)%></td>
  <td class="tblunderline" align="right"<%=bKWHflag%>><%=formatnumber(prevbuildingAvgKWH,0)%></td>
  <td class="tblunderline" align="right"<%=bCostKWflag%>><%=formatcurrency(avgBuildingCostKW)%></td>
  <td class="tblunderline" align="right"<%=bCostKWHflag%>><%=formatcurrency(avgBuildingCostKWH)%></td>
  <td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(avgBuildingFuelAdj,6)%></td>
  <td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(prevbuildingAvgBillAmt)%></td>
  <td align="right" bgcolor="#eeeeee">This&nbsp;Period:</td>
  <td bgcolor="#eeeeee" align="right"<%=tenantsFlag%>><a href="javascript:nullfunction()" onclick="window.open('tenantmeterlist.asp?building=<%=building%>&bperiod=<%=bperiod%>&byear=<%=byear%>&utilityid=<%=utilityid%>&checking=Tenant', '', 'toolbar=no,width=250,height=200, resizable=no,scrollbars=yes')"><%=numoftenants%></a></td>
  <td bgcolor="#eeeeee" align="right"<%=metersFlag%>><a href="javascript:nullfunction()" onclick="window.open('tenantmeterlist.asp?building=<%=building%>&bperiod=<%=bperiod%>&byear=<%=byear%>&utilityid=<%=utilityid%>&checking=Meter', '', 'toolbar=no,width=250,height=200, resizable=no,scrollbars=yes')"><%=numofmeters%></a></td>
</tr>
<tr>
  <td align="right" class="tblunderline">Current:</td>
  <td class="tblunderline" align="right"<%=bKWflag%>><%=formatnumber(currentkw)%></td>
  <td class="tblunderline" align="right"<%=bKWHflag%>><%=formatnumber(currentkwh,0)%></td>
  <td class="tblunderline" align="right"<%=bCostKWflag%>><%=formatcurrency(currentcostkw)%></td>
  <td class="tblunderline" align="right"<%=bCostKWHflag%>><%=formatcurrency(currentcostkwh)%></td>
  <td class="tblunderline" align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(currentfueladj,6)%></td>
  <td class="tblunderline"align="right" style="border-bottom:1px solid #dddddd;"><%=formatcurrency(currentBillAmt)%></td>
  <td align="right" bgcolor="#eeeeee">Last&nbsp;Period:</td>
  <td bgcolor="#eeeeee" align="right"<%=tenantsFlag%>><%=numoftenantsPrev%></td>
  <td bgcolor="#eeeeee" align="right"<%=metersFlag%>><%=numofmetersPrev%></td>
  </tr>
</table>

<form name="form1" method="post" action="<%=procPage%>">
<table border=0 cellpadding="3" cellspacing="0">
<tr>
  <td>
  <%if not super then%><b>Biller&nbsp;Validation</b><%else%>
	<b>Supervisor&nbsp;Validation</b>
  <%end if%>
  </td>
</tr>
</table>

<%
dim cellw
'if not isposted then cellw = "90" else 
cellw = "70"
%>
<div style="border: 1px solid #cccccc;margin:3px;width:100%;">
<table width="100%" border=0 cellspacing="1" cellpadding="3">
<tr bgcolor="#dddddd" valign="bottom" style="font-weight:bold;">
<td width="40">Accept</td>
<td width="<%=cellw%>">Tenant Number</td>
<td width="<%=cellw%>">Tenant Name</td>
<td width="<%=cellw%>">Meter</td>
<td width="<%=cellw%>">Average <%=usage%></td>
<td width="<%=cellw%>">Current <%=usage%> Usage</td>
<td width="<%=cellw%>">Variance <%=usage%></td>
<td width="<%=cellw%>">Average <%=demand%></td>
<td width="<%=cellw%>">Current <%=demand%></td>
<td>Variance <%=demand%></td>
<%if super then%>
<td width="<%=cellw%>">Bill Amount</td>
<td width="<%=cellw%>">Average Amount</td>
<td width="<%=cellw%>">Variance Amount</td>
<td>Updated <%=usage%>/<%=demand%></td>
<%end if%>
</tr></table>
<div id="meterlist" style="overflow:auto;height:200">
<table width="100%" border="0" cellspacing="1" cellpadding="3">
<%
previousMeterid = ""
do until rst1.eof
	isposted = false
	if rst1("posted")="True" then isposted = true else needAcceptButton = true
	if previousMeterid<>trim(rst1("meterid")) then
		dim rowColor, ischecked, tenantname, hilight, kwvartemp, kwhvartemp
		kwvartemp = rst1("kwvarience")
		if not isnumeric(trim(kwvartemp)) then kwvartemp = 0
		kwhvartemp = rst1("kwhvarience")
		if not isnumeric(trim(kwhvartemp)) then kwhvartemp = 0
		if cdbl(kwvartemp)>cdbl(rst1("variance"))*100 or cdbl(kwhvartemp)>cdbl(rst1("variance"))*100 or (clng(rst1("avgKWH")) = 0) then
			rowColor = flagColor
			hilight = flagColorHilight
			ischecked = ""
		else
			rowColor = "#666666"
			hilight = "#DDDDDD"
			ischecked = " CHECKED"
		end if
		if rst1("validate")="True" and Not(super) then ischecked = " CHECKED"
		if rst1("svalidate")="True" and super then ischecked = " CHECKED"
		tenantname = rst1("billingname")
		if len(tenantname)>10 then tenantname = left(tenantname,10)&"..."
		response.write "<tr style=""color: "& rowColor &";cursor:hand"" valign=""top"" onclick=""updatemeter("&rst1("meterid")&", "&byear&", "&bperiod&", '"&rst1("tenantnum")&"', '["&server.urlencode(rst1("billingname"))&"]', '"&isposted&"')"" onMouseOver=""this.style.backgroundColor='"& hilight &"';this.style.color='#ffffff';"" onMouseOut=""this.style.backgroundColor='#ffffff';this.style.color='"& rowcolor &"';"">"
		if not isposted then response.write "<td width=""40""><input type=""checkbox"" value="""&rst1("meterid")&""" name=""meters"" onMouseOver=""checkboxf=1"" onMouseOut=""checkboxf=0"" style=""cursor:auto"""& ischecked &"></td>" else response.write "<td width=""40"">S.&nbsp;Acc</td>"
		response.write "<td width="""&cellw&""">"&rst1("tenantnum")&"</td>"
		response.write "<td width="""&cellw&"""><nobr>"&tenantname&"</nobr></td>"
		response.write "<td width="""&cellw&""">"&rst1("meternum")&"</td>"
		response.write "<td width="""&cellw&""">"&formatnumber(rst1("avgKWH"),0)&"</td>"
		response.write "<td width="""&cellw&""">"&formatnumber(rst1("kwhused"),0)&"</td>"
		response.write "<td width="""&cellw&""">"&formatnumber(kwhvartemp,0)&"%</td>"
		response.write "<td width="""&cellw&""">"&formatnumber(rst1("AvgKW"),2)&"</td>"
		response.write "<td width="""&cellw&""">"&formatnumber(rst1("demand"),2)&"</td>"
		response.write "<td>"&formatnumber(kwvartemp,2)&"%</td>"
		if super then
			response.write "<td width="""&cellw&""">"&formatcurrency(rst1("totalamt"),2)&"&nbsp;</td>"
			response.write "<td width="""&cellw&""">"&formatcurrency(rst1("avgAmt"),2)&"&nbsp;</td>"
			response.write "<td width="""&cellw&""">"&formatnumber(rst1("Amtvarience"),2)&"%&nbsp;</td>"
			response.write "<td>"
			if rst1("validate")="True" then response.write "Accepted<br>"
			if trim(rst1("biller"))<>"" then response.write rst1("biller")&"<br>"&rst1("org_kwh")&usage&"/ "&rst1("org_kw")&demand
			response.write "&nbsp;</td>"
		end if
'		response.write "<td>"&formatnumber(rst1("variance")*100,2)&"%</td>"
		response.write "</tr>"&vbNewLine
		previousMeterid=trim(rst1("meterid"))
	end if
	rst1.movenext
loop
'response.write strsql
%>
</tr></table>
</div>
<%if needAcceptButton then%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td style="padding-top:4px;" bgcolor="#eeeeee"><input type="button" value="Accept" onclick="document.forms['form1'].yscroll.value=scrollpoint();document.forms['form1'].submit();" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;"></td>
</tr>
</table>
<%end if%>
</div>

<table cellpadding="3">
<tr style="font-family: Arial, Helvetica, sans-serif;font-size:13"><tr style="font-family: Arial, Helvetica, sans-serif;font-size:13"><td><table width="8" height="8" cellpadding="0" cellspacing="0" border="0" bgcolor="<%=flagColor%>"><tr><td></td></tr></table></td><td>Signifies a variance of over 20% (except in case of meter numbers)</td></tr>
</table>
<input type="hidden" name="byear" value="<%=byear%>">
<input type="hidden" name="utilityid" value="<%=utilityid%>">
<input type="hidden" name="bperiod" value="<%=bperiod%>">
<input type="hidden" name="building" value="<%=building%>">
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="buildingname" value="<%=buildingname%>">
<input type="hidden" name="yscroll" value="0">
</form>

</body>
</html>
