<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<basefont face="Arial">
<%
dim tenant, byear, bperiod, pid, lid, building, utilityid
lid = trim(Request("lid"))
tenant = trim(Request("tenant"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))
pid = trim(request("pid"))
building = trim(request("building"))
utilityid = trim(request("utilityid"))
dim pdfsession
pdfsession = request("pdf")

if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) or ( pdfsession ="yes" ) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if

dim cnn1, rst1, rst2, bldgrs, sql
set bldgrs = Server.CreateObject("ADODB.Recordset")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(pid,building,"billing")
dim DBmainIP
DBmainIP = ""

dim templid, tempypid, totalout
if trim(request("billid"))<>"" and lid<>"" then
	showtenantbill lid, request("billid")
elseif tenant<>"" and byear<>"" and bperiod<>"" then
	bldgrs.open "SELECT bbp.id as billid, ypid, lup.leaseutilityid, billingid FROM buildings b INNER JOIN tblbillbyperiod bbp ON bbp.bldgnum=b.bldgnum INNER JOIN tblleasesutilityprices lup ON bbp.leaseutilityid=lup.leaseutilityid WHERE bbp.billyear="&byear&" and bbp.billperiod="&bperiod&" and billingid="&tenant&" ORDER BY ypid, lup.leaseutilityid", cnn1
	if bldgrs.eof then
		response.write "No bill available for selected."
		response.end
	end if
	templid = trim(bldgrs("leaseutilityid"))
	tempypid = trim(bldgrs("ypid"))
	showtenantbill templid, bldgrs("billid")
	bldgrs.close
elseif lid<>"" and byear<>"" and bperiod<>"" then
	bldgrs.open "SELECT bbp.id as billid, ypid, billingid FROM buildings b INNER JOIN tblbillbyperiod bbp ON bbp.bldgnum=b.bldgnum INNER JOIN tblleasesutilityprices lup ON bbp.leaseutilityid=lup.leaseutilityid WHERE bbp.billyear="&byear&" and bbp.billperiod="&bperiod&" and lup.leaseutilityid="&lid&" ORDER BY ypid, lup.leaseutilityid", cnn1
	if bldgrs.eof then
		response.write "No bill available for selected."
		response.end
	end if
	templid = lid
	tempypid = trim(bldgrs("ypid"))
	showtenantbill templid, bldgrs("billid")
	bldgrs.close
elseif byear<>"" and bperiod<>"" and building<>"" then
  if isnumeric(utilityid) then sql = sql & " and bbp.utility="&utilityid
  if building="" then
  	sql = "SELECT bbp.id as billid, ypid, lup.leaseutilityid, billingid FROM buildings b INNER JOIN tblbillbyperiod bbp ON bbp.bldgnum=b.bldgnum INNER JOIN tblleasesutilityprices lup ON bbp.leaseutilityid=lup.leaseutilityid WHERE bbp.billyear="&byear&" and bbp.billperiod="&bperiod&" and portfolioid="&pid&" and reject = 0 "&sql&"  ORDER BY ypid, lup.leaseutilityid"
  else
	  sql = "SELECT bbp.id as billid, ypid, lup.leaseutilityid, billingid FROM buildings b INNER JOIN tblbillbyperiod bbp ON bbp.bldgnum=b.bldgnum INNER JOIN tblleasesutilityprices lup ON bbp.leaseutilityid=lup.leaseutilityid WHERE bbp.billyear="&byear&" and bbp.billperiod="&bperiod&" and portfolioid="&pid&" "&sql&" and bbp.bldgnum='"&building&"' and reject=0 ORDER BY ypid, lup.leaseutilityid"
  end if
  bldgrs.open sql, cnn1
	if bldgrs.eof then
		response.write "No bills available for selected."
		response.end
	end if
	do until bldgrs.eof
		templid = trim(bldgrs("leaseutilityid"))
		tempypid = trim(bldgrs("ypid"))
		showtenantbill templid, bldgrs("billid")
		bldgrs.movenext
	loop
	bldgrs.close
end if
set cnn1 = nothing
'showtenantbill 4005, 4
'response.write "SELECT ypid, lup.leaseutilityid, billingid FROM buildings b INNER JOIN tblbillbyperiod bbp ON bbp.bldgnum=b.bldgnum INNER JOIN tblleasesutilityprices lup ON bbp.leaseutilityid=lup.leaseutilityid WHERE bbp.billyear="&byear&" and bbp.billperiod="&bperiod&" and portfolioid="&pid&" ORDER BY ypid, lup.leaseutilityid"
'response.end






'### begin of showtenantbill, is rest of file ###

function showtenantbill(leaseid, billid)
dim cnn2, rst1, rst2, sql, metertitle, totalizernum, billcount, isonlinebill
Set cnn2 = Server.CreateObject("ADODB.Connection")
cnn2.Open getLocalConnect(building)
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
sql = "SELECT isnull(mintons,0) as minton, isnull(description,0) as creditdesc, isnull(actualtonsh,0) as atonsh, isnull(actualtons,0) as atons, bbp.*, c.*, lup.*, co.*, rc.*, l.onlinebill FROM tblbillbyperiod bbp INNER JOIN custom_oucbill c ON bbp.billyear=c.billyear and bbp.billperiod=c.billperiod and bbp.leaseutilityid=c.leaseutilityid INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=c.leaseutilityid and bbp.id = c.bill_id INNER JOIN tblleases l ON l.billingid=lup.billingid INNER JOIN custom_oucAccount co ON co.billingid=lup.billingid LEFT JOIN ratecodes rc ON rc.id=c.rateid WHERE bbp.id="&billid
rst1.open sql, cnn2, 2
dim BillCapBaseRate, BaseChargeBaseRate, AdjChargeBaseRate, BaseChargeBillRate, AdjChargeBillRate, BillCapBillRate, metercount, tenantnum, premiseid, billingname, strt, datestart, dateend, baseamt, CustCharge, credit, MonthCPI, DiffTemp, UPI, mintons, BillingCapacity, tstrt, tcity, tstate, tzip, totalTonHours, totalPeakCap, creditdesc, deduct, deduct_tons, deduct_tonsh, wavgdt,subtotal, billtotal,demandchg,usagechg, invoice_note
'response.write sql&"|"&rst1("demandbilled")
if not rst1.eof then
mintons = rst1("minton")
BillCapBillRate = rst1("demandbilledchg")
BillCapBaseRate = rst1("demandbase")
BaseChargeBaseRate = rst1("ConBbase")
BaseChargeBillRate = rst1("ConBBilled")
AdjChargeBaseRate = rst1("ConAbase")
AdjChargeBillRate = rst1("ConAbilled")
CustCharge = rst1("CustCharge")
metercount = rst1("metercount")
tenantnum = rst1("tenantnum")
premiseid = rst1("premiseid")
billingname = rst1("billingname")
strt = rst1("strt")
tstrt = rst1("tstrt")
tcity = rst1("tcity")
tstate = rst1("tstate")
tzip = rst1("tzip")
datestart = rst1("datestart")
dateend = rst1("dateend")
baseamt = rst1("baseamt")
credit = rst1("credit")
totalTonHours = cdbl(rst1("atonsh"))
totalPeakCap = cdbl(rst1("atons"))
MonthCPI = rst1("CPIadj")
DiffTemp = rst1("DTadj")
UPI = rst1("UPI")
creditdesc = rst1("creditdesc")
deduct = rst1("deduct")
deduct_tons = rst1("deduct_tons")
deduct_tonsh = rst1("deduct_tonsh")
BillingCapacity = rst1("demandbilled")
isonlinebill = rst1("onlinebill")
wavgdt = rst1("wavgdt")
subtotal = rst1("subtotal")
billtotal = rst1("totalamt")
demandchg= rst1("demand")
usagechg = rst1("energy")
invoice_note = rst1("invoice_note")
end if 
rst1.close

if trim(BillCapBaseRate) = "" or isnull(BillCapBaseRate) then BillCapBaseRate = 0
if trim(BillCapBillRate) = "" or isnull(BillCapBillRate) then BillCapBillRate = 0
if trim(BaseChargeBaseRate) = "" or isnull(BaseChargeBaseRate) then BaseChargeBaseRate = 0
if trim(BaseChargeBillRate) = "" or isnull(BaseChargeBillRate) then BaseChargeBillRate = 0
if trim(AdjChargeBaseRate) = "" or isnull(AdjChargeBaseRate) then AdjChargeBaseRate = 0
if trim(AdjChargeBillRate) = "" or isnull(AdjChargeBillRate) then AdjChargeBillRate = 0
if isnumeric(metercount) then 
	if cint(metercount)=0 or cint(metercount)>1 then metertitle = "Totalized"
end if
sql = "SELECT count(*) as billnumbers FROM tblleases WHERE billingid in (SELECT l.billingid FROM tblleasesutilityprices lup, tblleases l WHERE l.billingid=lup.billingid and leaseutilityid="&leaseid&" and l.leaseexpired=0) GROUP BY onlinebill"
rst2.open sql, cnn2
if not rst2.eof then billcount = cint(rst2("billnumbers")) else billcount = 0
rst2.close
sql = "SELECT * FROM tblmetersbyperiod mc LEFT JOIN (SELECT m2.meternum as totalizernum, m1.meterid FROM meters m1 INNER JOIN meters m2 ON m1.refmeterid=m2.meterid) ref ON ref.meterid=mc.meterid WHERE bill_id="&billid&"order by totalizernum desc"
rst2.open sql, cnn2
'response.write sql
'response.end
if not rst2.eof then 
  totalizernum = rst2("totalizernum")
  if metertitle="" then metertitle=rst2("meternum")
  if trim(totalizernum)<>"" and billcount>1 then metertitle = totalizernum
end if
%>
<html><head><title></title>
<link rel="Stylesheet" href="styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<br><br>
<table width="80%" border="0" cellpadding="0" cellspacing="0" align="center"><tr><td valign="top" height="925">
<table border=0 cellpadding="0" cellspacing="0" width="100%">
<tr>
  <td><img src="oucooling.gif" alt="OUCooling / A TRIGEN*CINERGY SOLUTIONS Service" width="203" height="123" border="0"></td>
  <td width="20">&nbsp;</td>
  <td align="right">
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td>Service&nbsp;Period: <nobr><%=datestart%> - <%=dateend%></nobr></td>
    <td width="10">&nbsp;</td>
  </tr>
  <tr>
    <td>Billing&nbsp;Date: <%=date%></td>
    <td width="10">&nbsp;</td>
  </tr>
  <tr><td>&nbsp;</td></td width="10">&nbsp;</td></tr>
  <tr>
    <td><i>Orlando Utilities Commission<br>500 South Orange Ave<br>Orlando FL 32801</i></td>
    <td width="10">&nbsp;</td>
  </tr>
  </table>
  </td>
</tr>
</table>


<br>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#eeeeee">
<tr valign="bottom"><td colspan="4" align="center" bgcolor="#e0e6f2"><b>Chilled Water Billing Summary</b></td></tr>
<tr bgcolor="#ffffff" valign="top">
  <td width="15%">Customer No</td>
  <td width="35%"><%=tenantnum%></td>
  <td width="15%">Premise No</td>
  <td width="35%"><%=premiseid%></td>
</tr>
<tr bgcolor="#ffffff"><td>Customer</td><td><%=billingname%></td><td>Meter No</td><td><%=metertitle%></td></tr>
<tr bgcolor="#ffffff"><td>Address</td><td colspan="3"><%=tstrt%>,&nbsp;<%=tcity%>,&nbsp;<%=tstate%>&nbsp;<%=tzip%></td></tr>
</table>

<br><br>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#eeeeee">
<tr valign="bottom"><td colspan="3" align="center" bgcolor="#e0e6f2"><b>Consumption and Capacity</b></td></tr>
<tr bgcolor="#ffffff"><td width="34%"><%if not hasMultiMeters then%><b>Meter ID</b><%end if%></td><td width="33%"><b>Ton Hours Consumed</b></td><td width="33%"><b>Service Period Peak Capacity In Tons</b></td></tr>
	<%dim hasMultiMeters, meterDisplay
	
	hasMultiMeters = false
	do until rst2.eof
		meterDisplay = meterDisplay & "<tr bgcolor=""#ffffff""><td>"&rst2("meternum")&"&nbsp;</td><td>"&formatnumber(rst2("used"),0)&"&nbsp;</td><td>"&formatnumber(rst2("demand_p"),1)&" &nbsp;</td></tr>"
		rst2.movenext
		if not rst2.eof then hasMultiMeters = true
	loop
	rst2.close
	if not hasMultiMeters then 
		response.write meterDisplay
	else%>
	<tr bgcolor="#ffffff"><td><b>Totalizer</b></td><td><%=formatnumber(totalTonHours,0)%>&nbsp;</td><td><%=formatnumber(totalPeakCap,1)%>&nbsp;</td></tr>
	<%end if%>
	<%if trim(deduct)="True" then%>
    
			<tr bgcolor="#ffffff"><td><b>Deduction</b></td><td><%=formatnumber(deduct_tonsh,0)%>&nbsp;</td><td><%=formatnumber(deduct_tons,1)%>&nbsp;</td></tr>
	  <%end if%>

	<tr bgcolor="#ffffff"><td><b>Total</b></td><td><%=formatnumber(totalTonHours,0)%>&nbsp;</td><td><%=formatnumber(totalPeakCap,1)%>&nbsp;</td></tr>
</table>

	<%if trim(deduct)<>"True" then%><br><br><%end if%>
<%
dim billingcapCharge, baseCharge, adjCharge, total
billingcapCharge = cdbl(baseamt)+0
%>
<table width="100%" border="0" cellpadding="5" cellspacing="0">
<tr valign="bottom"><td colspan="3" align="center" bgcolor="#e0e6f2"><b>Charges For Service</b></td></tr>
	<tr><td>Contracted Cooling Capacity in Tons</td><td align="right"><%=formatnumber(mintons,1)%></td></tr>
	<tr><td>Billing Capacity</td><td align="right"><%=formatnumber(BillingCapacity,1)%></td></tr>
</table>

<br>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#e0e6f2">
<tr bgcolor="#ffffff"><td><b>Description Of Charge</b></td><td><b>Quantity</b></td><td><b>Base Rate</b></td><td><b>Billed Rate</b></td><td><b>Charge</b></td></tr>
<tr bgcolor="#ffffff"><td width="28%">Billing Capacity Charge/Tons</td>
  <td width="18%"><%=formatnumber(BillingCapacity,1)%></td>
  <td width="18%"><%=formatcurrency(BillCapBaseRate,2)%></td>
  <td width="18%"><%=formatcurrency(BillCapBillRate,2)%></td>
  <td width="18%" align="right"><%=formatcurrency(demandchg)%></td>
</tr>
<tr bgcolor="#ffffff"><td colspan="5"><b>Consumption Charges</b></td></tr>
<tr bgcolor="#ffffff"><td>Base Charge Ton Hours</td>
  <td><%=formatnumber(totalTonHours,0)%></td>
  <td><%=formatcurrency(BaseChargeBaseRate,5)%></td>
  <td><%=formatcurrency(BaseChargeBillRate,5)%></td>
  <td align="right"><%baseCharge = (cdbl(BaseChargeBillRate)*totalTonHours)%><%=formatcurrency(baseCharge)%></td>
</tr>
<tr bgcolor="#ffffff"><td>Adj. Charge Ton Hours</td>
  <td><%=formatnumber(totalTonHours,0)%></td>
  <td><%=formatcurrency(AdjChargeBaseRate,5)%></td>
  <td><%=formatcurrency(AdjChargeBillRate,5)%></td>
  <td align="right"><%adjCharge = (cdbl(AdjChargeBillRate)*totalTonHours)%><%=formatcurrency(adjCharge)%></td>
</tr>
<tr bgcolor="#ffffff"><td>Customer Charge</td>
  <td><%=metercount%></td>
  <td></td>
  <td><%=formatcurrency(CustCharge)%></td>
  <td align="right"><%CustCharge=cdbl(metercount) * cdbl(CustCharge)%><%=formatcurrency(CustCharge)%></td>
</tr>
</table>
<br>
<table align="right" border="0" cellpadding="3" cellspacing="0">
	<tr><td align="right">Chilled Water Charges Subtotal:</td><td align="right"><%total = billingcapCharge+baseCharge+adjCharge+CustCharge%><%=formatcurrency(subtotal)%></td></tr>
	<%if cdbl(credit)<>0 then%><tr><td align="right"><%=creditdesc%>:</td><td align="right"><%=formatcurrency(cdbl(credit)*-1)%></td></tr><%end if%>
	<%miscCredits 2, leaseid, billid%>
	<tr><td align="right"><font size="+2"><b>Total Chilled Water Billing:</b></font></td><td align="right"><font size="+2"><b><%=formatcurrency(billtotal)%></b></font></td></tr>
</table>
<br clear="all">
<%showNote(invoice_note)%>&nbsp;<br>
<div align="center"><img src="MakeChartyrly.asp?lid=<%=leaseid%>&by=<%=byear%>&bp=<%=bperiod%>&billid=<%=billid%>&building=<%=building%>&unittype=tons&isOUC=true" width="600" height="175"></div>
</td>
</tr>

<tr>
  <td valign="bottom">
  <table border="0" cellpadding="5" cellspacing="1" width="100%" bgcolor="#e0e6f2">
  <tr bgcolor="#ffffff">
    <td>Monthly CPI Adj: <%=MonthCPI%></td>
    <td>Differential Temperature Adjustment: <%=DiffTemp%></td>
    <td>Weighted Average Delta T: <%=wavgdt%>&deg;</td>
    <td>Electric Price Index: <%=UPI%></td>
  </tr>
  </table>
  <br>
  </td>
</tr></table>
<%if hasMultiMeters then%>
<WxPrinter PageBreak>
<table width="80%" border="0" cellpadding="0" cellspacing="0" align="center"><tr><td valign="top" height="925">
<table border=0 cellpadding="0" cellspacing="0" width="100%">
<tr>
  <td><img src="oucooling.gif" alt="OUCooling / A TRIGEN*CINERGY SOLUTIONS Service" width="203" height="123" border="0"></td>
  <td width="20">&nbsp;</td>
  <td align="right">
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td>Service&nbsp;Period: <nobr><%=datestart%> - <%=dateend%></nobr></td>
    <td width="10">&nbsp;</td>
  </tr>
  <tr>
    <td>Billing&nbsp;Date: <%=date%></td>
    <td width="10">&nbsp;</td>
  </tr>
  <tr><td>&nbsp;</td></td width="10">&nbsp;</td></tr>
  <tr>
    <td><i>Orlando Utilities Commission<br>500 South Orange Ave<br>Orlando FL 32801</i></td>
    <td width="10">&nbsp;</td>
  </tr>
  </table>
  </td>
</tr>
</table>
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#eeeeee">
<tr valign="bottom"><td colspan="3" align="center" bgcolor="#e0e6f2"><b>Consumption and Capacity Meter Break Down</b></td></tr>
<tr bgcolor="#ffffff"><td width="34%"></td><td width="33%"><b>Ton Hours Consumed</b></td><td width="33%"><b>Service Period Peak Capacity In Tons</b></td></tr>
	<%=meterDisplay%>
	<tr bgcolor="#ffffff"><td><b>Total Ton Hours</b></td><td><%=formatnumber(totalTonHours,0)%>&nbsp;</td><td><%=formatnumber(totalPeakCap,1)%>&nbsp;</td></tr>
</table>
</td></tr></table>
<%end if%>
<center><%if trim(isonlinebill)="True" then%>To view online bill, login to www.genergyonline.com with access code <b><%=tenantnum%>.<%=building%></b>.<%end if%></center>
<WxPrinter PageBreak>
<%dim outputstring
outputstring = ""



'response.write "<table width=""100%"" border=""0""><tr><td height=""68""><img src=""invoice-logo-1.jpg"" width=""202"" height=""143""></td></tr></table>"
'response.write outputstring
%>

<!-- </div> -->
<%
set rst1 = nothing
showtenantbill = outputstring
end function


sub showNote(note)
	if trim(note)<>"" then%>
	<br clear="all">
	<table cellpadding="0" cellspacing="0" width="640"><tr><td width="320">
	<table width="320" bgcolor="black" cellpadding="3" cellspacing="1">
	<tr><td bgcolor="white"><%=note%></td></tr>
	</table>
	<td width="320">&nbsp;</td>
	</td></tr></table>
	<%end if
end sub%>
</body></html>
<%
sub miscCredits(numcells, tmpleaseid, billid)
	'if cint(rst2("creditnonull")) <> 0 then
		dim rstMiscCred, credSql
		credSql = "select isnull(description,'Misc Credit') as [desc], credit, convert(integer,adj) as adj FROM tblcreditbyperiod where bill_id="&billid&" and credit<>0 ORDER BY adj"
		set rstMiscCred = server.createobject("adodb.recordset")
		rstMiscCred.open credSql, getLocalConnect(building)
		'response.write credSql
		if not rstMiscCred.eof then
			do while not rstMiscCred.eof
				dim desc
				desc = rstMiscCred("desc")	%>
					<tr>
						<td align="right" nowrap><%if rstMiscCred("adj")=1 then%>Adjustment:<%else%>Credit:<%end if%> <%=desc%></td>
						<td align="right"><%=formatcurrency(cdbl(rstMiscCred("credit")),2)%></td>
					</tr><%
				rstMiscCred.movenext
			loop
		end if
	'end if
end sub%>