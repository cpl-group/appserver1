<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<head>
<title>Invoice</title>
<link rel="Stylesheet" href="styles.css" type="text/css">
<style type="text/css">
h6 { font-weight:normal; }
</style>
</head>
<body bgcolor="#FFFFFF">
<basefont face="Arial">
<%
dim pdfsession
pdfsession = request("pdf")
if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) or ( pdfsession ="yes" ) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if


dim leaseid, ypid, building, pid, byear, bperiod, groupname, utilityid,hidedemand, reject, billingid
leaseid = trim(Request("l"))
billingid = trim(Request("billingid"))
ypid = trim(request("y"))
building = trim(request("building"))
pid = trim(request("pid"))
utilityid = trim(request("utilityid"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))

groupname = ""
if trim(request("reject"))="1" then reject = 1 else reject = 0

dim cnn1, rst,rst1, rst2, rst3, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(building)
dim DBmainIP
DBmainIP = ""

dim eTotal, sTotal, cTotal, gTotal, wTotal, tenantnum, strt, DateEnd, BillingName, billtype, scTotal
dim eTotalTEN, sTotalTEN, cTotalTEN, scTotalTEN, gTotalTEN, wTotalTEN
if trim(request("billid"))<>"" and trim(leaseid)<>"" then
	makeInvoice leaseid, utilityid, trim(request("billid"))
elseif trim(leaseid)<>"" and trim(leaseid)<>"0" and trim(byear)<>"" and trim(bperiod)<>"" then
  eTotal = 0
  gTotal = 0
  wTotal = 0
  sTotal = 0
  cTotal = 0
  scTotal = 0
  sql = "SELECT bp.id as billid, * FROM tblleasesutilityprices lup, billyrperiod byp, tblleases l, tblbillbyPeriod bp WHERE bp.reject=0 and bp.leaseutilityid=lup.leaseutilityid and bp.ypid=byp.ypid and lup.billingid=l.billingid and l.bldgnum=byp.bldgnum and byp.billperiod="&bperiod&" and byp.billyear="&byear&" and lup.leaseutilityid="&leaseid
  rst1.open sql, cnn1
	
  do until rst1.eof
  	makeInvoice leaseid, utilityid, rst1("billid")
    rst1.movenext
  loop
  rst1.close

elseif billingid<>"0" and billingid<>"" then
	sql = "SELECT lup.utility, lup.leaseutilityid, bp.id as billid FROM tblleasesutilityprices lup, tblbillbyPeriod bp where reject=0 and bp.leaseutilityid=lup.leaseutilityid and billyear="&byear&" and billperiod="&bperiod&" and billingid="&billingid
	rst1.open sql, getConnect(pid,building,"Billing")

	do until rst1.eof
		leaseid = rst1("leaseutilityid")
		utilityid = rst1("utility")
	  	makeInvoice leaseid, utilityid, rst1("billid")
		rst1.movenext
		response.write "<WxPrinter PageBreak>"
	loop
	header strt,DateEnd,TenantNum,"Total Tenant ",BillingName,0, bperiod&"/"&byear
	summary bperiod, byear, eTotal, sTotal, cTotal, scTotal, gTotal, wTotal
elseif trim(building)<>"" and trim(byear)<>"" and trim(bperiod)<>"" then
  sql = "SELECT * FROM [group] WHERE type=2 and bldgnum='"&building&"'"
  rst1.open sql, getConnect(pid,0,"billing")

  do until rst1.eof
  	if trim(groupname)<>"" then groupname = groupname&"|"
  	groupname = groupname & rst1("groupname")
  	rst1.movenext
  loop
  dim grouplist
  grouplist = split(groupname,"|")
  rst1.close
  eTotal = 0
  sTotal = 0
  cTotal = 0
  scTotal= 0
  gTotal= 0
  wTotal = 0
  for each groupname in grouplist
  	tenantnum = ""
  	strt = ""
  	DateEnd = ""
  	BillingName = ""
  	billtype = 0
  	sql = "SELECT distinct bp.id as billid, g.LeaseUtilityId, g.utility, bp.ypid, id, lup.billingid FROM ["&groupname&"] g, tblleasesutilityprices lup, tblbillbyPeriod bp where g.LeaseUtilityId=lup.LeaseUtilityId and reject=0 and bp.leaseutilityid=g.leaseutilityid and billyear="&byear&" and billperiod="&bperiod
    if utilityid<>"0" and utilityid<>"" then sql = sql & " and bp.utility="&utilityid
	sql = sql & " ORDER BY billingid"
  	rst1.open sql, cnn1
	dim doTenantSummary
  	do until rst1.eof
		billingid=rst1("billingid")
		doTenantSummary = false
  		makeInvoice cint(rst1("LeaseUtilityId")), cint(rst1("utility")), clng(rst1("billid"))
		response.write "<WxPrinter PageBreak>"
  		rst1.movenext
		if rst1.eof then doTenantSummary = true
		if not rst1.eof then if billingid<>rst1("billingid") then doTenantSummary = true
		if doTenantSummary then
			header strt,DateEnd,TenantNum,"Total Tenant ",BillingName,0, bperiod&"/"&byear
		  	summary bperiod, byear, eTotalTEN, sTotalTEN, cTotalTEN, scTotalTEN, gTotalTEN, wTotalTEN
			response.write "<WxPrinter PageBreak>"
			  eTotalTEN = 0
			  sTotalTEN = 0
			  cTotalTEN = 0
			  scTotalTEN= 0
			  gTotalTEN= 0
			  wTotalTEN = 0
		end if
  	loop
  	rst1.close
	if utilityid="" or utilityid="0" then
		sql = "SELECT bldgname FROM buildings WHERE bldgnum='"&building&"'"
		rst1.open sql, cnn1
	  	'response.write "<hr>totals<br>"
		header "", DateEnd, "", rst1("bldgname")&" Total", rst1("bldgname"), 0, bperiod&"/"&byear
	  	summary bperiod, byear, eTotal, sTotal, cTotal, scTotal, gTotal, wTotal
	  	rst1.close
	end if
  next
end if
%>

</body>

<%
'###################################'
'## makeInvoice                   ##'
'## decides which invoice to make ##'
'###################################'
sub makeInvoice(luid, utype, billid)
  dim billtotal, isonlinebill
  if utype = 3 then 
	sql = "SELECT bbp.fuelAdj as theRealFuelAdj, bbp.btbldgname, bbp.btstrt, isnull(bbp.credit,0) as creditnotnull, bbp.btcity, bbp.btstate, bbp.tstrt, bbp.tcity, bbp.tstate, bbp.tzip, bbp.tName, isnull(totalamt,0) as totalamt2, energydetail, isnull(cust_chrg,0) as customchrgnull,isnull(subtotal,0) as subtotalnull, bbp.*, cbch.*, l.onlinebill FROM tblbillbyperiod bbp INNER JOIN buildings b ON b.bldgnum=bbp.bldgnum INNER JOIN custom_bchbill_water cbch ON cbch.bill_id=bbp.id INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=bbp.leaseutilityid INNER JOIN tblleases l ON l.billingid=lup.billingid WHERE bbp.id="&billid
  elseif utype = 4 then 
	sql = "SELECT bbp.fuelAdj as theRealFuelAdj, bbp.btbldgname, bbp.btstrt, isnull(bbp.credit,0) as creditnotnull, bbp.btcity, bbp.btstate, bbp.tstrt, bbp.tcity, bbp.tstate, bbp.tzip, bbp.tName, isnull(totalamt,0) as totalamt2, energydetail, isnull(cust_chrg,0) as customchrgnull,isnull(subtotal,0) as subtotalnull, bbp.*, cbch.*, l.onlinebill FROM tblbillbyperiod bbp INNER JOIN buildings b ON b.bldgnum=bbp.bldgnum INNER JOIN custom_bchbill_gas cbch ON cbch.bill_id=bbp.id INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=bbp.leaseutilityid INNER JOIN tblleases l ON l.billingid=lup.billingid WHERE bbp.id="&billid
  elseif utype = 10 then
	sql = "SELECT bbp.fuelAdj as theRealFuelAdj, bbp.btbldgname, bbp.btstrt, isnull(bbp.credit,0) as creditnotnull, bbp.btcity, bbp.btstate, bbp.tstrt, bbp.tcity, bbp.tstate, bbp.tzip, bbp.tName, isnull(totalamt,0) as totalamt2, energydetail, 0 as returntempnonull,isnull(cust_chrg,0) as customchrgnull,isnull(subtotal,0) as subtotalnull, total_water_charge, total_sewer_charge, total_Steam_charge, bbp.*, cbch.*, l.onlinebill FROM tblbillbyperiod bbp INNER JOIN buildings b ON b.bldgnum=bbp.bldgnum INNER JOIN custom_bchbill_water_steam  cbch ON cbch.bill_id=bbp.id INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=bbp.leaseutilityid INNER JOIN tblleases l ON l.billingid=lup.billingid WHERE bbp.id="&billid
  elseif utype = 16 then
	sql = "SELECT bbp.fuelAdj as theRealFuelAdj, bbp.btbldgname, bbp.btstrt, isnull(bbp.credit,0) as creditnotnull, bbp.btcity, bbp.btstate, bbp.tstrt, bbp.tcity, bbp.tstate, bbp.tzip, bbp.tName, isnull(totalamt,0) as totalamt2, energydetail, 0 as returntempnonull, 0 as customchrgnull,isnull(subtotal,0) as subtotalnull, bbp.*, l.onlinebill, sc.cooling, sc.heating FROM tblbillbyperiod bbp INNER JOIN custom_bchbill_sc sc ON sc.bill_id=bbp.id INNER JOIN buildings b ON b.bldgnum=bbp.bldgnum INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=bbp.leaseutilityid INNER JOIN tblleases l ON l.billingid=lup.billingid WHERE bbp.id="&billid
  else
	sql = "SELECT bbp.fuelAdj as theRealFuelAdj, bbp.btbldgname, bbp.btstrt, isnull(bbp.credit,0) as creditnotnull, bbp.btcity, bbp.btstate, bbp.tstrt, bbp.tcity, bbp.tstate, bbp.tzip, bbp.tName, isnull(totalamt,0) as totalamt2, energydetail, isnull(returntemp,0) as returntempnonull,isnull(cust_chrg,0) as customchrgnull,isnull(subtotal,0) as subtotalnull, bbp.*, cbch.*, l.onlinebill FROM tblbillbyperiod bbp INNER JOIN buildings b ON b.bldgnum=bbp.bldgnum INNER JOIN custom_bchbill cbch ON cbch.bill_id=bbp.id INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=bbp.leaseutilityid INNER JOIN tblleases l ON l.billingid=lup.billingid WHERE bbp.id="&billid
  end if
'response.write sql'&"<br>"&cnn1
'response.write groupname
'response.end
	rst2.open sql, cnn1,1
  if not rst2.eof then isonlinebill = rst2("onlinebill")
  sql = "SELECT isnull(sum(mbp.used),0) as used, isnull(sum(demand_p),0) as demand_p, isnull(sum(onpeak),0) as onpeak, isnull(sum(offpeak),0) as offpeak, isnull(sum(kwhp),0) as penalty, utilitydisplay, utilityid FROM meters m, "&DBmainIP&"tblutility u, tblleasesutilityprices lup, tblmetersbyperiod mbp WHERE mbp.meterid=m.meterid and lup.leaseutilityid=m.leaseutilityid and lup.utility=u.utilityid and mbp.bill_id="&billid&" group by utilitydisplay, utilityid"
'response.write sql
	rst3.open sql, cnn1
	if not(rst3.eof) and not(rst2.eof) then 
		billtype=rst3("utilitydisplay")
		tenantnum = rst2("TenantNum")
		strt = rst2("tstrt") & "<br>" & rst2("tcity") &", "& rst2("tstate") &" "& rst2("tzip")
		DateEnd = rst2("DateEnd")
		BillingName = rst2("BillingName")
    	utype = cint(rst3("utilityid"))
		header strt,DateEnd,TenantNum,billtype,BillingName,utype, bperiod&"/"&byear
    	billtotal = cdbl(rst2("totalamt2"))
		
	'Get Degree Day Information
	Dim cCDD,pCDD,cHDD,pHDD
	sql = "select billyear, sum(deg) as DD from (select billyear, b.bldgnum, convert(smalldatetime,left(date,11)) as date, dd.region, avg(deg) as deg from deg_day dd full join buildings b on b.region = dd.region inner join billyrperiod bp on bp.bldgnum = b.bldgnum where dd.date between bp.datestart and bp.dateend and deg > 0 and billperiod = "&bperiod&" and b.bldgnum = '"&building&"' group by billyear, b.bldgnum, convert(smalldatetime,left(date,11)), dd.region) a where billyear = "&byear-1&" or billyear = "&byear&" group by billyear"
	rst.open sql, cnn1
	if not rst.eof then 
	while not rst.eof 
		select case rst("billyear")
		case byear 
			cCDD = abs(formatnumber(rst("DD"),1)) & " CDD"
		case byear-1
			pCDD = abs(formatnumber(rst("DD"),1)) & " CDD"
		end select	
	rst.movenext
	wend	
	end if
	rst.close
	if cCDD = "" then 
		cCDD = "NA CDD"
	end if 
	if pCDD = "" then 
		pCDD = "NA CDD"
	end if 
	
	sql = "select billyear, sum(deg) as DD from (select billyear, b.bldgnum, convert(smalldatetime,left(date,11)) as date, dd.region, avg(deg) as deg from deg_day dd full join buildings b on b.region = dd.region inner join billyrperiod bp on bp.bldgnum = b.bldgnum where dd.date between bp.datestart and bp.dateend and deg < 0 and billperiod = "&bperiod&" and b.bldgnum = '"&building&"' group by billyear, b.bldgnum, convert(smalldatetime,left(date,11)), dd.region) a where billyear = "&byear-1&" or billyear = "&byear&" group by billyear"
	
	rst.open sql, cnn1
	if not rst.eof then 
	while not rst.eof 
		select case rst("billyear")
		case byear 
			cHDD = abs(formatnumber(rst("DD"),1)) & " HDD"
		case byear-1
			pHDD = abs(formatnumber(rst("DD"),1)) & " HDD"
		end select	
	rst.movenext
	wend	
	end if
	rst.close
	if cHDD = "" then 
		cHDD = "NA HDD"
	end if 
	if pHDD = "" then 
		pHDD = "NA HDD"
	end if 
    'End Get Degree Day Information
		
if utype=2 then 'electricity
	eTotalTEN = eTotalTEN + billtotal 'tenant summary totals
    eTotal = eTotal + billtotal%>
    <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Usage Charge</td>
      <td align="right"><%=formatcurrency(rst2("energy"))%></td>
    </tr>
<!--     <tr bgcolor="#ffffff">
      <td>Operating Charge</td>
      <td align="right"><%=formatcurrency(0)%></td>
    </tr> -->
	<%if cint(rst2("ratetenant"))<>58 then 'is BCH average cost%>
    <tr bgcolor="#ffffff">
      <td>Capacity Charge</td>
      <td align="right"><%=formatcurrency(rst2("demand"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Hefa Charge</td>
      <td align="right"><%=formatcurrency(rst2("hefaamt"))%></td>
    </tr>
	<%end if%>
	<% if cint(rst2("customchrgnull")) <> 0 then %>
	<tr bgcolor="#ffffff">
      <td>Customer Charge</td>
      <td align="right"><%=formatcurrency(rst2("customchrgnull"))%></td>
    </tr>
	<%end if%>
    </table>
    <table border=0 cellpadding="3" cellspacing="0" align="right">
	<%miscCredits(billid)%>
    <tr>
      <td align="right" colspan=2><b>Subtotal:</b></td>
      <td align="right"><b><%=formatcurrency(rst2("subtotalnull"))%></b></td>
    </tr>
<!--     <tr>
      <td align="right" colspan=2><b>Adjustments:</b></td>
      <td align="right"><b></b></td>
    </tr> -->
    <tr>
      <td align="right" colspan=2><b>Subtotal Adjustments:</b></td>
      <td align="right"><b><%=formatcurrency(0)%></b></td>
    </tr>
    <tr>
      <td align="right" colspan=2><b>Electric Total:</b></td>
      <td align="right"><b><%=formatcurrency(billtotal)%></b></td>
    </tr>
    </table>
	<%showNote()%>
  </td>
</tr>
<tr>
  <td align="left" valign="bottom">
  <table border=0 cellspacing="1" cellpadding="3" bgcolor="#eeeeee">
  <tr bgcolor="#ffffff">
    <td>Demand (KW)</td>
    <td><%=formatnumber(rst3("demand_p"),2)%></td>
  </tr>
  <%if cint(rst2("ratetenant"))=58 then 'is BCH average cost%>
	  <tr bgcolor="#ffffff">
	    <td>Total Usage</td>
	    <td><%=formatnumber(rst2("TEnergy"),2)%></td>
	  </tr>
  <%end if%>
  <tr bgcolor="#ffffff">
    <td>Peak Use (KWH)</td>
    <td><%=formatnumber(rst3("onpeak"),2)%></td>
  </tr>
  <tr bgcolor="#ffffff">
    <td>Off-Peak Use (KWH)</td>
    <td><%=formatnumber(rst3("offpeak"),2)%></td>
  </tr>
  <%if cint(rst2("ratetenant"))=58 then 'is BCH average cost%>
	  <tr bgcolor="#ffffff">
	    <td>Average Cost ($/KWH)</td>
	    <td><%=formatnumber(rst2("avgcost"),6)%></td>
	  </tr>
  <%else%>
	  <tr bgcolor="#ffffff">
	    <td>HEFA Price ($/KWH)</td>
	    <td><%=formatnumber(rst2("hefarate"),6)%></td>
	  </tr>
  <%
  end if
    ShowDD cCDD,cHDD,pCDD,pHDD
  %>
</table>
    <br clear="all">
    </td>
  </tr>
<%elseif utype=6 then 'chilled water
	cTotalTEN = cTotalTEN + billtotal 'tenant summary totals
    cTotal = cTotal + billtotal%>
    <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Usage Charge</td>
      <td align="right"><%=formatcurrency(rst2("energy"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Hefa Charge</td>
      <td align="right"><%=formatcurrency(rst2("hefaamt"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Operating Charge</td>
      <td align="right"><%=formatcurrency(rst2("opamt"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Capacity Charge</td>
      <td align="right"><%if not isnull(rst2("demand")) then response.write formatcurrency(rst2("demand"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Return Temperature Charge</td>
      <td align="right"><%=formatcurrency(rst2("returntempnonull"))%></td>
    </tr>
	<% if cint(rst2("customchrgnull")) <> 0 then %>
	<tr bgcolor="#ffffff">
      <td>Customer Charge</td>
      <td align="right"><%=formatcurrency(rst2("customchrgnull"))%></td>
    </tr>
	<%end if%>
    </table>
    <table border=0 cellpadding="3" cellspacing="0" align="right">
	<%miscCredits(billid)%>
    <tr bgcolor="#ffffff">
      <td align="right" colspan=2><b>Subtotal:</b></td>
      <td align="right"><b><%=formatcurrency(rst2("subtotal"))%></b></td>
    </tr>
<!--     <tr bgcolor="#ffffff">
      <td align="right" colspan=2><b>Adjustments:</b></td>
      <td align="right"><b></b></td>
    </tr> -->
    <tr bgcolor="#ffffff">
      <td align="right" colspan=2><b>Subtotal Adjustments:</b></td>
      <td align="right"><b><%=formatcurrency(0)%></b></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td align="right" colspan=2><b>Chilled Water Total:</b></td>
      <td align="right"><b><%=formatcurrency(billtotal)%></b></td></tr>
    </table>
	<%showNote()%>
    </td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
    <table border=0 cellspacing="1" cellpadding="3" bgcolor="#cccccc">
    <tr bgcolor="#ffffff">
      <td>Demand (Ton)</td>
      <td><%=formatnumber(rst3("demand_p"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Billing Demand (Ton)</td>
      <td><%=formatnumber(rst3("demand_p"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Peak Use (Ton-Hrs)</td>
      <td><%=formatnumber(rst3("onpeak"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Off-Peak Use (Ton-Hrs)</td>
      <td><%=formatnumber(rst3("offpeak"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Penalty (KWH)</td>
      <td><%=formatnumber(rst2("kwhp"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>HEFA Price ($/KWH)</td>
      <td><%=formatnumber(rst2("hefarate"),6)%></td></tr>
    <tr bgcolor="#ffffff">
      <td>Operating Price ($/KW)</td>
      <td><%=formatnumber(rst2("oprate"),6)%></td>
	</tr>
	<%ShowDD cCDD,cHDD,pCDD,pHDD%>	  
    </table>
    </td>
  </tr>

<%  

elseif utype=1 then 'steam
	sTotalTEN = sTotalTEN + billtotal 'tenant summary totals
    sTotal = sTotal + billtotal%>
    <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Usage Charge</td>
      <td align="right"><%=formatcurrency(cdbl(rst2("energy"))+cdbl(rst2("demand")))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Fuel Adjustment</td>
      <td align="right"><%=formatcurrency(rst2("fuel"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Steam Connection Charge</td>
      <td align="right"><%=formatcurrency(rst2("connectioncharge"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Condensate Return Charge</td>
      <td align="right"><%=formatcurrency(rst2("condreturncharge"))%></td>
    </tr>
	<% if cint(rst2("customchrgnull")) <> 0 then %>
	<tr bgcolor="#ffffff">
      <td>Customer Charge</td>
      <td align="right"><%=formatcurrency(rst2("customchrgnull"))%></td>
    </tr>
	<%end if%>
    </table>
    <table border=0 cellpadding="3" cellspacing="0" align="right">
	<%miscCredits(billid)%>
    <tr>
      <td align="right" colspan=2><b>Subtotal:</b></td>
      <td><b><%=formatcurrency(rst2("subtotal"))%></b></td>
    </tr>
<!--     <tr>
      <td align="right" colspan=2><b>Adjustments:</b></td>
      <td><b>0</b></td>
    </tr> -->
    <tr>
      <td align="right" colspan=2><b>Steam Total:</b></td>
      <td><b><%=formatcurrency(billtotal)%></b></td>
    </tr>
    </table>
    <%showNote()%>
    </td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
    <table border="0" cellpadding="3" cellspacing="1" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Actual Demand (M#/Hr)</td>
      <td><%=formatnumber(rst3("demand_p"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Billed Demand (M#/Hr)</td>
      <td><%=formatnumber(rst2("billed_demand"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Usage (M#)</td>
      <td><%=formatnumber(rst3("used"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Cond Return (M#)</td>
      <td><%=formatnumber(rst2("condreturn"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Fuel Adjustment ($/M#)</td>
      <td><%=formatnumber(cdbl(rst2("theRealFuelAdj")),2)%></td>
    </tr>
	<%ShowDD cCDD,cHDD,pCDD,pHDD%>
    </table>
    <br clear="all">
    </td>
  </tr>
<%
elseif utype=3 then 'Cold Water
	wTotalTEN = wTotalTEN + billtotal 'tenant summary totals
    wTotal = wTotal + billtotal%>
    <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Water Charge</td>
      <td align="right"><%=formatcurrency(cdbl(rst2("total_water_charge")))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Sewer Charge</td>
      <td align="right"><%=formatcurrency(rst2("total_sewer_charge"))%></td>
    </tr>
	<% if cint(rst2("customchrgnull")) <> 0 then %>
	<tr bgcolor="#ffffff">
      <td>Customer Charge</td>
      <td align="right"><%=formatcurrency(rst2("customchrgnull"))%></td>
    </tr>
	<%end if%>
    </table>
    <table border=0 cellpadding="3" cellspacing="0" align="right">
	<%miscCredits(billid)%>
    <tr>
      <td align="right" colspan=2><b>Subtotal:</b></td>
      <td><b><%=formatcurrency(rst2("subtotal"))%></b></td>
    </tr>
    <tr>
      <td align="right" colspan=2><b>Water Total:</b></td>
      <td><b><%=formatcurrency(billtotal)%></b></td>
    </tr>
    </table>
    <%showNote()%><br>
    </td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
    <table border="0" cellpadding="3" cellspacing="1" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td align="left">Total CCF</td>
      <td align="right"><%=formatnumber(rst2("total_ccf"),2)/100%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td align="left">Total Gallons</td>
      <td align="right" ><%=formatnumber(rst2("total_gallons"),2)%></td>
    </tr>
	<%ShowDD cCDD,cHDD,pCDD,pHDD%>
    </table>
    <br clear="all">
    </td>
  </tr>
<%
elseif utype=4 then 'Gas
	gTotalTEN = gTotalTEN + billtotal 'tenant summary totals
    gTotal = gTotal + billtotal%>
    <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Gas Delivery Charge</td>
      <td align="right"><%=formatcurrency(cdbl(rst2("total_deliv_chrg")))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Gas Supply Charge</td>
      <td align="right"><%=formatcurrency(rst2("total_supply_chrg"))%></td>
    </tr>
	<% if cint(rst2("customchrgnull")) <> 0 then %>
	<tr bgcolor="#ffffff">
      <td>Customer Charge</td>
      <td align="right"><%=formatcurrency(rst2("customchrgnull"))%></td>
    </tr>
	<%end if%>
    </table>
    <table border=0 cellpadding="3" cellspacing="0" align="right">
	<%miscCredits(billid)%>
    <tr>
      <td align="right" colspan=2><b>Subtotal:</b></td>
      <td><b><%=formatcurrency(rst2("subtotal"))%></b></td>
    </tr>
    <tr>
      <td align="right" colspan=2><b>Gas Total:</b></td>
      <td><b><%=formatcurrency(billtotal)%></b></td>
    </tr>
    </table>
    <%showNote()%>
    </td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
    <table border="0" cellpadding="3" cellspacing="1" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Total CCF</td>
      <td  align="right"><%=formatnumber(rst2("total_ccf"),2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Total Therms</td>
      <td align="right"><%=formatnumber(rst2("total_therms"),2)%></td>
    </tr>
	<%ShowDD cCDD,cHDD,pCDD,pHDD%>
    </table>
    <br clear="all">
    </td>
  </tr>  
<%
elseif utype=10 then 'Hot Water
	wTotalTEN = wTotalTEN + billtotal 'tenant summary totals
    wTotal = wTotal + billtotal%>
    <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Water Charge</td>
      <td align="right"><%=formatcurrency(cdbl(rst2("total_water_charge")))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Sewer Charge</td>
      <td align="right"><%=formatcurrency(rst2("total_sewer_charge"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Hot Water Charge</td>
      <td align="right"><%=formatcurrency(rst2("total_Steam_charge"))%></td>
    </tr>
	<% if cint(rst2("customchrgnull")) <> 0 then %>
	<tr bgcolor="#ffffff">
      <td>Customer Charge</td>
      <td align="right"><%=formatcurrency(rst2("customchrgnull"))%></td>
    </tr>
	<%end if%>
    </table>
    <table border=0 cellpadding="3" cellspacing="0" align="right">
	<%miscCredits(billid)%>
    <tr>
      <td align="right" colspan=2><b>Subtotal:</b></td>
      <td><b><%=formatcurrency(rst2("subtotal"))%></b></td>
    </tr>
    <tr>
      <td align="right" colspan=2><b>Hot Water Total:</b></td>
      <td><b><%=formatcurrency(billtotal)%></b></td>
    </tr>
    </table>
    <%showNote()%>
    </td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
    <table border="0" cellpadding="3" cellspacing="1" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Total CCF</td>
      <td  align="right"><%=formatnumber(cdbl(rst2("total_ccf"))/100,2)%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Total Gallons</td>
      <td  align="right"><%=formatnumber(rst2("total_gallons"),2)%></td>
    </tr>
<%
	'    <tr bgcolor="#ffffff">
   	'   <td>Total Mlbs(steam)</td>
    '  <td align="right"><%=formatnumber(rst2("total_mlbs"),3)</td>
    '</tr>
%>
	<% ShowDD cCDD,cHDD,pCDD,pHDD%>
    </table>
    <br clear="all">
    </td>
  </tr>  
<%elseif utype=16 then 'Space Conditioning (HVAC)
	scTotalTEN = scTotalTEN + billtotal 'tenant summary totals
    scTotal = scTotal + billtotal%>
    <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#ffffff">
      <td>Cooling Charge</td>
      <td align="right"><%=formatcurrency(rst2("cooling"))%></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td>Heating Charge</td>
      <td align="right"><%=formatcurrency(rst2("heating"))%></td>
    </tr>
	<% if cint(rst2("customchrgnull")) <> 0 then %>
	<tr bgcolor="#ffffff">
      <td>Customer Charge</td>
      <td align="right"><%=formatcurrency(rst2("customchrgnull"))%></td>
    </tr>
	<%end if%>
    </table>
    <table border=0 cellpadding="3" cellspacing="0" align="right">
	<%miscCredits(billid)%>
    <tr bgcolor="#ffffff">
      <td align="right" colspan=2><b>Subtotal:</b></td>
      <td align="right"><b><%=formatcurrency(rst2("subtotal"))%></b></td>
    </tr>
    <tr bgcolor="#ffffff">
      <td align="right" colspan=2><b>Space Conditioning Total:</b></td>
      <td align="right"><b><%=formatcurrency(billtotal)%></b></td></tr>
    </table>
	<%showNote()%>
    </td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
    <table border=0 cellspacing="1" cellpadding="3" bgcolor="#cccccc">
	<% ShowDD cCDD,cHDD,pCDD,pHDD%>
    </table>
	    <br clear="all">
    </td>
  </tr>
<%end if%>
  </table>
  <br>
  <!-- end Billing table -->
  </td>
</tr>
<tr><td height="1"></td></tr>
<%
Dim footSpacing
select case utype 
case 3, 4, 100, 10
	hidedemand = "true"
	footSpacing = 3
case 16
	hidedemand = "true"
	footSpacing = 150
case else
	hidedemand = "false"
end select
%>
<%if utype<>16 then%>
<tr>
  <td align="center"><img src="MakeChartyrly.asp?unittype=<%=utype%>&building=<%=building%>&lid=<%=luid%>&by=<%=byear%>&billid=<%=billid%>&bp=<%=bperiod%>&hidedemand=<%=hidedemand%>" width="600" height="175"></td>
</tr>
<%end if%>
<tr>
  <td height="<%=footSpacing%>" valign="bottom">
    </td>
  </tr>
  </table>
  <table border=0 cellpadding="8" cellspacing="1" width="100%" bgcolor="#eeeeee">
  <tr bgcolor="#ffffff" valign="top"><td>
       <%footer rst2("tname"), tenantnum, rst2("tstrt"), rst2("tcity"), rst2("tstate"), rst2("tzip"), rst2("btbldgname"), rst2("btstrt"), rst2("btcity"), rst2("btstate"), rst2("btzip"), isonlinebill%>
</td></tr>
</table>
<center><%if trim(isonlinebill)="True" then%>To view online bill, login to www.genergyonline.com with access code <b><%=tenantnum%>.<%=building%></b>.<%end if%></center>
<!-- end outer table -->

<%
end if
	rst2.close
	rst3.close
end sub
sub header(strt, DateEnd, TenantNum, billtype, BillingName, utype, byperiod)%>
<!-- begin outer table -->
<table border=0 cellpadding="0" cellspacing="0" align="center" width="80%">
<tr><td height="10">&nbsp;</td></tr>	<!-- this height determines the top margin -->
<tr valign="top">
  <td>
  <!-- begin Billing table -->
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr valign="top">
  	<%
	dim spacerHeight
	if cint(utype) = 2  then		'electricity
		spacerHeight = 500
	elseif cint(utype) = 6 then		'chilled water
		spacerHeight = 480
	elseif cint(utype) = 1	then	'steam
		spacerHeight = 480
	elseif cint(utype) = 3	then	'cold water
		spacerHeight = 570
	elseif cint(utype) = 4	then	'gas
		spacerHeight = 570
	elseif cint(utype) = 16	then	'HVAC
		spacerHeight = 570
	elseif cint(utype) = 10	then	'Hot Water
		spacerHeight = 550
	else	'?
		spacerHeight = 570
	end if
	%>
    <td align="center" height="<%=spacerHeight%>"> <!-- this height determines the margin between the totals and the center table -->
    <table border=0 cellpadding="5" cellspacing="0" width="100%">
    <tr>
      <td><img src="invoice-logo-1.jpg" width="202" height="143" border="0"></td>
      <td align="right" valign="middle">
      <table border=0 cellpadding="1" cellspacing="0">
      <tr><td>Date: <%=DateEnd%><br><br></td></tr>
      <tr><td><%if trim(TenantNum)<>"" then%>Customer No. <%=TenantNum%><%end if%></td></tr>
      <tr><td><%=strt%></td></tr>
      </table>
      </td>
    </tr>
    <tr>
      <td colspan="2"><h4><%=BillingName%></h4><b><%if utype<>0 then%><%=billtype%>&nbsp;<%end if%>Billing<%if utype=0 then%>&nbsp;Summary<%end if%></b></td>
    </tr>
    <tr>
      <td colspan="2"><b>Billing Period <%=byperiod%></b></td>
    </tr>
    </table>
    
<%end sub%>


<%sub summary(bperiod, byear, eTotal, sTotal, cTotal, scTotal, gTotal, wTotal)%>
    <table border=0 cellpadding="5" cellspacing="0" width="100%">
    <tr>
      <td>Utility Services Provided For: <%=monthname(bperiod)&" "&byear%><br>See Attached Detail(s)</td>
    </tr>
    </table>
    
    <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#eeeeee">
	<%
	if eTotal>0 then
	%>
    <tr bgcolor="#ffffff">
      <td>Electricity Charge</td>
      <td align="right"><%=formatcurrency(eTotal)%></td>
    </tr>
	<%
	end if
	if sTotal>0 then
	%>
    <tr bgcolor="#ffffff">
      <td>Steam Charge</td>
      <td align="right"><%=formatcurrency(sTotal)%></td>
    </tr>
	<%
	end if
	if gTotal>0 then
	%>
    <tr bgcolor="#ffffff">
      <td>Gas Charge</td>
      <td align="right"><%=formatcurrency(gTotal)%></td>
    </tr>
	<%
	end if
	if cTotal>0 then
	%>
    <tr bgcolor="#ffffff">
      <td>Chilled Water Charge</td>
      <td align="right"><%=formatcurrency(cTotal)%></td>
    </tr>
	<%
	end if
	if scTotal>0 then
	%>
    <tr bgcolor="#ffffff">
      <td>Space Conditioning Charge</td>
      <td align="right"><%=formatcurrency(scTotal)%></td>
    </tr>
	<%
	end if
	if wTotal>0 then
	%>
    <tr bgcolor="#ffffff">
      <td>Cold/Hot Water Charge</td>
      <td align="right"><%=formatcurrency(wTotal)%></td>
    </tr>
	<%
	end if
	if 0>0 then
	%>
    <tr bgcolor="#ffffff">
      <td>Finance Charges</td>
      <td align="right"><b><%=formatcurrency(0)%></b></td>
    </tr>
	<%
	end if
	%>
    </table>
    
    <table border=0 cellpadding="3" cellspacing="0" align="right">
    <tr>
      <td align="right"><b>Total Due:</b></td>
      <td align="right"><b><%=formatcurrency(cdbl(eTotal)+cdbl(sTotal)+cdbl(cTotal)+cdbl(scTotal)+cdbl(gTotal)+cdbl(wTotal))%></b></td></tr>
    </table>
    </td>
  </tr>
  </table>
  <!-- end Billing table -->
  </td>
</tr>
<tr><td height="10">&nbsp;</td></tr>
<tr>
  <td height="125" valign="bottom">
  <table border=0 cellpadding="8" cellspacing="1" width="100%"><tr><td>
     <%'footer rst2("tname"), tenantnum, rst2("tstrt"), rst2("tcity"), rst2("tstate"), rst2("tzip"), rst2("btbldgname"), rst2("btstrt"), rst2("btcity"), rst2("btstate"), rst2("btzip"), "false"%>
 </td>
</tr>
</table>
 </td>
</tr>
</table>
<!-- end outer table -->
<%end sub

sub footer(tenantname, tenantnum, tstrt, tcity, tstate, tzip, btbldgname, billingaddress, btcity, btstate, btzip, isonlinebill)%>
&nbsp;
<table width="80%" border="0" align="center">
	<tr><td colspan="3"><hr width="100%" align="center" noshade size="1"></td></tr>
	<tr>
		<td width="50%" valign="top">Tenant Name and Address:<br>
								<b><%=tenantname%> (<%=tenantnum%>)<br>
	  								<%=replace(tstrt,vbNewLine,"<br>")%><br>
	  								<%=tcity%>, <%=tstate%>&nbsp;<%=tzip%></b>
  		</td>
		<td width="50%" valign="top">Make Check Payable To:<br>
									<b><%=btbldgname%><br>
    								<%if not isnull(billingaddress) then
										response.write (replace(billingaddress,vbNewLine,"<br>"))
									end if%><br>
									<%=btcity%>, <%=btstate%>&nbsp;<%=btzip%><br>
	 								 *Payment due upon receipt</b><br>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="center">&nbsp;<br>If you have questions concerning your bill, please call Colleen Olson of the Real Estate Dept @ 617-355-8309.<br>
			<%if trim(isonlinebill)="True" then%>
	  			To view online bill, login to www.genergyonline.com with <b><%=tenantnum%>.<%=building%></b>.
			<%end if%>
		</td>
	</tr>
</table>
<%end sub

sub miscCredits(billid)
	if cint(rst2("creditnotnull")) <> 0 then
		dim rstMiscCred, credSql
		credSql = "select isnull(description,'Misc Credit') as [desc], credit, convert(integer,adj) as adj FROM tblcreditbyperiod where bill_id="&billid&" and credit<>0 ORDER BY adj"
		set rstMiscCred = server.createobject("adodb.recordset")
		rstMiscCred.open credSql, getLocalConnect(building)
		'response.write credSql
		if not rstMiscCred.eof then
			do while not rstMiscCred.eof
				dim desc
				desc = rstMiscCred("desc")%>	
	
				<tr><td>&nbsp;</td>
					<td align="right"><B><%if rstMiscCred("adj")=1 then%>Adjustment&nbsp;<%else%>Credit&nbsp;<%end if%><%=desc%>:</b></td>
					<td align="right"><B><%=formatcurrency(abs(cdbl(rstMiscCred("credit"))),2)%></b></td>
				</tr>		<%
	
				rstMiscCred.movenext
			loop
		end if
	end if
end sub

sub showNote()
	if trim(rst2("invoice_note"))<>"" then%>
	<br clear="all">
	<table cellpadding="0" cellspacing="0" width="640"><tr><td width="320">
	<table width="320" bgcolor="black" cellpadding="3" cellspacing="1">
	<tr><td bgcolor="white"><%=rst2("invoice_note")%></td></tr>
	</table>
	<td width="320">&nbsp;</td>
	</td></tr></table>
	<%end if
end sub
function ShowDD(cCDD,cHDD,pCDD,pHDD)
%>
	  <tr bgcolor="#eeeeee">
		<td colspan=2><font size="2"><b>Degree Day Summary</b></font></td>
	  </tr>
	  <tr bgcolor="#ffffff" >
		<td align="right"><font size="2">This Period</font></td>
		<td><font size="2"><%=cCDD%> / <%=cHDD%></font></td>
	  </tr>
	  <tr bgcolor="#ffffff">
		<td align="right"><font size="2">Same Period Last Year</font></td>
		<td><font size="2"><%=pCDD%> / <%=pHDD%></font></td>
	  </tr>
<%	 
end function
%>