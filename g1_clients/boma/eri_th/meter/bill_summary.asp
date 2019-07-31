<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
function getNumber(number)
'	response.write "|"&number&"|"
	if not(isNumeric(number)) then number = 0
	getNumber = number
end function

dim bperiod, building, byear, pid
bperiod = request("bperiod")
building = request("building")
byear = request("byear")
pid = request("pid")

if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if

dim rst1, rst2, cnn1
set rst1 = server.createobject("ADODB.Recordset")
set rst2 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
cnn1.open getLocalConnect(building)

if trim(bperiod)="" or trim(byear)="" then
	rst1.open "select top 1 BillYear, BillPeriod from tblmetersbyperiod WHERE bldgnum='"&building&"' ORDER BY billyear desc, billperiod desc", cnn1
	if rst1.eof then
		response.write "No information for this building"
		response.end
	else
		byear = cint(rst1("billyear"))
		bperiod = cint(rst1("billperiod"))
	end if
	rst1.close
end if
dim DBlocalIP
if trim(building)<>"" then DBlocalIP = "["&getBuildingIP(building)&"].genergy2.dbo."

dim summarylogo
rst1.open "SELECT logo FROM portfolio p, billtemplates bt WHERE p.templateid=bt.id and p.id=(SELECT portfolioid FROM "&DBlocalIP&"buildings WHERE bldgnum='"&building&"')", application("cnnstr_supermod")
if not rst1.eof then summarylogo = rst1("logo")
rst1.close

rst1.open "SELECT r.unit_credit, b.strt, r.sqft, r.adminfee, isnull(r.RateModify,0) as RateModify, r.fuelAdj, rt.[type], AvgKWH, tenantname, datestart-1 as datestart, datediff(day, datestart-1, dateend) as days, dateend, ypid, r.leaseutilityid, billingname, tenantnum, isnull(credit,0) as credit, isnull(subtotal,0) as subtotal, isnull(energy,0) as energy, isnull(demand,0) as demand, isnull(serviceFee,0) as serviceFee, isnull(tax,0) as tax, isnull(totalamt,0) as totalamt, b.btstrt, lup.calcintpeak,isnull(adminfeedollar,0) as admindollar FROM rpt_bill_summary r, ["&application("superIP")&"].mainModule.dbo.ratetypes rt, buildings b, tblleasesutilityprices lup WHERE lup.leaseutilityid=r.leaseutilityid and b.bldgnum=r.bldgnum AND r.[type]=rt.id and billyear="&byear&" and billperiod="&bperiod&" and r.bldgnum='"&building&"' ORDER BY TenantName", cnn1

%>
<html><head><title>Bill Summary</title></head><body bgcolor="#FFFFFF">
<basefont size="1" face="Arial,Helvetica,sans-serif">
<%
if rst1.eof then
  response.write "This building utility does not have this feature available.(" & application("superIP")&":"&building&":"&byear&":"&bperiod&")"
  response.end
end if

dim bldgOnPeak, bldgOffPeak,bldgintpeak, bldgTotalPeak, bldgTotalKW, bldgAdmin, bldgService, bldgCredit, bldgSubtotal, bldgTax, bldgTotalAmt
dim totaldemand_PC, totaldemand_PCint, totaldemand_PCoff, totalOnpeak, totalOffPeak, totalIntPeak, totalKWH, meterdemandtemp, subsubtotal, meterdemandtempint, meterdemandtempoff

'Zero out building peak numbers
bldgonpeak 	= 0
bldgoffpeak = 0
bldgintpeak = 0

if not rst1.eof and trim(request("noheader"))="" then%>
<table width="100%" border="0" bgcolor="#FFFFFF"><tr>
<td height="68"><img src="http://appserver1.genergy.com/eri_th/pdfMaker/<%=summarylogo%>" hspace="0" width="202" height="143"></td>
<td width="90%" valign="top" align="center"><b><%=rst1("Strt")%></b><br>Submetering Summary Report</td>
<td width="10%" valign="bottom">&nbsp;<br>&nbsp;<table border="0" cellspacing="2" cellpadding="3" bgcolor="#000000">
<tr bgcolor="white"><td align="center"><font size="1">Bill&nbsp;Year</font></td><td align="center"><font size="1">Bill&nbsp;Period</font></td></tr>
<tr bgcolor="white"><td align="center"><font size="1"><%=byear%></font></td><td align="center"><font size="1"><%=bperiod%></font></td></tr>
</table></td></tr></table>
<%end if

dim pagepart, calcintpeak
if rst1("calcintpeak")="True" then calcintpeak = true else calcintpeak = false
pagepart = 1
do until rst1.eof
	totaldemand_PC = 0
	totaldemand_PCint = 0
	totaldemand_PCoff = 0
	totalOnpeak = 0
	totalOffPeak = 0
  	totalIntPeak = 0
	totalKWH = 0
	rst2.open "SELECT * FROM tblmetersbyperiod WHERE leaseutilityid="&cINT(rst1("leaseutilityid"))&" and ypid="&cINT(rst1("ypid")), cnn1
	%>
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<tr><td colspan="2" bgcolor="#3399CC"><font size="2" color="#FFFFFF"><b>Demo Tenant (<%=rst1("TenantNum")%>)</b></font></td></tr>
	<tr><td width="80%" valign="top"><!-- outer table -->
	
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<tr bgcolor="#CCCCCC">
		<td colspan="2" align="center"><font size="1">Tenant</font></td>
		<td colspan="2" align="center"><font size="1">Readings</font></td>
		<td colspan="<%if calcintpeak then%>4<%else%>3<%end if%>" align="center"><font size="1">Consumption</font></td>
		<td align="center" <%if calcintpeak then%>colspan="3"<%end if%>><font size="1">Demand</font></td>
	</tr>
	<tr bgcolor="#CCCCCC" align="center">
		<td><font size="1">Meter No.</font></td>
		<td><font size="1">Multi.</font></td>
		<td><font size="1">Previous</font></td>
		<td><font size="1">Current</font></td>
		<td><font size="1">On Peak</font></td>
    <%if calcintpeak then%><td><font size="1">Int Peak</font></td><%end if%>
		<td><font size="1">Off Peak</font></td>
		<td><font size="1">Total Usage</font></td>
		<td><font size="1"><%if calcintpeak then%>On<%else%>KW<%end if%></font></td>
    <%if calcintpeak then%>
  		<td><font size="1">Off</font></td>
	  	<td><font size="1">Int</font></td>
    <%end if%>
	</tr>
	<%
	dim metercount, intpeak
	metercount = 0
	do until rst2.eof
		metercount = metercount+1
		meterdemandtemp = rst2("Demand_P")
    meterdemandtempint = rst2("Demand_int")
    meterdemandtempoff = rst2("Demand_off")
		intpeak = 0
		if isnumeric(rst2("IntPeak")) then intpeak = cdbl(rst2("IntPeak"))
		if rst2("coincident")="True" then
			meterdemandtemp = 0
			totaldemand_PC = cdbl(rst2("Demand_C")) + cdbl(rst2("Demand_int")) + cdbl(rst2("Demand_off"))
		else
			totaldemand_PC = totaldemand_PC + cdbl(meterdemandtemp)
      if calcintpeak then
        totaldemand_PCoff = totaldemand_PCoff + cdbl(meterdemandtempoff)
        totaldemand_PCint = totaldemand_PCint + cdbl(meterdemandtempint)
      end if
		end if
		
    	totalOffPeak = totalOffPeak + formatnumber(rst2("OffPeak"),2)
		
    if calcintpeak then 
		totalIntPeak 	= totalIntPeak + formatnumber(rst2("IntPeak"),2)
		totalOnPeak 	= totalOnpeak + formatnumber(rst2("OnPeak"),2)
	else
		totalOnpeak = totalOnpeak + formatnumber(cdbl(rst2("OnPeak"))+IntPeak,2)
	end if
		totalKWH = totalKWH + formatnumber(rst2("Used"),2)
		%>
		<tr>
			<td><font size="1"><%=rst2("Meternum")%></font></td>
			<td><font size="1"><%=rst2("ManualMultiplier")%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawPrevious"),0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawCurrent"),0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(cdbl(rst2("OnPeak"))+IntPeak,0)%>&nbsp;&nbsp;</font></td>
      <%if calcintpeak then%><td><font size="1"><%=formatnumber(rst2("IntPeak"),0)%>&nbsp;&nbsp;</font></td><%end if%>
			<td><font size="1"><%=formatnumber(rst2("OffPeak"),0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("Used"),0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%if rst2("coincident")="True" then%>0<%else%><%=formatnumber(meterdemandtemp,2)%><%end if%>&nbsp;&nbsp;</font></td>
      <%if calcintpeak then%>
  			<td><font size="1"><%if rst2("coincident")="True" then%>0<%else%><%=formatnumber(meterdemandtempint,2)%><%end if%>&nbsp;&nbsp;</font></td>
  			<td><font size="1"><%if rst2("coincident")="True" then%>0<%else%><%=formatnumber(meterdemandtempoff,2)%><%end if%>&nbsp;&nbsp;</font></td>
      <%end if%>
		</tr>
		
<%	if (metercount>24 or (metercount>11 and pagepart>1)) and trim(request("noheader"))<>"" then
			pagepart = 1
			metercount=1%>
			<tr><td colspan="8" align="right"><font size="1">Continued on next page...</font></td></tr>
			</table>
			</tr></table>
			<WxPrinter PageBreak>
			<table cellspacing="1" cellpadding="2" border="0" width="100%">
			<tr><td colspan="2" bgcolor="#3399CC"><font size="2" color="#FFFFFF"><b><%=rst1("billingname")%> (<%=rst1("TenantNum")%>)</b></font></td></tr>
			<tr><td width="80%" valign="top">
			<table cellspacing="1" cellpadding="2" border="0" width="100%">
			<tr bgcolor="#CCCCCC">
				<td colspan="2" align="center"><font size="1">Tenant</font></td>
				<td colspan="2" align="center"><font size="1">Readings</font></td>
				<td colspan="3" align="center"><font size="1">Consumption</font></td>
				<td align="center"><font size="1">Demand</font></td>
			</tr>
			<tr bgcolor="#CCCCCC" align="center">
				<td><font size="1">Meter No.</font></td>
				<td><font size="1">Multi.</font></td>
				<td><font size="1">Previous</font></td>
				<td><font size="1">Current</font></td>
				<td><font size="1">On Peak</font></td>
				<td><font size="1">Off Peak</font></td>
				<td><font size="1">KWHR</font></td>
				<td><font size="1">KW</font></td>
			</tr>
<%		end if
		rst2.movenext
	loop
	%>
	<tr>
		<td></td>
		<td></td>
		<td></td>
		<td><font size="1"><b>Meter Totals</b></font></td>
		<td><font size="1"><b><%=formatnumber(totalOnPeak,0)%>&nbsp;&nbsp;</b></font></td>
    <%if calcintpeak then%><td><font size="1"><b><%=formatnumber(totalIntPeak,0)%>&nbsp;&nbsp;</b></font></td><%end if%>
		<td><font size="1"><b><%=formatnumber(totalOffPeak,0)%>&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b><%=formatnumber(totalKWH,0)%>&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b><%=formatnumber(totaldemand_PC,2)%>&nbsp;&nbsp;</b></font></td>
    <%if calcintpeak then%>
  		<td><font size="1"><b><%=formatnumber(totaldemand_PCint,2)%>&nbsp;&nbsp;</b></font></td>
  		<td><font size="1"><b><%=formatnumber(totaldemand_PCoff,2)%>&nbsp;&nbsp;</b></font></td>
    <%end if%>
	</tr>
	</table>
	<%'response.write bldgOnPeak
    bldgOnPeak = bldgOnPeak + totalOnPeak
		if calcintpeak then 
			bldgIntPeak = bldgIntPeak + totalIntPeak 
		end if
		bldgOffPeak = bldgOffPeak + totalOffPeak
		bldgTotalPeak = bldgTotalPeak + totalKWH
		bldgTotalKW = bldgTotalKW + totaldemand_PC
		bldgAdmin = bldgAdmin + (cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDbl(rst1("credit")))*cdbl(rst1("adminfee"))
		bldgService = bldgService + cDbl(rst1("serviceFee"))
		bldgCredit = bldgCredit + cDbl(rst1("credit"))
		bldgSubtotal = bldgSubtotal + cDbl(rst1("subtotal"))
		bldgTax = bldgTax + cDbl(rst1("tax"))
		bldgTotalAmt = bldgTotalAmt + cDbl(rst1("TotalAmt"))
	rst2.close
	%>
	&nbsp;
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<tr bgcolor="#CCCCCC" align="center">
		<td align="center"><font size="1">Service Class</font></td>
		<td align="center"><font size="1">Admin Fee</font></td>
		<td align="center"><font size="1">El. Adj. Factor</font></td>
		<td align="center"><font size="1">Demand Charge</font></td>
		<td align="center"><font size="1">Energy Charge</font></td>
		<td align="center"><font size="1"><%if isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" then%>LMEP Rate<%else%>Modify Rate<%end if%></font></td>
		<td align="center"><font size="1">Service Fee</font></td>
		<td align="center"><font size="1">Admin Fee</font></td>
		<td align="center"><font size="1">SqFt</font></td>
		<td align="center"><font size="1">Watts/SqFt</font></td>
	</tr>
	<tr>
		<td><nobr><font size="1"><%=rst1("type")%><%if ucase(trim(rst1("type")))="AVG COST 1" then response.write "&nbsp;&nbsp;"&formatnumber(rst1("AvgKWH"),6)%></font></nobr></td>
		<td align="center"><font size="1"><%=formatpercent(rst1("AdminFee"),2)%></font></td>
		<td align="center"><font size="1"><%=getNumber(rst1("fuelAdj"))%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(rst1("demand"),2)%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(rst1("energy"),2)%></font></td>
		<td align="center"><font size="1"><%if isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" then%><%=formatcurrency(rst1("unit_credit"),6)%><%else%><%=formatcurrency(rst1("RateModify"),6)%><%end if%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(cDbl(rst1("serviceFee")))%></font></td>
		<td align="center"><font size="1"><%=formatcurrency((cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDBL(rst1("credit")))*cdbl(rst1("adminfee")),2)%></font></td><!--[demand]-[text210])*[adminfee] -->
		<td align="center"><font size="1"><%=getNumber(rst1("sqft"))%></font></td>
		<td align="center"><font size="1"><%if getNumber(rst1("sqft"))=0 then%>0<%else response.write formatnumber((totaldemand_PC*1000)/cDbl(rst1("sqft"))) end if%></font></td>
	</tr>
	</table>

	</td><td width="20%" valign="top">
<table cellspacing="2" cellpadding="3" border="0" bgcolor="#000000">
	<tr bgcolor="white">
		<td><font size="1">From</font></td>
		<td><font size="1">To</font></td>
		<td><font size="1">No. Days</font></td>
	</tr><tr bgcolor="white">
		<td><font size="1"><%=rst1("DateStart")%></font></td>
		<td><font size="1"><%=rst1("DateEnd")%></font></td>
		<td><font size="1"><%=rst1("days")%></font></td>
	</tr>
	</table>
	<table cellspacing="1" cellpadding="2" border="0">
	<tr align="right"><td><font size="1">Sub Total:</font></td><td><font size="1"><%subsubtotal = subsubtotal + cDbl(rst1("energy"))+cDBL(rst1("demand"))%><%=formatcurrency(cDbl(rst1("energy"))+cDBL(rst1("demand")),2)%></font></td></tr>
	<%if pid=49 then%><tr align="right"><td><font size="1">Admin/Service Fee:</font></td><td><font size="1"><%=formatcurrency(cdbl(rst1("servicefee"))+cdbl(rst1("admindollar")),2)%></font></td></tr><%end if%>
	<tr align="right"><td><font size="1"><%if isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" then%>LMEP Credit:<%elseif pid=49 then%>Restructuring Rate Reduction:<%else%>Credit:<%end if%></font></td><td><font size="1">(<%=formatcurrency(rst1("credit"),2)%>)</font></td></tr>
	<%if pid<>49 then'not sjp properties%><tr align="right"><td><font size="1">Admin/Service Fee:</font></td><td><font size="1"><%=formatcurrency(       formatcurrency(rst1("subtotal"),2)+cDbl(rst1("credit"))-formatcurrency(cDbl(rst1("energy"))+cDBL(rst1("demand")),2)      ,2)%></font></td></tr><%end if%>
	<tr align="right"><td><font size="1">Sub Total:</font></td><td><font size="1"><%=formatcurrency(rst1("subtotal"),2)%></font></td></tr>
	<tr align="right"><td><font size="1">Sales Tax:</font></td><td><font size="1"><%=formatcurrency(cDbl(rst1("tax")),2)%></font></td></tr>
	<tr align="right"><td><font size="1"><b>Total Charges:</b></font></td><td><font size="1"><b><%=formatcurrency(cDbl(rst1("TotalAmt")))%></b></font></td></tr>
	</table>
	</td></tr></table>
	&nbsp;
	<%
	rst1.movenext
	pagepart = pagepart+1
	if pagepart>2 or metercount>11 then
		if not rst1.eof then response.write "<WxPrinter PageBreak>&nbsp;"
		pagepart = 1
	elseif trim(request("noheader"))<>"" then
		response.write "&nbsp;<br>&nbsp;<br>"
	end if
loop
rst1.close

rst1.open "select sum(OnPeakKWH) as OnPeakKWH, sum(OffPeakKWH) as OffPeakKWH, sum(TotalKWH) as TotalKWH, sum(CostKWH) as CostKWH, (case when sum(TotalKWH)=0 then 0 else sum(CostKWH)/sum(TotalKWH) end) as UnitCostKWH, sum(TotalKW) as TotalKW, sum(CostKW) as CostKW, (case when sum(TotalKW)=0 then 0 else sum(CostKW)/sum(TotalKW) end) as UnitCostKW, sum(TotalBillAmt) as TotalBillAmt, (case when (sum([totalkw])*24*DateDiff(day,[datestart],[DateEnd]))=0 then 0 else sum([totalkwh])/(sum([totalkw])*24*DateDiff(day,[datestart],[DateEnd])) end) as loadfactor from utilitybill ub INNER JOIN billyrperiod byp ON byp.ypid=ub.ypid WHERE bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod&" GROUP BY datestart, dateend", cnn1
'response.write "select sum(OnPeakKWH) as OnPeakKWH, sum(OffPeakKWH) as OffPeakKWH, sum(TotalKWH) as TotalKWH, sum(CostKWH) as CostKWH, (case when sum(TotalKWH)=0 then 0 else sum(CostKWH)/sum(TotalKWH) end) as UnitCostKWH, sum(TotalKW) as TotalKW, sum(CostKW) as CostKW, (case when sum(TotalKW)=0 then 0 else sum(CostKW)/sum(TotalKW) end) as UnitCostKW, sum(TotalBillAmt) as TotalBillAmt, (case when (sum([totalkw])*24*DateDiff(day,[datestart],[DateEnd]))=0 then 0 else sum([totalkwh])/(sum([totalkw])*24*DateDiff(day,[datestart],[DateEnd])) end) as loadfactor from utilitybill ub INNER JOIN billyrperiod byp ON byp.ypid=ub.ypid WHERE bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod&" GROUP BY datestart, dateend"
'response.end
if not rst1.eof then%>
<WxPrinter PageBreak>
<table cellspacing="1" cellpadding="2" border="0" width="100%">
<tr><td colspan="3" bgcolor="#3399CC"><font size="2" color="#FFFFFF"><b>Building Totals</b></font></td></tr>
<tr><td colspan="3"><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p></td></tr>
<tr><td align="center" valign="top">
	<table border="0" cellspacing="2" cellpadding="5" bgcolor="black" width="250"><tr bgcolor="white"><td><div align="center">Utility Expenses</div>
		<table width="100%">
		<tr><td><font size="1">On Peak KWH</font></td><td><font size="1"><%if isnumeric(rst1("OnPeakKWH")) then response.write formatnumber(rst1("OnPeakKWH"),0)%></font></td></tr>
		<tr><td><font size="1">Off Peak KWH</font></td><td><font size="1"><%if isnumeric(rst1("OffPeakKWH")) then response.write formatnumber(rst1("OffPeakKWH"),0)%></font></td></tr>
		<tr><td><font size="1">Total KWH</font></td><td><font size="1"><%if isnumeric(rst1("TotalKWH")) then response.write formatnumber(rst1("TotalKWH"),0)%></font></td></tr>
		<tr><td><font size="1">Cost KWH</font></td><td><font size="1"><%if isnumeric(rst1("CostKWH")) then response.write formatcurrency(rst1("CostKWH"))%></font></td></tr>
		<tr><td><font size="1">Unit Cost KWH</font></td><td><font size="1"><%if isnumeric(rst1("UnitCostKWH")) then response.write formatcurrency(rst1("UnitCostKWH"),4)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<tr><td><font size="1">Total KW</font></td><td><font size="1"><%if isnumeric(rst1("TotalKW")) then response.write formatnumber(rst1("TotalKW"),0)%></font></td></tr>
		<tr><td><font size="1">Cost KW</font></td><td><font size="1"><%if isnumeric(rst1("CostKW")) then response.write formatcurrency(rst1("CostKW"))%></font></td></tr>
		<tr><td><font size="1">Unit Cost KW</font></td><td><font size="1"><%if isnumeric(rst1("UnitCostKW")) then response.write formatcurrency(rst1("UnitCostKW"),2)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<tr><td><font size="1">Load Factor</font></td><td><font size="1"><%if isnumeric(rst1("loadfactor")) then response.write formatpercent(rst1("loadfactor"),2)%></font></td></tr>
		<tr><td colspan="2">&nbsp;</td></tr>
		<tr><td><font size="1">Utility Bill Amount</font></td><td><font size="1"><%if isnumeric(rst1("TotalBillAmt")) then response.write formatcurrency(rst1("TotalBillAmt"),2)%></font></td></tr>
		<tr><td><font size="1">Average Cost</font></td><td><font size="1"><%if isnumeric(rst1("TotalBillAmt")) and clng(rst1("TotalKWH"))<>0 then response.write formatcurrency(rst1("TotalBillAmt")/rst1("TotalKWH"),6)%></font></td></tr>
		</table>
	</td></tr></table>
</td><td>&nbsp;</td><td align="center" valign="top">
	<table border="0" cellspacing="2" cellpadding="5" bgcolor="black" width="250"><tr bgcolor="white"><td><div align="center">Sub-Meter Revenue</div>
		    <table width="100%">
              <tr>
                <td><font size="1">On Peak</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgOnPeak) then response.write formatnumber(bldgOnPeak,0)%>
                  </font></td>
              </tr>
			  <%if calcintpeak then %>
              <tr>
                <td><font size="1">Int Peak</font></td>
                <td><font size="1"><%if isnumeric(bldgIntPeak) then response.write formatnumber(bldgIntPeak,0)%></font></td>
              </tr>
			  <%end if%>
              <tr>
                <td><font size="1">Off Peak</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgOffPeak) then response.write formatnumber(bldgOffPeak,0)%>
                  </font></td>
              </tr>
              <tr>
                <td><font size="1">Total KWH</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgTotalPeak) then response.write formatnumber(bldgTotalPeak,0)%>
                  </font></td>
              </tr>
              <tr>
                <td colspan="2"><hr size="1" color="#000000" noshade></td>
              </tr>
              <tr>
                <td><font size="1">Total KW
                  <%if calcintpeak then%>
                  On
                  <%end if%>
                  </font></td>
                <td><font size="1">
                  <%if isnumeric(bldgTotalKW) then response.write formatnumber(bldgTotalKW,0)%>
                  </font></td>
              </tr>
              <tr>
                <td colspan="2"><hr size="1" color="#000000" noshade></td>
              </tr>
              <tr>
                <td><font size="1">Subtotal</font></td>
                <td><font size="1">
                  <%if isnumeric(subsubtotal) then response.write formatcurrency(subsubtotal)%>
                  </font></td>
              </tr>
              <tr>
                <td><font size="1">Admin Fee</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgAdmin) then response.write formatcurrency(bldgAdmin)%>
                  </font></td>
              </tr>
              <tr>
                <td><font size="1">Service Fee</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgService) then response.write formatcurrency(bldgService)%>
                  </font></td>
              </tr>
              <tr>
                <td><font size="1">Credit</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgCredit) then response.write "("&formatcurrency(bldgCredit)&")"%>
                  </font></td>
              </tr>
              <tr>
                <td><font size="1">Subtotal</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgSubtotal) then response.write formatcurrency(bldgSubtotal)%>
                  </font></td>
              </tr>
              <tr>
                <td><font size="1">Tax</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgTax) then response.write formatcurrency(bldgTax,2)%>
                  </font></td>
              </tr>
              <tr>
                <td><font size="1">Total</font></td>
                <td><font size="1">
                  <%if isnumeric(bldgTotalAmt) then response.write formatcurrency(bldgTotalAmt,2)%>
                  </font></td>
              </tr>
              <tr>
                <td colspan="2"><hr size="1" color="#000000" noshade></td>
              </tr>
              <tr>
                <td><font size="1">% Recup</font></td>
                <td><font size="1">
                  <%if isnumeric(rst1("TotalBillAmt")) and isnumeric(bldgTotalAmt) and trim(rst1("TotalBillAmt"))<>"0" then response.write formatpercent(bldgTotalAmt/cdbl(rst1("TotalBillAmt")),2)%>
                  </font></td>
              </tr>
            </table>
	</td></tr></table>
</td></tr></table>


<%end if%>
