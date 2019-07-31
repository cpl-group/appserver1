<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
function getNumber(number)
'	response.write "|"&number&"|"
	if not(isNumeric(number)) then number = 0
	getNumber = number
end function

dim bperiod, building, byear,demo
bperiod = request("bperiod")
building = request("building")
byear = request("byear")
demo = request("demo")

if demo = "" then demo = false end if

dim rst1, rst2, cnn1
set rst1 = server.createobject("ADODB.Recordset")
set rst2 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
cnn1.open getConnect(0,building,"billing")
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

rst1.open "SELECT distinct *, datestart-1 as datestart, datediff(day, datestart-1, dateend) as days from  rpt_bill_summary WHERE bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod&" ORDER BY TenantName", cnn1
%>
<html><head><title>Bill Summary</title></head><body bgcolor="#FFFFFF">
<basefont size="1" face="Arial">
<%
dim bldgOnPeak, bldgOffPeak, bldgTotalPeak, bldgTotalKW, bldgAdmin, bldgService, bldgCredit, bldgSubtotal, bldgTax, bldgTotalAmt
dim totaldemand_PC, totalOnpeak, totalOffPeak, totalKWH, meterdemandtemp, subsubtotal

if not rst1.eof and trim(request("noheader"))="" then%>
<a href="http://pdfmaker.genergyonline.com/pdfMaker/pdfBillSummary.asp?demo=<%=demo%>&building=<%=building%>&byear=<%=byear%>&bperiod=<%=bperiod%>&strt=<%=server.urlencode(rst1("Strt"))%>" target="_blank">Download printable PDF of Bill Summary</a>
<table width="100%" border="0" bgcolor="#FFFFFF"><tr>
<td height="68"><img src="http://appserver1.genergy.com/eri_th/pdfMaker/invoice_logo_1.jpg" hspace="0" width="202" height="143"></td>
<td width="90%" valign="top" align="center"><b><%if demo then%>Demo Property<%else%><%=rst1("Strt")%><%end if%></b><br>Submetering Summary Report</td>
<td width="10%" valign="bottom">&nbsp;<br>&nbsp;<table border="1" cellspacing="0" cellpadding="3" bordercolor="#000000">
<tr><td align="center"><font size="1">Bill&nbsp;Year</font></td><td align="center"><font size="1">Bill&nbsp;Period</font></td></tr>
<tr><td align="center"><font size="1"><%=byear%></font></td><td align="center"><font size="1"><%=bperiod%></font></td></tr>
</table></td></tr></table>
<%end if

dim pagepart
pagepart = 1
do until rst1.eof
	totaldemand_PC = 0
	totalOnpeak = 0
	totalOffPeak = 0
	totalKWH = 0
	rst2.open "SELECT *, isnull(c.estimated,0) as cest, isnull(p.estimated,0) as pest FROM tblmetersbyperiod mbp LEFT JOIN (SELECT estimated, meterid, billyear, billperiod FROM consumption) c ON c.meterid=mbp.meterid and c.billyear=mbp.billyear and c.billperiod=mbp.billperiod LEFT JOIN (SELECT estimated, meterid, billyear, billperiod FROM Peakdemand) p ON p.meterid=mbp.meterid and p.billyear=mbp.billyear and p.billperiod=mbp.billperiod WHERE leaseutilityid="&cINT(rst1("leaseutilityid"))&" and ypid="&cINT(rst1("ypid")), cnn1
	%>
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<tr><td colspan="2" bgcolor="#3399CC"><font color="#FFFFFF"><b><%if demo then%>Demo Tenant<%else%><%=rst1("billingname")%><%end if%> (<%=rst1("TenantNum")%>)</b></font></td></tr>
	<tr><td width="80%" valign="top"><!-- outer table -->
	
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
	<%
	dim metercount, intpeak, meterEst
	metercount = 0
  meterEst = false
	do until rst2.eof
    if cint(rst2("pest")) or cint(rst2("cest")) then meterEst = True
		metercount = metercount+1
		meterdemandtemp = rst2("Demand_P")
		intpeak = 0
		if isnumeric(rst2("IntPeak")) then intpeak = cdbl(rst2("IntPeak"))
		if rst2("coincident")="True" then
			meterdemandtemp = 0
			totaldemand_PC = rst2("Demand_C")
		else
			totaldemand_PC = totaldemand_PC + cdbl(meterdemandtemp)
		end if
		totalOnpeak = totalOnpeak + formatnumber(cdbl(rst2("OnPeak"))+IntPeak,0)
		totalOffPeak = totalOffPeak + formatnumber(rst2("OffPeak"),0)
		totalKWH = totalKWH + formatnumber(rst2("KWHUsed"),0)
		%>
		<tr>
			<td><font size="1"><%=rst2("Meternum")%></font></td>
			<td><font size="1"><%=rst2("manualMultiplier")%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawPrevious"),0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawCurrent"),0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(cdbl(rst2("OnPeak"))+IntPeak,0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("OffPeak"),0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("KWHUsed"),0)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%if rst2("coincident")="True" then%>0<%else%><%=formatnumber(meterdemandtemp,2)%><%end if%>&nbsp;&nbsp;</font></td>
		</tr>
		
<%	if (metercount>24 or (metercount>11 and pagepart>1)) and trim(request("noheader"))<>"" then
			pagepart = 1
			metercount=1%>
			<tr><td colspan="8" align="right"><font size="1">Continued on next page...</font></td></tr>
			</table>
			</tr></table>
			<WxPrinter PageBreak>
			<table cellspacing="1" cellpadding="2" border="0" width="100%">
			<tr><td colspan="2" bgcolor="#3399CC"><font size="2" color="#FFFFFF"><b><%if demo then%>Demo Tenant<%else%><%=rst1("billingname")%><%end if%> (<%=rst1("TenantNum")%>)</b></font></td></tr>
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
		<td><font size="1"><b><%=formatnumber(totalOffPeak,0)%>&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b><%=formatnumber(totalKWH,0)%>&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b><%=formatnumber(totaldemand_PC,2)%>&nbsp;&nbsp;</b></font></td>
	</tr>
	</table>
  <font size="1">&nbsp;</font>
	<%
		bldgOnPeak = bldgOnPeak + totalOnPeak
		bldgOffPeak = bldgOffPeak + totalOffPeak
		bldgTotalPeak = bldgTotalPeak + totalKWH
		bldgTotalKW = bldgTotalKW + totaldemand_PC
		bldgAdmin = bldgAdmin + (cDbl(rst1("energy"))+cDBL(rst1("demand")))*cdbl(rst1("adminfee"))
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
		<td align="center"><font size="1">Modify Rate</font></td>
		<td align="center"><font size="1">Service Fee</font></td>
		<td align="center"><font size="1">Admin Fee</font></td>
		<td align="center"><font size="1">SqFt</font></td>
		<td align="center"><font size="1">Watts/SqFt</font></td>
	</tr>
	<tr>
		<td><font size="1"><%=rst1("RateTenant")%><%if lcase(trim(rst1("RateTenant")))="avg" then response.write "&nbsp;&nbsp;"&formatnumber(rst1("AvgKWH"),4)%></font></td>
		<td align="center"><font size="1"><%=formatpercent(rst1("AdminFee"),2)%></font></td>
		<td align="center"><font size="1"><%=getNumber(rst1("fuelAdj"))%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(rst1("demand"),2)%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(rst1("energy"),2)%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(rst1("RateModify"),4)%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(cDbl(rst1("serviceFee")))%></font></td>
		<td align="center"><font size="1"><%=formatcurrency((cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDbl(rst1("credit")))*cdbl(rst1("adminfee")),2)%></font></td><!--[demand]-[text210])*[adminfee] -->
		<td align="center"><font size="1"><%=getNumber(rst1("sqft"))%></font></td>
		<td align="center"><font size="1"><%if getNumber(rst1("sqft"))=0 then%>0<%else response.write formatnumber((totaldemand_PC*1000)/cDbl(rst1("sqft"))) end if%></font></td>
	</tr>
	</table>

	</td><td width="20%" valign="top">
<table cellspacing="0" cellpadding="3" border="1" bordercolor="black">
	<tr>
		<td><font size="1">From</font></td>
		<td><font size="1">To</font></td>
		<td><font size="1">No. Days</font></td>
	</tr><tr>
		<td><font size="1"><%=rst1("DateStart")%></font></td>
		<td><font size="1"><%=rst1("DateEnd")%></font></td>
		<td><font size="1"><%=rst1("days")%></font></td>
	</tr>
	</table>
	<table cellspacing="1" cellpadding="2" border="0">
	<tr align="right"><td><font size="1">Sub Total:</font></td><td><font size="1"><%subsubtotal = subsubtotal + cDbl(rst1("energy"))+cDBL(rst1("demand"))%><%=formatcurrency(cDbl(rst1("energy"))+cDBL(rst1("demand")),2)%></font></td></tr>
	<tr align="right"><td><font size="1">Credit:</font></td><td><font size="1"><%=formatcurrency(rst1("credit"),2)%></font></td></tr>
	<tr align="right"><td><font size="1">Admin/Service Fee:</font></td><td><font size="1"><%=formatcurrency(cDbl(rst1("AdminFee")*(cDbl(rst1("energy"))+cDBL(rst1("demand"))))+cDbl(rst1("serviceFee")))%></font></td></tr>
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
<tr><td colspan="3" bgcolor="#3399CC"><font color="#FFFFFF"><b>Building Totals</b></font></td></tr>
<tr><td colspan="3"><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p></td></tr>
<tr><td align="center" valign="top">
	<table border="1" cellspacing="0" cellpadding="5" bordercolor="black" width="250"><tr><td><div align="center">Utility Expenses</div>
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
		<tr><td><font size="1">Average Cost</font></td><td><font size="1"><%if isnumeric(rst1("TotalBillAmt")) and clng(rst1("TotalKWH"))<>0 then response.write formatcurrency(rst1("TotalBillAmt")/rst1("TotalKWH"),4)%></font></td></tr>
		</table>
	</td></tr></table>
</td><td>&nbsp;</td><td align="center" valign="top">
	<table border="1" cellspacing="0" cellpadding="5" bordercolor="black" width="250"><tr><td><div align="center">Sub-Meter Revenue</div>
		<table width="100%">
		<tr><td><font size="1">On Peak</font></td><td><font size="1"><%if isnumeric(bldgOnPeak) then response.write formatnumber(bldgOnPeak,0)%></font></td></tr>
		<tr><td><font size="1">Off Peak</font></td><td><font size="1"><%if isnumeric(bldgOffPeak) then response.write formatnumber(bldgOffPeak,0)%></font></td></tr>
		<tr><td><font size="1">Total KWH</font></td><td><font size="1"><%if isnumeric(bldgTotalPeak) then response.write formatnumber(bldgTotalPeak,0)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<tr><td><font size="1">Total KW</font></td><td><font size="1"><%if isnumeric(bldgTotalKW) then response.write formatnumber(bldgTotalKW,0)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%if isnumeric(subsubtotal) then response.write formatcurrency(subsubtotal)%></font></td></tr>
		<tr><td><font size="1">Admin Fee</font></td><td><font size="1"><%if isnumeric(bldgAdmin) then response.write formatcurrency(bldgAdmin)%></font></td></tr>
		<tr><td><font size="1">Service Fee</font></td><td><font size="1"><%if isnumeric(bldgService) then response.write formatcurrency(bldgService)%></font></td></tr>
		<tr><td><font size="1">Credit</font></td><td><font size="1"><%if isnumeric(bldgCredit) then response.write "("&formatcurrency(bldgCredit)&")"%></font></td></tr>
		<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%if isnumeric(bldgSubtotal) then response.write formatcurrency(bldgSubtotal)%></font></td></tr>
		<tr><td><font size="1">Tax</font></td><td><font size="1"><%if isnumeric(bldgTax) then response.write formatcurrency(bldgTax,2)%></font></td></tr>
		<tr><td><font size="1">Total</font></td><td><font size="1"><%if isnumeric(bldgTotalAmt) then response.write formatcurrency(bldgTotalAmt,2)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<tr><td><font size="1">% Recup</font></td><td><font size="1"><%if isnumeric(rst1("TotalBillAmt")) and isnumeric(bldgTotalAmt) and trim(rst1("TotalBillAmt"))<>"0" then response.write formatpercent(bldgTotalAmt/cdbl(rst1("TotalBillAmt")),2)%></font></td></tr>
		</table>
	</td></tr></table>
</td></tr></table>


<%end if%>
