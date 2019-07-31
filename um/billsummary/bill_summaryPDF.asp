<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
function getNumber(number)
'	response.write "|"&number&"|"
	if not(isNumeric(number)) then number = 0
	getNumber = number
end function

dim bperiod, building, byear
bperiod = request.querystring("bperiod")
building = request.querystring("building")
byear = request.querystring("byear")

dim rst1, rst2, cnn1
set rst1 = server.createobject("ADODB.Recordset")
set rst2 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
cnn1.open getConnect(0,building,"billing")

rst1.open "SELECT *, datediff(day, datestart, dateend) as days from  rpt_bill_summary WHERE bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod&" ORDER BY TenantName", cnn1

%>
<html>
<head>
<title>Bill Summary</title>
<style type="text/css">
body,td,p { font-family:Helvetica,Arial,sans-serif;font-size:8pt; }
</style>
</head>
<body style="font-family:arial, helvetica, sans-serif; font-size:12px">
<%
dim totaldemand_PC, totalOnpeak, totalOffPeak, totalKWH, meterdemandtemp
if not rst1.eof then
%>

<table width="640" cellpadding="0" cellspacing="1" border="0">
<tr>
	<td width="90%" valign="middle"><font size="+1"><b><%=rst1("Strt")%>: Period <%=bperiod%>, Year <%=byear%></b></font></td>
	<td width="10%" valign="top"><table cellspacing="1" cellpadding="3" border="0" width="200">
		<tr bgcolor="#eeeeee" align="center">
			<td>From</td>
			<td>To</td>
			<td>No. Days</td>
		</tr>
		<tr align="center">
			<td><%=rst1("DateStart")%></td>
			<td><%=rst1("DateEnd")%></td>
			<td><%=rst1("days")%></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
end if
do until rst1.eof
	totaldemand_PC = 0
	totalOnpeak = 0
	totalOffPeak = 0
	totalKWH = 0
	rst2.open "SELECT * FROM tblmetersbyperiod WHERE leaseutilityid="&cINT(rst1("leaseutilityid"))&" and ypid="&cINT(rst1("ypid")), cnn1
	%>
  <table cellpadding="0" cellspacing="0" border="0">
	<tr><td colspan="2" valign="top"><!-- outer table -->
	
  <table width="640" cellpadding="5" cellspacing="2" border="0">
  <tr><td colspan="8" bgcolor="#3399cc"><font color="#ffffff"><b><%=rst1("billingname")%> (<%=rst1("TenantNum")%>)</b></font></td></tr>
	<tr bgcolor="#dddddd">
		<td colspan="2" align="center">Tenant</td>
		<td colspan="2" align="center">Readings</td>
		<td colspan="3" align="center">Consumption</td>
		<td align="center">Demand</td>
	</tr>
	<tr bgcolor="#eeeeee" align="center">
		<td>Meter No.</td>
		<td>Multi.</td>
		<td>Previous</td>
		<td>Current</td>
		<td>On Peak</td>
		<td>Off Peak</td>
		<td>KWHR</td>
		<td>KW</td>
	</tr>
	<%
	dim metercount
	metercount = 0
	do until rst2.eof
		metercount = metercount+1
		meterdemandtemp = rst2("Demand_P")
		if cint(meterdemandtemp)=0 then meterdemandtemp = rst2("Demand_C")
		totaldemand_PC = totaldemand_PC + cdbl(meterdemandtemp)
		totalOnpeak = totalOnpeak + formatnumber(rst2("OnPeak"),0)
		totalOffPeak = totalOffPeak + formatnumber(rst2("OffPeak"),0)
		totalKWH = totalKWH + formatnumber(rst2("KWHUsed"),0)
		%>
		<tr>
			<td><%=rst2("Meternum")%></td>
			<td align="right"><%=rst2("Multiplier")%>&nbsp;&nbsp;</td>
			<td align="right"><%=formatnumber(rst2("PrevKWH"),0)%>&nbsp;&nbsp;</td>
			<td align="right"><%=formatnumber(rst2("CurrentKWH"),0)%>&nbsp;&nbsp;</td>
			<td align="right"><%=formatnumber(rst2("OnPeak"),0)%>&nbsp;&nbsp;</td>
			<td align="right"><%=formatnumber(rst2("OffPeak"),0)%>&nbsp;&nbsp;</td>
			<td align="right"><%=formatnumber(rst2("KWHUsed"),0)%>&nbsp;&nbsp;</td>
			<td align="right"><%=formatnumber(meterdemandtemp,2)%>&nbsp;&nbsp;</td>
		</tr>
<%		rst2.movenext
	loop%>
	<tr>
		<td></td>
		<td></td>
		<td></td>
		<td align="right"><b>Meter Totals</b></td>
		<td align="right"><b><%=formatnumber(totalOnPeak,0)%>&nbsp;&nbsp;</b></td>
		<td align="right"><b><%=formatnumber(totalOffPeak,0)%>&nbsp;&nbsp;</b></td>
		<td align="right"><b><%=formatnumber(totalKWH,0)%>&nbsp;&nbsp;</b></td>
		<td align="right"><b><%=formatnumber(totaldemand_PC,2)%>&nbsp;&nbsp;</b></td>
	</tr>
	</table>
	<%
	rst2.close
	%>
	&nbsp;

	</td>
	</tr>
	<tr valign="top" bgcolor="#eeeeee">
	<td valign="middle">
	<table cellpadding="3" cellspacing="1" border="0">
	<tr>
		<td>Service Class</td>
		<td><%=rst1("RateTenant")%></td>
		<td rowspan="5" width="30">&nbsp;</td>
		<td>Admin Fee</td>
		<td align="right"><%=formatpercent(rst1("AdminFee"),2)%></td>
	</tr>
	<tr>
		<td>Fuel Factor</td>
		<td align="right"><%=getNumber(rst1("fuelAdj"))%></td>
		<td>Demand Charge</td>
		<td align="right"><%=formatcurrency(rst1("demand"),2)%></td>
	</tr>
	<tr>
		<td>Energy Charge</td>
		<td align="right"><%=formatcurrency(rst1("energy"),2)%></td>
		<td>Modify Rate</td>
		<td align="right"><%=formatcurrency(rst1("RateModify"),4)%></td>
	</tr>
	<tr>
		<td>Service Fee</td>
		<td align="right"><%=formatcurrency(cDbl(rst1("serviceFee"))*metercount)%></td>
		<td>Admin Fee</td>
		<td align="right"><%=formatcurrency((cDbl(rst1("energy"))+cDbl(rst1("demand")))*cDbl(rst1("AdminFee")),2)%></td><!--[demand]-[text210])*[adminfee] -->
	</tr>
	<tr>
		<td>SqFt</td>
		<td align="right"><%=getNumber(rst1("sqft"))%></td>
		<td>Watts/SqFt</td>
		<td align="right"><%if getNumber(rst1("sqft"))=0 then%>0<%else response.write formatnumber((totaldemand_PC*1000)/cDbl(rst1("sqft"))) end if%></td>
	</tr>
	</table>
	</td>
	<td valign="top" align="right"><!-- outer table -->

	<table cellpadding="3" cellspacing="1" border="0">
	<tr><td height="6"></td></tr>
	<tr align="right"><td>Sub Total:</td><td><%=formatcurrency(cDbl(rst1("energy"))+cDBL(rst1("demand")),2)%></td></tr>
	<tr align="right"><td>Credit:</td><td><%=formatcurrency(rst1("credit"),2)%></td></tr>
	<tr align="right"><td>Admin/Service Fee:</td><td><%=formatcurrency((cDbl(rst1("credit"))+formatcurrency(rst1("subtotal"),2))-formatcurrency(cDbl(rst1("energy"))+cDBL(rst1("demand")),2),2)%></td></tr>
	<tr align="right"><td>Sub Total:</td><td><%=formatcurrency(rst1("subtotal"),2)%></td></tr>
	<tr align="right"><td>Sales Tax:</td><td><%=formatcurrency(cDbl(rst1("subtotal"))*cDbl(rst1("salestax")),2)%></td></tr>
	<tr align="right"><td><b>Total Charges:</b></td><td><b><%=formatcurrency(cDbl(rst1("TotalAmt")))%></b></td></tr>
	</table>
	
	</td>
	</tr></table><!-- outer table -->
	&nbsp;
	<%
	
	rst1.movenext
loop
rst1.close
%>
