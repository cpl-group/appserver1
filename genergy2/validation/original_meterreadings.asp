<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim byear, bperiod, meterid, building, isposted, action
meterid = request("meterid")
building = request("building")
byear = request("byear")
bperiod = request("bperiod")
action = request("action")

dim rst1, cnn1, strsql, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(building)

if action="Revert" then
	cnn1.CursorLocation = adUseClient
	'specify stored procedure to run
	cmd.CommandText = "sp_ReviewEdit_RevertReadings"
	cmd.CommandType = adCmdStoredProc
	Set prm = cmd.CreateParameter("meterid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adSmallInt, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm
	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("meterid")	= meterid
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.Parameters("user")		= getKeyValue("user")
	cmd.Execute()
	%><script>
	window.opener.document.location.reload();
	window.close();
	</script>
	<%response.end
end if

strsql = "SELECT m.meternum, c.billyear, c.billperiod, current1, used1, onpeak1, offpeak1, demand1, datepeak1, isnull(bbp.posted,0) as post FROM Consumption c INNER JOIN Meters m ON m.MeterId = c.MeterId INNER JOIN PeakDemand pd ON pd.MeterId = c.MeterId AND c.BillYear = pd.BillYear AND c.BillPeriod = pd.BillPeriod left OUTER JOIN validation v ON v.meterid = c.MeterId AND v.billperiod = c.BillPeriod AND v.billyear = c.BillYear join tblleasesutilityprices tlup on tlup.leaseutilityid=m.leaseutilityid LEFT OUTER JOIN coincidentdemand cd on cd.leaseutilityid=tlup.leaseutilityid and cd.billyear=c.billyear and cd.billperiod=c.billperiod LEFT JOIN tblbillbyperiod bbp ON m.leaseutilityid=bbp.leaseutilityid and bbp.billperiod=c.billperiod and bbp.billyear=c.billyear WHERE c.MeterId="& MeterId &" and c.BillYear="& byear &" and c.BillPeriod="& bperiod &" and tlup.leaseutilityid = m.leaseutilityid"
'response.write strsql
'response.end
rst1.open strsql, cnn1
%>
<html>
<head><title>Original readings</title>
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #dddddd; }
.bordercell { border-left:1px solid #ffffff; border-bottom:1px solid #ffffff; border-top:1px solid #ffffff; }
.bordercolumn { border-left:1px solid #eeeeee; }
</style>
</head>
<body leftmargin="0" topmargin="0"><form>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
	<tr>
		<td bgcolor="#6699cc"><span class="standardheader">
			Original readings for Meter: <%=rst1("meternum")%><br>period <%=rst1("billperiod")%>, <%=rst1("billyear")%></span>
		</td>
	</tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="3">
<tr bgcolor="#dddddd"><td class="bordercell" width="5%"></td>
	<td class="bordercell">Original&nbsp;Reading</td></tr>
<tr><td bgcolor="#dddddd" class="bordercell">Usage</td>
	<td><%=rst1("current1")%></td></tr>
<tr><td bgcolor="#dddddd" class="bordercell">Usage&nbsp;Delta</td>
	<td><%=rst1("used1")%></td></tr>
<tr><td bgcolor="#dddddd" class="bordercell">OnPeak</td>
	<td><%=rst1("onpeak1")%></td></tr>
<tr><td bgcolor="#dddddd" class="bordercell">OffPeak</td>
	<td><%=rst1("offpeak1")%></td></tr>
<tr><td bgcolor="#dddddd" class="bordercell">Demand</td>
	<td><%=rst1("demand1")%></td></tr>
<tr><td bgcolor="#dddddd" class="bordercell">Demand&nbsp;Date&nbsp;Peak</td>
	<td><%=rst1("datepeak1")%></td></tr>
</table>
<%if not rst1("post") then%><center><input type="Submit" name="action" value="Revert" size="15" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;"></center><%end if%>
<input type="hidden" name="meterid" 	value="<%=meterid%>">
<input type="hidden" name="building" 	value="<%=building%>">
<input type="hidden" name="byear" 		value="<%=byear%>">
<input type="hidden" name="bperiod" 	value="<%=bperiod%>">

</form>
</body>
</html>
<%
Function breakDate(strng)
	dim RegularExpressionObject, ReplacedString, RetStr
	Set RegularExpressionObject = New RegExp
	
	With RegularExpressionObject
	.Pattern = " "
	.IgnoreCase = True
	.Global = False
	End With
	if not isnull(rst1("datelastread")) then 
		ReplacedString = RegularExpressionObject.Replace(rst1("datelastread"), " <br>")
		RetStr = ReplacedString
		Set RegularExpressionObject = nothing
		
		breakDate = RetStr
	end if
End Function

%>

