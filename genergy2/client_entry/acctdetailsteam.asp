<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<title>Account Detail</title>
<script>
function roundNumber(num,precision)
{	var i,sZeros = '';
	for(i = 0;i < precision;i++)
		sZeros += '0';
	i = Number(1 + sZeros);
	return Math.round(num * i) / i;
}

function calcAvg_DD()
{	frm = document.forms['detail']
	var CDD = parseFloat(frm.CDD.value);
	var HDD = parseFloat(frm.HDD.value);
	if(isNaN(CDD)) CDD=0;
	if(isNaN(HDD)) HDD=0;
	frm.Avg_DD.value = roundNumber(parseFloat(frm.mlb.value)/(CDD + HDD),2);
}

function calcAvgCost()
{	frm = document.forms['detail'];
	frm.AvgCost.value = roundNumber(parseFloat(frm.totalamt.value)/parseFloat(frm.mlb.value),4)
}

</script>
</head>
<%
dim acctid, ypid, by, bp, utility
acctid=Request.querystring("acctid")
ypid=Request.querystring("ypid")
by=Request.querystring("by")
bp=Request.querystring("bp")
utility = Request.querystring("utility")

Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_genergy1")



sqlstr= "select * from utilitybill_steam where acctid='" &acctid& "' and ypid='"&ypid&"'"
'response.write sqlstr
'response.end	
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
dim id, fuel, mlb, st, AvgCost, totalamt, CDD, HDD, Avg_DD
if not rst1.EOF then
    id = rst1("id")
	fuel = rst1("fueladj")
	mlb = rst1("mlbusage")
	st = rst1("salestax")
	AvgCost = rst1("AvgCost")
	totalamt = rst1("totalbillamt")
	CDD = rst1("CDD")
	HDD = rst1("HDD")
	Avg_DD = rst1("Avg_DD")
else
    id = request("id")
	fuel = request("fuel")
	mlb = request("mlb")
	st = request("st")
	AvgCost = request("AvgCost")
	totalamt = request("totalamt")
	CDD = request("CDD")
	HDD = request("HDD")
	Avg_DD = request("Avg_DD")
end if

%>

<body bgcolor="#FFFFFF">
<table width="100%" border="0">
<tr><td bgcolor="#3399CC" height="30"> 
	<font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Utility Bill for Period</i></font><b><font size="4" color="#FFFFFF"> <%=Request.querystring("bp")%></font></b><font size="4" color="#FFFFFF"><i>Year</i></font><b><font size="4" color="#FFFFFF"> <%=Request.querystring("by")%></font></b></font>
    </td>
</tr>
</table>


<form name="detail" method="post" action=" updsutilitybill.asp">
<table border="0">
<tr><td> 
		<input type="hidden" name="id"value="<%=id%>">
		<input type="hidden" name="by"value="<%=by%>">
		<input type="hidden" name="bp"value="<%=bp%>">
		<input type="hidden" name="ypid"value="<%=ypid%>">
		<input type="hidden" name="acctid"value="<%=acctid%>">
		<u>Cost</u></td><td></td>
	<td></td>
	<td><u>Usage</u></td>
	<td></td>
	<td>&nbsp; </td>
</tr>
<tr><td>Total Bill Amount</td><td>$</td>
	<td><input type="text" name="totalamt" value="<%=totalamt%>" onKeyUp="calcAvgCost()"></td>
	<td>MLbs</td><td></td>
	<td><input type="text" name="mlb" value="<%=mlb%>" onKeyUp="calcAvgCost();calcAvg_DD()"></td>
	<td>&nbsp;</td>
</tr>
<tr><td>Sales Tax</td><td>$</td>
	<td><input type="text" name="st" value="<%=st%>"></td>
	<td>HDD</td>
	<td></td>
	<td><input type="text" name="HDD" value="<%=HDD%>" onKeyUp="calcAvg_DD()"></td>
</tr>
<tr><td>Avg Cost/M#</td><td>$</td>
	<td><input type="text" name="AvgCost" value="<%=AvgCost%>"></td>
	<td>CDD</td>
	<td>&nbsp;</td>
	<td><input type="text" name="CDD" value="<%=CDD%>" onKeyUp="calcAvg_DD()"></td>
</tr>
<tr> 
	<td>Fuel Adjustment</td><td>$</td>
	<td><input type="text" name="fuel" value="<%=fuel%>"></td>
	<td>Avg M# per HDD/CDD</td>
	<td></td>
	<td><input type="text" name="Avg_DD" value="<%=Avg_DD%>" onKeyUp="calcAvg_DD()"></td>
	<td></td>
</tr>
<tr> 
	<td></td><td></td><td></td><td></td><td></td>
	  <td>&nbsp; </td>
</tr>
</table>
  <%if not rst1.EOF then%>
  <font size="2" face="Arial, Helvetica, sans-serif">GL CODE 
  <input name="glcode" type="text" size="10" maxlength="20">
  <input type="submit" name="action" value="UPDATE">
  <input type="button" name="download" value="DOWNLOAD DATA">
  <%else%>
  GL CODE 
  <input name="glcode2" type="text" size="10" maxlength="20">
  <input type="submit" name="action" value="SAVE">
  <input type="button" name="download" value="DOWNLOAD DATA">
  </font> 
  <%end if%>
</form>
</body>
</html>
<%
rst1.close
set cnn1=nothing
%>