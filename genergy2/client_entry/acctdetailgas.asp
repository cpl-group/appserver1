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

function calcAvgCostTherm()
{	frm = document.forms['detail'];
	frm.AvgCostTherm.value = roundNumber(parseFloat(frm.tamt2.value)/parseFloat(frm.therms.value),4);
}

function calcSalesTax()
{	frm = document.forms['detail'];
	frm.st2.value = roundNumber(parseFloat(frm.tamt2.value)-(parseFloat(frm.tamt2.value)/1.0825),2)
}

function zeroNaNs(frm)
{	for(i=0;i<frm.elements.length;i++)
	{	if((frm.elements[i].type=='text')&&(isNaN(frm.elements[i].value)||!(isFinite(frm.elements[i].value))||(frm.elements[i].value=='')))
		{	frm.elements[i].value='0'
		}
	}
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
sqlstr= "select * from utilitybill_gas where acctid='" &acctid& "' and ypid='"&ypid&"'"
	
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
dim id, fuel, therms, st2, ccf, AvgCostTherm, tamt2
if not rst1.EOF then
	id = rst1("id")
	fuel = rst1("fueladj")
	therms = rst1("thermusage")
	st2 = rst1("salestax")
	ccf = rst1("ccfusage")
	tamt2 = rst1("totalbillamt")
	AvgCostTherm = rst1("AvgCostTherm")
else
	id = request("id")
	fuel = request("fuel")
	therms = request("therms")
	st2 = request("st2")
	ccf = request("ccf")
	tamt2 = request("tamt2")
	AvgCostTherm = request("AvgCostTherm")
end if
%>

<body bgcolor="#FFFFFF">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Utility Bill for Period</i></font><b><font size="4" color="#FFFFFF"> <%=Request.querystring("bp")%></font></b><font size="4" color="#FFFFFF"><i>Year</i></font><b><font size="4" color="#FFFFFF"> <%=Request.querystring("by")%></font></b></font>
    </td>
  </tr>
</table>
<form name="detail" method="post" action="updgutilitybill.asp" onSubmit="zeroNaNs(frm)">
<table border="0" style="font-family: Arial, Helvetica, sans-serif">
      <td> 
        <input type="hidden" name="id" value="<%=id%>">
		<input type="hidden" name="utility" value="<%=utility%>">
        <input type="hidden" name="by" value="<%=by%>">
        <input type="hidden" name="bp" value="<%=bp%>">
        <input type="hidden" name="ypid" value="<%=ypid%>">
        <input type="hidden" name="acctid" value="<%=acctid%>">
        <u>Cost</u></td>
      <td></td><td></td>
      <td><u>Usage</u></td>
      <td></td>
      <td>&nbsp; </td>
    </tr>
    <tr> 
      <td>Total Bill Amount</td><td>$</td>
      <td><input type="text" name="tamt2" value="<%=tamt2%>" onKeyUp="calcAvgCostTherm();calcSalesTax();"></td>
      <td>Therms</td><td></td>
      <td><input type="text" name="therms" value="<%=therms%>" onKeyUp="calcAvgCostTherm()"></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>Sales Tax</td><td>$</td>
      <td><input type="text" name="st2" value="<%=st2%>"></td>
      <td>Total (ccf)</td><td></td>
      <td><input type="text" name="ccf" value="<%=ccf%>"></td>
      <td></td>
    </tr>
    <tr> 
      <td>Avg Cost P/Therm.</td><td>$</td>
      <td><input type="text" name="AvgCostTherm" value="<%=AvgCostTherm%>"></td>
      <td></td>
      <td>&nbsp; </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>Fuel Adjustment</td><td>$</td>
      <td><input type="text" name="fuel" value="<%=fuel%>"></td>
      <td></td>
      <td></td>
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
  <input name="glcode" type="text" size="10" maxlength="20">
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
