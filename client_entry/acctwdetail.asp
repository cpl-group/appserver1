<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<title>Account Detail</title>
<%
dim acctid, ypid, by, bp
acctid=Request.querystring("acctid")
ypid=Request.querystring("ypid")
by = Request.querystring("by")
bp = Request.querystring("bp")

Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_genergy1")

sqlstr="select DateStart, DateEnd, datediff(day, DateStart, DateEnd) as daysinbillperiod from billyrperiod where billperiod='" &bp & "'"

rst1.Open sqlstr, cnn1, 0, 1, 1

dim daysinbillperiod
daysinbillperiod = rst1("daysinbillperiod")

%>
<script>
function roundNumber(num,precision)
{	var i,sZeros = '';
	for(i = 0;i < precision;i++)
		sZeros += '0';
	i = Number(1 + sZeros);
	return Math.round(num * i) / i;
}

function calcAvgCost()
{	frm = document.forms['detail'];
	frm.avgcost.value = roundNumber(parseFloat(frm.totalamt.value)/parseFloat(frm.totalhcf.value),4);
}

function calcTotalAmt()
{	frm = document.forms['detail'];
	frm.totalamt.value = roundNumber(parseFloat(frm.sewer.value)+parseFloat(frm.watercharge.value),2);
}

function zeroNaNs(frm)
{	for(i=0;i<frm.elements.length;i++)
	{	//alert(frm.elements[i].name+':'+(frm.elements[i].type=='text')+((isNaN(frm.elements[i].value))||(!(isFinite(frm.elements[i].value)))) )
		if((frm.elements[i].type=='text')&&(isNaN(frm.elements[i].value)||!(isFinite(frm.elements[i].value))))
		{	frm.elements[i].value=''
		}
	}
}

function calcAvgDailyUsage(){
  frm = document.forms['detail'];
	frm.avgdailyusage.value = roundNumber(parseFloat(<%=daysinbillperiod%>)/parseFloat(frm.totalhcf.value),4);
}
</script>

</head>
<%
rst1.close

sqlstr= "select * from utilitybill_coldwater where acctid='" &acctid& "' and ypid='"&ypid&"'"
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1

dim id, watercharge, totalhcf, sewer, avgcost, totalamt, avgdailyusage
if not rst1.EOF then
	id = rst1("id")
	watercharge = rst1("watercharge")
	totalhcf = rst1("totalhcf")
	sewer = rst1("sewercharge")
	avgcost = rst1("avgcost")
	totalamt = rst1("totalamt")
	avgdailyusage = rst1("avgdailyusage")
else
	id = request("id")
	watercharge = request("watercharge")
	totalhcf = request("totalhcf")
	sewer = request("sewer")
	avgcost = request("avgcost")
	totalamt = request("totalamt")
	avgdailyusage = request("avgdailyusage")
end if
%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0">
	<tr><td bgcolor="#3399CC" height="30"> 
		<font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Utility Bill for Period</i></font><b><font size="4" color="#FFFFFF"> <%=Request.querystring("bp")%></font></b><font size="4" color="#FFFFFF"><i>Year</i></font><b><font size="4" color="#FFFFFF"> <%=Request.querystring("by")%></font></b></font>
	</td></tr>
</table>
 <form name="detail" method="post" action=" updwutilitybill.asp" onsubmit="zeroNaNs(this)">
  <table width="100%" border="0" style="font-family: Arial, Helvetica, sans-serif">
    <tr> </tr>
    <tr> 
      <td> 
        <input type="hidden" name="id" value="<%=id%>">
        <input type="hidden" name="ypid"value="<%=ypid%>">
        <input type="hidden" name="acctid"value="<%=acctid%>">
		<input type="hidden" name="by"value="<%=by%>">
        <input type="hidden" name="bp"value="<%=bp%>">
        <u>Cost</u></td>
      <td></td><td></td>
      <td><u>Usage</u></td>
      <td></td><td></td>
      <td>&nbsp; </td>
      <td>&nbsp; </td>
    </tr>
    <tr> 
      <td>Water&nbsp;Charge</td><td>$</td>
      <td><input type="text" name="watercharge" value="<%=watercharge%>" onKeyUp="calcTotalAmt()"></td>
      <td>Total&nbsp;Usage&nbsp;(HCF)</td><td></td>
      <td><input type="text" name="totalhcf" value="<%=totalhcf%>" onKeyUp="calcAvgCost();calcAvgDailyUsage()"></td>
      <td>&nbsp; </td>
      <td></td>
    </tr>
    <tr> 
      <td>Sewer&nbsp;Charge</td><td>$</td>
      <td><input type="text" name="sewer" value="<%=trim(sewer)%>" onKeyUp="calcTotalAmt()"></td>
      <td>Avg&nbsp;Cost</td><td>$</td>
      <td><input type="text" name="avgcost" value="<%=trim(avgcost)%>"></td>
      <td></td>
      <td></td>
    </tr>
    <tr> 
      <td>Total&nbsp;Bill&nbsp;Amount</td><td>$</td>
      <td><input type="text" name="totalamt" value="<%=totalamt%>" onKeyUp="calcAvgCost()"></td>
      <td>Average&nbsp;Daily&nbsp;Usage</td><td></td>
      <td><input type="text" name="avgdailyusage" value="<%=avgdailyusage%>"></td>
      <td>
		<%if not rst1.EOF then%>
		      <input type="submit" name="action" value="UPDATE">
		<%else%>
		      <input type="submit" name="action" value="SAVE">
		<%end if%>
      </td>
      <td width="24%"></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
rst1.close

'response.write sqlstr
'response.write rst1("daysinbillperiod") & vbCrLf
'response.write rst1("DateStart") & vbCrLf
'response.write rst1("DateEnd") & vbCrLf
'response.end

set cnn1=nothing
%>