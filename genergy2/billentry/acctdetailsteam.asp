<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
	if ((CDD==0) && (HDD ==0)){
		frm.Avg_DD.value = 0;
	}
	else {
	frm.Avg_DD.value = roundNumber(parseFloat(frm.mlb.value)/(CDD + HDD),2);
	}
}

function calcAvgCost()
{	frm = document.forms['detail'];
	var taxtemp = 0;
	if(frm.taxIncluded.checked) taxtemp = parseFloat(frm.st.value);
	
	frm.AvgCost.value = roundNumber((parseFloat(frm.totalamt.value)-taxtemp)/parseFloat(frm.mlb.value),4)
	//alert(taxtemp);
}

function calcTaxDollars(){
//	frm = document.forms['detail'];
//	if(frm.taxIncluded.checked)
//		frm.st.value = roundNumber( (parseFloat(frm.stp.value) * parseFloat(frm.totalamt.value)) / ( 1 + parseFloat(frm.stp.value)) , 2)
}

function calcTaxPercent(){
//frm = document.forms['detail'];
	//if(frm.taxIncluded.checked)
		//frm.stp.value = roundNumber(   parseFloat(frm.st.value) / (parseFloat(frm.totalamt.value) - parseFloat(frm.st.value))    , 5)
}

</script>
<link rel="Stylesheet" href="/GENERGY2/styles.css" type="text/css">	
</head>
<%
dim acctid, ypid, by, bp, utility, bldg
bldg=Request.querystring("bldg")
acctid=Request.querystring("acctid")
ypid=Request.querystring("ypid")
by=Request.querystring("by")
bp=Request.querystring("bp")
utility = Request.querystring("utility")

Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(bldg)



sqlstr= "select *, isnull(taxincluded,0) as taxin from utilitybill_steam where acctid='" &acctid& "' and ypid='"&ypid&"'"
'response.write sqlstr
'response.end	
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
dim id, fuel, mlb, st, AvgCost, totalamt, CDD, HDD, Avg_DD, MAC, GRT, stp, mlbs_hr, note, taxincluded
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
	GRT = rst1("grtpercent")
	stp = rst1("salestaxpercent")
	MAC = rst1("macdollar")
	mlbs_hr = rst1("mlbs_hr")
	note = rst1("note")
	taxincluded = rst1("taxin")
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
	GRT = request("GRT")
	stp = request("stp")
	MAC = request("MAC")
	mlbs_hr = request("mlbs_hr")
	note = request("note")
	taxincluded = false
end if

%>

<body bgcolor="#eeeeee">
<table width="100%" border="0">
<td bgcolor="#6699cc"><span class="standardheader">Utility Bill for Period <%=request("bp")%>, Year <%=request("by")%></span></td>
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
	<td><input type="text" name="totalamt" value="<%=totalamt%>" onKeyUp="calcAvgCost();calcTaxPercent()"><input type="Checkbox" name="taxIncluded" value="1" onclick="calcAvgCost()" <%if taxincluded then%>CHECKED<%end if%>>&nbsp;Tax&nbsp;Included</td>
	<td>MLbs</td><td></td>
	<td><input type="text" name="mlb" value="<%=mlb%>" onKeyUp="calcAvgCost();calcAvg_DD()"></td>
</tr>
<tr><td>Sales Tax</td><td>$</td>
	<td><input type="text" name="st" value="<%=st%>" onKeyUp="calcTaxPercent();calcAvgCost()"></td>
	<td>MLbs/H</td><td></td>
	<td><input type="text" name="mlbs_hr" value="<%=mlbs_hr%>" onKeyUp=""></td>
</tr>
<tr><td>Sales Tax</td><td>%</td>
	<td><input type="text" name="stp" value="<%=stp%>" onKeyUp="calcTaxDollars()"></td>
	<td>HDD</td>
	<td></td>
	<td><input type="text" name="HDD" value="<%=HDD%>" onKeyUp="calcAvg_DD()"></td>
</tr>
<tr>
	<td>Avg Cost/M#</td><td>$</td>
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
</tr>
<tr> 
	<td>Gross Receipt Tax</td>
	<td>%</td>
	<td><input type="text" name="GRT" value="<%=GRT%>"></td>
	<td>Fuel Unit Amount</td>
	<td></td>	
	<td><input type="text" name="MAC" value="<%=MAC%>"></td>
</tr>
</table>
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr><td rowspan="2" valign="top">
<%if not(isBuildingOff(bldg)) then%>
	  <%if not rst1.EOF then%>
	  <input type="submit" name="action" value="UPDATE">
	  <%else%>
	  <input type="submit" name="action" value="SAVE">
	  <%end if%>
<%end if%>
	  <!--<input type="button" name="download" value="DOWNLOAD DATA">-->
	  <input type="hidden" name="bldg" value="<%=bldg%>">
		</td>
		<td>Additional Notes <font size="-2">(500 character limit)</font></td></tr>
<tr><!-- <td></td> -->
		<td><textarea cols="120" rows="4" name="note" onkeyup="if(500<this.value.length) this.value=this.value.substr(0,500);"><%=note%></textarea></td>
</tr>
<tr><td><!-- <font size="2" face="Arial, Helvetica, sans-serif"> GL CODE  --><input name="glcode" type="hidden" size="10" maxlength="20"><!-- </font> --></td>
</table>
</form>
</body>
</html>
<%
rst1.close
set rst1 = nothing  'TK: 04/28/2006
set cnn1=nothing
%>