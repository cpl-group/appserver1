<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Account Detail</title>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<%
dim acctid, ypid, by, bp, bldg, utility
bldg=Request.querystring("bldg")
acctid=Request.querystring("acctid")
ypid=Request.querystring("ypid")
by = Request.querystring("by")
bp = Request.querystring("bp")
utility = trim(request("utility"))
Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(bldg)

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
	frm.avgcost.value = roundNumber(parseFloat(frm.totalamt.value)/parseFloat(frm.totalccf.value),4);
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
	frm.avgdailyusage.value = roundNumber(parseFloat(<%=daysinbillperiod%>)/parseFloat(frm.totalccf.value),4);
}

function zeroNaNsChill(frm)
{	for(i=0;i<frm.elements.length;i++)
	{	//alert(frm.elements[i].name+':'+(frm.elements[i].type=='text')+((isNaN(frm.elements[i].value))||(!(isFinite(frm.elements[i].value)))) )
		if((frm.elements[i].type=='text') && (frm.elements[i].value != "") && isNaN(frm.elements[i].value) && frm.elements[i] != frm.MiscChargeDesc)
		{	frm.elements[i].value="";
		}
	}
}


function updatecwTotal(){
	frm = document.forms['cw'];
	frm.subtotal.value = roundDecimal ( makeItANumber(frm.usageCharge.value) + makeItANumber(frm.eeCharge.value) + makeItANumber(frm.opCharge.value) + makeItANumber(frm.capCharge.value) + makeItANumber(frm.tempCharge.value) + makeItANumber(frm.MiscCharge.value), 2);//makeItANumber(frm.auCharge.value)
	frm.subtotalAdj.value = roundDecimal ( makeItANumber(frm.adjustments.value) + makeItANumber(frm.subtotal.value) , 2);
	frm.totalBillAmt.value = roundDecimal (makeItANumber(frm.subtotalAdj.value) + makeItANumber(frm.taxDollar.value) + (getTaxPercent() * makeItANumber(frm.subtotalAdj.value)) , 2);
	frm.totalBillAmt.style.backgroundColor=getPinkEye(frm.totalBillAmt.value);
}

function getTaxPercent(){
	frm = document.forms['cw'];
	var percentage = (makeItANumber(frm.taxPercent.value) / 100.0);
	return percentage;
}

function calccwTotalTH(){
	frm = document.forms['cw'];
	frm.totalTonHrs.value = makeItANumber(frm.onPeakTonHrs.value) + makeItANumber(frm.offPeakTonHrs.value);
	frm.totalTonHrs.style.backgroundColor=getPinkEye(frm.totalTonHrs.value);
	CalcUnitCostTonh();
}

function cwSubtotalChanged(){
	frm = document.forms['cw'];
	frm.usageCharge.value = "";
	frm.eeCharge.value = "";
	//frm.auCharge.value = "";
	frm.opCharge.value = "";
	frm.capCharge.value = "";
	frm.tempCharge.value = "";
	frm.MiscCharge.value = "";
	frm.subtotalAdj.value = roundDecimal(makeItANumber(frm.adjustments.value) + makeItANumber(frm.subtotal.value),2);
	frm.totalBillAmt.value = roundDecimal(makeItANumber(frm.subtotalAdj.value) + makeItANumber(frm.taxDollar.value), 2);
	frm.totalBillAmt.style.backgroundColor=getPinkEye(frm.totalBillAmt.value);
}
function cwSubtotalAdjChanged(){
	frm = document.forms['cw'];
	frm.usageCharge.value = "";
	frm.eeCharge.value = "";
	//frm.auCharge.value = "";
	frm.opCharge.value = "";
	frm.capCharge.value = "";
	frm.tempCharge.value = "";
	frm.subtotal.value = "";
	frm.adjustments.value = "";
	frm.MiscCharge.value = "";
	frm.totalBillAmt.value = roundDecimal (makeItANumber(frm.subtotalAdj.value) + makeItANumber(frm.taxDollar.value) , 2);
	frm.totalBillAmt.style.backgroundColor=getPinkEye(frm.totalBillAmt.value);
}
function cwTotalChanged(){
	frm = document.forms['cw'];
	frm.usageCharge.value = "";
	frm.eeCharge.value = "";
	//frm.auCharge.value = "";
	frm.opCharge.value = "";
	frm.capCharge.value = "";
	frm.tempCharge.value = "";
	frm.subtotal.value = "";
	frm.adjustments.value = "";
	frm.subtotalAdj.value = "";
	frm.taxPercent.value = "";
	frm.taxDollar.value = "";
	frm.MiscCharge.value = "";
}
function cwAdjChanged(){
	frm = document.forms['cw'];
	frm.subtotalAdj.value = roundDecimal( makeItANumber(frm.adjustments.value) + makeItANumber(frm.subtotal.value) , 2);
	frm.totalBillAmt.value = roundDecimal( makeItANumber(frm.subtotalAdj.value) + makeItANumber(frm.taxDollar.value) , 2);
	frm.totalBillAmt.style.backgroundColor=getPinkEye(frm.totalBillAmt.value);
}

function cwTaxPercentChanged(){
	frm = document.forms['cw'];
	var taxDollarTemp = getTaxPercent() * makeItANumber(frm.subtotalAdj.value)
	frm.taxDollar.value = roundDecimal( taxDollarTemp , 3);
	frm.totalBillAmt.value = roundDecimal(  makeItANumber(frm.subtotalAdj.value) + makeItANumber(frm.taxDollar.value) , 2);
}

function cwTaxDollarChanged(){
	frm = document.forms['cw'];
	if (makeItANumber(frm.subtotalAdj.value) * 100 != 0) {
		var taxPercentTemp = makeItANumber(frm.taxDollar.value) / makeItANumber(frm.subtotalAdj.value) * 100;
		frm.taxPercent.value = roundDecimal( taxPercentTemp , 3);
	}	else {
		frm.taxPercent.value = ""
		frm.taxPercent.style.backgroundColor=getPinkEye(frm.taxPercent.value);
	}
	frm.totalBillAmt.value = roundDecimal (makeItANumber(frm.subtotalAdj.value) + makeItANumber(frm.taxDollar.value) , 2);
}	

function makeItANumber(value){
	var toBeReturned
	if (isNaN(value) || value==null || value==""){
		toBeReturned = 0;
	} else {
		toBeReturned = parseFloat(value);
	}
	return toBeReturned;
}

function onChangeMakeNumber(value){
	if (value == "") {
		return ""
	} else{
		return new Number(value);
	}
}
		

function roundDecimal(number, decPlaces){
	var temp = number * Math.pow(10, decPlaces);
	temp = Math.round(temp);
	temp = temp / Math.pow(10, decPlaces);
	return temp;
}

function getPinkEye(someVal){
	if ((someVal == null) || (someVal == "") || isNaN(someVal)) {
		return "#FFCCCC";
	}else {
		return "#FFFFFF";
	}
}

function miscChargeDescFocus(field){
	frm = document.forms['cw'];
	if ((frm.MiscCharge.value == "") || (frm.MiscCharge.value == null)){
		field.blur();
	}
}

function CalcUnitCostTonh(){
	frm = document.forms[0];
	var usageCharge = frm.usageCharge.value;
	var totalTonHrs = frm.totalTonHrs.value;
	if(isNaN(usageCharge)) usageCharge = 0;
	if(isNaN(totalTonHrs)||totalTonHrs==0){frm.unitCostTonh.value = 0;}
	else{frm.unitCostTonh.value = usageCharge/totalTonHrs;}
}
</script>

</head>
<body bgcolor="#eeeeee">
<%
rst1.close

if utility = "3" then			'cold water
	sqlstr= "select * from utilitybill_coldwater where acctid='" &acctid& "' and ypid='"&ypid&"'"
elseif utility = "6" then		'chilled water
	sqlstr= "select * from utilitybill_chilledwater where acctid='" &acctid& "' and ypid='"&ypid&"'"
end if

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1

dim id, watercharge, totalccf, sewer, avgcost, totalamt, avgdailyusage		'these are for cold water
dim usageCharge, eeCharge, opCharge, capCharge, tempCharge,  adjustments, totalTons, totalTonHrs, onPeakTonHrs, offPeakTonHrs, penalty, totalBillAmt,auCharge, taxPercent, MiscCharge, MiscChargeDesc, unitCostTonh, note
if utility = "3" then			'cold water
	if not rst1.EOF then
		id = rst1("id")
		watercharge = rst1("watercharge")
		totalccf = rst1("totalccf")
		sewer = rst1("sewercharge")
		avgcost = rst1("avgcost")
		totalamt = rst1("totalbillamt")
		avgdailyusage = rst1("avgdailyusage")
		note = rst1("note")
	else
		id = request("id")
		watercharge = request("watercharge")
		totalccf = request("totalccf")
		sewer = request("sewer")
		avgcost = request("avgcost")
		totalamt = request("totalamt")
		avgdailyusage = request("avgdailyusage")
		note = request("note")
	end if
elseif utility = "6" then		'chilled water
	if not rst1.EOF then
		id = rst1("id")
		usagecharge = rst1("usagecharge")
		totalBillAmt = rst1("totalbillamt")
		eeCharge = rst1("eeCharge")
		opCharge = rst1("opCharge")
		capCharge = rst1("capCharge")
		tempCharge = rst1("tempCharge")
		adjustments = rst1("adjustments")
		totalTons = rst1("totalTons")
		totaltonHrs = rst1("totalTonHrs")
		onPeakTonHrs = rst1("onPeakTonh")
		offPeakTonHrs = rst1("offPeakTonh")
		penalty= rst1("penalty_kwh")
		'auCharge = rst1("auCharge")
		taxPercent = rst1("salesTax")
		MiscCharge = rst1("MiscCharge")
		MiscChargeDesc = rst1("MiscChargeDesc")
		unitCostTonh = rst1("unitCostTonh")
		note = rst1("note")
	end if
end if

if utility = "3" then 		'cold water %>
	<form name="detail" method="post" action="updwutilitybill.asp" onsubmit="zeroNaNs(this)">
	<input type="hidden" name="id" value="<%=id%>">
	<input type="hidden" name="ypid" value="<%=ypid%>">
	<input type="hidden" name="acctid" value="<%=acctid%>">
	<input type="hidden" name="by" value="<%=by%>">
	<input type="hidden" name="bp" value="<%=bp%>">
	<input type="hidden" name="utility" value="<%=utility%>">
	<input type="hidden" name="bldg" value="<%=bldg%>">
	<table width="100%" border="0" style="font-family: Arial, Helvetica, sans-serif">
		<tr>
			<td bgcolor="#6699cc" colspan="8">
				<span class="standardheader">Utility Bill for Period <%=request("bp")%>, Year <%=request("by")%></span>
			</td>
		</tr>		
		<tr> 
			<td><u>Cost</u></td>
      		<td colspan="2"></td>
      		<td><u>Usage</u></td>
      		<td colspan="4"></td>
    	</tr>
    	<tr> 
    		<td>Water&nbsp;Charge</td><td>$</td>
      		<td><input type="text" name="watercharge" value="<%=watercharge%>" onKeyUp="calcTotalAmt()"></td>
      		<td>Total&nbsp;Usage&nbsp;(CCF)</td><td></td>
      		<td><input type="text" name="totalccf" value="<%=totalccf%>" onKeyUp="calcAvgCost();calcAvgDailyUsage()"></td>
      		<td colspan="2"></td>
		</tr>
		<tr> 
			<td>Sewer&nbsp;Charge</td><td>$</td>
			<td><input type="text" name="sewer" value="<%=trim(sewer)%>" onKeyUp="calcTotalAmt()"></td>
			<td>Avg&nbsp;Cost</td><td>$</td>
			<td><input type="text" name="avgcost" value="<%=trim(avgcost)%>"></td>
			<td colspan=2></td>
		</tr>
		<tr> 
			<td>Total&nbsp;Bill&nbsp;Amount</td><td>$</td>
			<td><input type="text" name="totalamt" value="<%=totalamt%>" onKeyUp="calcAvgCost()"></td>
			<td>Average&nbsp;Daily&nbsp;Usage</td><td></td>
			<td><input type="text" name="avgdailyusage" value="<%=avgdailyusage%>"></td>
			<td>
      		</td>
     		<td width="24%"></td>
    	</tr>
  	</table>
<%
elseif utility = "6" then%>
	<table width="100%" border="0" cellspacing="0" cellpadding="1">
		<tr>
			<td bgcolor="#6699cc" colspan="10">
				<span class="standardheader">Utility Bill for Period <%=request("bp")%>, Year <%=request("by")%></span>
			</td>
		</tr>
 		<form name="cw" method="post" action=" updwutilitybill.asp" onSubmit="zeroNaNsChill(this)">
		<input type="hidden" name="id" value="<%=id%>">
		<input type="hidden" name="ypid"value="<%=ypid%>">
		<input type="hidden" name="acctid"value="<%=acctid%>">
		<input type="hidden" name="by"value="<%=by%>">
		<input type="hidden" name="bp"value="<%=bp%>">
		<input type="hidden" name="utility" value="<%=utility%>">
		<input type="hidden" name="bldg" value="<%=bldg%>">
		<tr> 
			<td colspan=5 align="center"><u>Cost</u></td>
      		<td colspan="2" align="center"><u>Usage</u></td>
			<td colspan="2" align="center"></td>
		</tr>
		<tr>  
			<td>Capacity Charge</td>
			<td>$</td>
			<td colspan=3><input type="text" name="capCharge" value="<%=capCharge%>" onKeyUp="updatecwTotal()" tabindex=1 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td>Demand (Tons)</td>
			<td><input type="text" name="totalTons" value="<%=totalTons%>" tabindex=10 onChange="this.value = onChangeMakeNumber(this.value);this.style.backgroundColor=getPinkEye(this.value);"><b><font color="#FF0000">*</font></b></td>
			<td colspan="3"></td>
		</tr>	
		<tr>
			<td>Usage Charge</td>
			<td>$</td>
			<td colspan=3><input type="text" name="usageCharge" value="<%=usagecharge%>" onKeyUp="CalcUnitCostTonh();updatecwTotal();"  tabindex=2 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td>Peak Use (Ton-Hrs)</td>
			<td><input type="text" name="onPeakTonHrs" value="<%=onPeakTonHrs%>" onKeyUp="calccwTotalTH()" tabindex=11 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td colspan=3></td>
		</tr>
<!-- 		<tr> 
 			<td>Adjusted Usage Charge</td>
			<td>$</td>
			<td colspan=3><input type="text" name="auCharge" value="<%=auCharge%>" onKeyUp="updatecwTotal()" tabindex=3 onChange="this.value = onChangeMakeNumber(this.value);"></td> 
			<td>Off-Peak Use (Ton-Hrs)</td>
			<td>$</td>
			<td><input type="text" name="offPeakTonHrs" value="<%=offPeakTonHrs%>" onKeyUp="calccwTotalTH()" tabindex=12 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td colspan=3></td>
    	</tr>-->
		<tr>		
			<td>Electric Energy Charge</td>
			<td>$</td>
			<td colspan=3><input type="text" name="eeCharge" value="<%=eeCharge%>" onKeyUp="updatecwTotal()" tabindex=4 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td nowrap>Off-Peak Use (Ton-Hrs)</td>
			<td><input type="text" name="offPeakTonHrs" value="<%=offPeakTonHrs%>" onKeyUp="calccwTotalTH()" tabindex=12 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td colspan=3></td>
    	</tr>
		<tr> 
			<td nowrap>Return Temperature Charge</td>
			<td>$</td>
			<td colspan=3><input type="text" name="tempCharge" value="<%=tempCharge%>" onKeyUp="updatecwTotal()" tabindex=5 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td>Total Usage (Ton-Hrs)</td>
			<td><input type="text" name="totalTonHrs" value="<%=totalTonHrs%>" onKeyp="this.form.elements['offPeakTonHrs'].value='';this.form.elements['onPeakTonHrs'].value='';CalcUnitCostTonh()" tabindex=13 onChange="this.value = onChangeMakeNumber(this.value);this.style.backgroundColor=getPinkEye(this.value);"><b><font color="#FF0000">*</font></b></td>
			<td width="1%" align="right">Unit Cost</td>
			<td><input type="text" name="unitCostTonh" value="<%=unitCostTonh%>" size="7" readonly></td>
			<td width="30%">&nbsp;</td>
    	</tr>
		<tr>
			<td>Operating Charge</td>
			<td>$</td>
			<td colspan=3><input type="text" name="opCharge" value="<%=opCharge%>" onKeyUp="updatecwTotal()" tabindex=6 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td>Penalty (KWH)</td>
			<td><input type="text" name="penalty" value="<%=penalty%>" tabindex=14 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td colspan=3></td>
    	</tr>
		<tr>
			<td>Misc. Charge</td>
			<td>$</td>
			<td colspan=3><input type="text" name="MiscCharge" value="<%=MiscCharge%>" onKeyUp="updatecwTotal()" tabindex=6 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td>Misc.&nbsp;Charge&nbsp;Description</td>
			<td><input type="text" name="MiscChargeDesc" value="<%=MiscChargeDesc%>" tabindex=15 onFocus="miscChargeDescFocus(this);"></td>
			<td colspan=3></td>
    	</tr>
		<tr>
			<td>Subtotal</td>
			<td>$</td>
			<td colspan=3><input type="text" name="subtotal" value="" onKeyUp="cwSubtotalChanged()" onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td colspan=2 bgcolor="#dddddd" align="center" style="border-left:1px solid #ffffff;border-top:1px solid #ffffff;border-right:1px solid #bbbbbb;">
				<span class="standardheader"><font color="black">Key</font></span>
			</td>
    	</tr>
		<tr> 
			<td>Adjustments</td>
			<td>$</td>
			<td colspan=3><input type="text" name="adjustments" value="<%=adjustments%>" onKeyUp="cwAdjChanged()" tabindex=7 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td align="right" bgcolor="#dddddd" style="border-left:1px solid #ffffff;"><b><font color="#FF0000">*</font></b></td>
			<td colspan="1" bgcolor="#dddddd" style="border-right:1px solid #bbbbbb;">: Required Field</td>
		</tr>
		<tr> 
			<td>Subtotal Adjustments</td>
			<td>$</td>
			<td colspan=3><input type="text" name="subtotalAdj" onKeyUp="cwSubtotalAdjChanged()" onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td align="right" bgcolor="#dddddd" style="border-left:1px solid #ffffff;border-bottom:1px solid #bbbbbb;">NaN</td>
			<td colspan="1" bgcolor="#dddddd" style="border-bottom:1px solid #bbbbbb;border-right:1px solid #bbbbbb;">: Not a Number - please reenter input.</td>
		</tr>
		<tr>
			<td>Sales Tax</td>
			<td>$</td>
			<td><input type="text" name="taxDollar" onKeyUp="cwTaxDollarChanged()" size=3 tabindex=8 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td>%<input type="text" name="taxPercent" value="<%=taxPercent%>" onKeyUp="cwTaxPercentChanged()" size=2 tabindex=9 onChange="this.value = onChangeMakeNumber(this.value);"></td>
			<td colspan=5></td>
		</tr>
		<tr> 	
			<td><b>Total Amount</b></td>
			<td>$</td>
			<td colspan=3><input type="text" name="totalBillAmt" value="<%=totalBillAmt%>" onKeyUp="cwTotalChanged()" onChange="this.value = onChangeMakeNumber(this.value);this.style.backgroundColor=getPinkEye(this.value);"><b><font color="#FF0000">*</font></b></td>
			<td colspan=5 align="left"></td>
		</tr>
		<tr height="20">
		</tr>
		<tr>
			<td colspan=6>
			<td>
			</td>
			<td width="24%"></td>
    	</tr>

  	</table>
<%if not rst1.EOF then%><script>updatecwTotal();cwTaxPercentChanged()</script><%end if%>
<%end if%>

<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr><td rowspan="2" valign="top">
		<%if not(isBuildingOff(bldg)) then%>
			<%if not rst1.EOF then%>
				<input type="submit" name="action" value="UPDATE">
			<%else%>
				<input type="submit" name="action" value="SAVE">
			<%end if%>
		<%end if%>
		</td>
		<td>Additional Notes <font size="-2">(500 character limit)</font></td></tr>
<tr><!-- <td></td> -->
		<td><textarea cols="120" rows="4" name="note" onkeyup="if(500<this.value.length) this.value=this.value.substr(0,500);"><%=note%></textarea></td>
</tr>
<tr><td><!-- <font size="2" face="Arial, Helvetica, sans-serif"> GL CODE  --><input name="glcode" type="hidden" size="10" maxlength="20"><!-- </font> --></td>
</table>
</form>


<%
rst1.close
set rst1 = nothing 'TK: 04/28/2006
set cnn1=nothing
%>

</body>
</html>