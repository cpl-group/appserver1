<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Account Detail</title>
<link rel="Stylesheet" href="/GENERGY2/styles.css" type="text/css">
<script>
function roundNumber(num,precision)
{	var i,sZeros = '';
	for(i = 0;i < precision;i++)
		sZeros += '0';
	i = Number(1 + sZeros);
	return Math.round(num * i) / i;
}

function calcAvgCostTherm()
{	frm = document.forms.detail;
	frm.AvgCostTherm.value = roundNumber(parseFloat(frm.tamt2.value)/(parseFloat(frm.therms.value)),4);
}

function calcSalesTax()
{	frm = document.forms.detail;
	frm.st2.value = roundNumber(parseFloat(frm.tamt2.value)*parseFloat(frm.salestaxrate.value),2)
}

function calcTherms(){
	frm = document.forms.detail;
	frm.therms.value = roundNumber(parseFloat(frm.conversion.value)*parseFloat(frm.ccf.value),2);
	calcFuel();
	calcAvgCostTherm();
}

function calcFuel(){
	frm = document.forms.detail;
	frm.fuel.value = roundNumber(parseFloat(frm.fuelrate.value)*parseFloat(frm.therms.value),2);
}

function zeroNaNs(frm)
{	for(i=0;i<frm.elements.length;i++)
	{	if((frm.elements[i].type=='text')&&(isNaN(frm.elements[i].value)||!(isFinite(frm.elements[i].value))||(frm.elements[i].value=='')))
		{	frm.elements[i].value='0'
		}
	}
}

function updatePriTherms(){
	frm = document.forms.detail;
	frm.therms.value = frm.sec_therm.value;
	calcAvgCostTherm();
}

function updateSecTherms(){
	frm = document.forms.detail;
	frm.sec_therm.value = frm.therms.value;
	recalcSecUll();
	recalcTax();
	updateSecTotal();
}
</script>
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">
</head>
<%
dim acctid, ypid, by, bp, utility, bldg
bldg=Request("bldg")
acctid=Request("acctid")
ypid=Request("ypid")
by=Request("by")
bp=Request("bp")
utility = Request("utility")

Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(bldg)
sqlstr= "select * from utilitybill_gas where acctid='" &acctid& "' and ypid='"&ypid&"'"
	
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
dim id, fuel, therms, st2, ccf, AvgCostTherm, tamt2

dim sec_CostPerTherm,sec_Therm,sec_TransCharge,sec_salestax,sec_ullp,sec_ull,sec_total, sec_salestaxp,note, GRTAmt, GRTrate, salestaxrate, fuelrate, conversion, taxincluded

if not rst1.EOF then
	id = rst1("id")
	fuel = rst1("fueladj")
	therms = rst1("thermusage")
	st2 = rst1("salestax")
	ccf = rst1("ccfusage")
	tamt2 = rst1("totalbillamt")
	AvgCostTherm = rst1("AvgCostTherm")
	GRTAmt = rst1("GRTAmt")
	GRTrate = rst1("GRTrate")
	salestaxrate = rst1("salestaxrate")
	fuelrate = rst1("fuelrate")
	conversion = rst1("conversion")
	
	sec_CostPerTherm = rst1("sec_CostPerTherm")
	sec_total = rst1("sec_total")
	sec_TransCharge = rst1("sec_TransCharge")
	sec_salestax = rst1("sec_salestax")
	sec_ullp = rst1("sec_ullp")
	sec_ull = rst1("sec_ull")
	sec_salestaxp = rst1("sec_salestaxp")
	note = rst1("note")
	taxincluded = rst1("taxincluded")
else
	id = request("id")
	fuel = request("fuel")
	therms = request("therms")
	st2 = request("st2")
	ccf = request("ccf")
	tamt2 = request("tamt2")
	AvgCostTherm = request("AvgCostTherm")
	GRTAmt = request("GRTAmt")
	GRTrate = request("GRTrate")
	salestaxrate = request("salestaxrate")
	fuelrate = request("fuelrate")
	conversion = request("conversion")
	
	sec_CostPerTherm = request("sec_CostPerTherm")
	sec_Therm = request("sec_Therm")
	sec_TransCharge = request("sec_TransCharge")
	sec_salestax = request("sec_salestax")
	sec_ullp = request("sec_ullp")
	sec_ull = request("sec_ull")
	sec_salestaxp = request("sec_salestaxp")
	note = request("note")
	taxincluded = request("taxincluded")
end if
%>

<body bgcolor="#eeeeee">
<table width="100%" border="0">
<td bgcolor="#6699cc"><span class="standardheader">Utility Bill for Period <%=request("bp")%>, Year <%=request("by")%></span></td>
</table>
<form name="detail" method="post" action="updgutilitybill.asp" onSubmit="zeroNaNs(this)">
<script>
function primbill(caller){
	caller.id = 'white' 
	document.getElementById('primbill').style.display = 'block';
	document.getElementById('secbill').style.display = 'none';
	document.getElementById('primspan').style.backgroundColor = '6699CC';
	document.getElementById('secspan').style.backgroundColor = 'CCCCCC';
}
function secbill(){
	document.getElementById('primbill').style.display = 'none';
	document.getElementById('secbill').style.display = 'block';
	document.getElementById('primspan').style.backgroundColor = 'CCCCCC';
	document.getElementById('secspan').style.backgroundColor = '6699CC';	
}
</script>
<table cellpadding=1 cellspacing=0><tr>
	<td id='primspan' style="background-color:#6699CC">
		<span  onclick="javascript:primbill(this);" onMouseOver="javascript:this.style.color='black';this.style.cursor = 'hand';" 
			onMouseOut="javascript:this.style.color='white';" style="color:white;"><b><u>Primary Billing</u></b>
		</span>
	</td>
	<td id='secspan' style="background-color:#CCCCCC">
		<span  onclick="javascript:secbill(this);" onMouseOver="javascript:this.style.color='black';this.style.cursor = 'hand';" 
			onMouseOut="javascript:this.style.color='white';" style="color:white;"><b><u>Secondary Billing</u></b>
		</span>
	</td>
</tr></table>
<div id="primbill" style="height=205px;">
		<table cellpadding=2 cellspacing=0 width="60%">
			<tr>
				<td style="border-top:2px solid #6699cc;" colspan=3 align="left"> 
					<input type="hidden" name="id" value="<%=id%>">
					<input type="hidden" name="utility" value="<%=utility%>">
					<input type="hidden" name="by" value="<%=by%>">
					<input type="hidden" name="bp" value="<%=bp%>">
					<input type="hidden" name="ypid" value="<%=ypid%>">
					<input type="hidden" name="acctid" value="<%=acctid%>">
					<u>Cost</u>
				</td>
				<td style="border-top:2px solid #6699cc;border-right:2px solid #6699cc;" colspan=2 align="left"><u>Usage</u></td>
			</tr>
			<tr> 
				<td width="30%">Total Bill Amount</td><td width="3%">$</td>
				<td width="27%" ><input type="text" name="tamt2" value="<%=tamt2%>" onKeyUp="calcAvgCostTherm();calcSalesTax();"></td>
				<td width="10%">Therms</td>
				<td width="30%" style="border-right:2px solid #6699cc;"><input type="text" name="therms" value="<%=therms%>" onKeyUp="javascript:updateSecTherms();calcAvgCostTherm();calcFuel();"></td>
			</tr>
			<tr> 
				<td valign="top">Sales Tax</td><td valign="top">$</td>
				<td><input type="text" name="st2" size="6" value="<%=st2%>">&nbsp;%<input type="text" size="6" name="salestaxrate" value="<%=salestaxrate%>" onkeyup="calcSalesTax()"><br>
					<input type="Checkbox" name="taxIncluded" value="1" <%if taxincluded then%>CHECKED<%end if%>>&nbsp;Tax&nbsp;Included
				</td>
				<td>Total (ccf)</td>
				<td style="border-right:2px solid #6699cc;"><input type="text" name="ccf" value="<%=ccf%>" onkeyup="calcTherms()"></td>
			</tr>
			<tr> 
				<td>Avg Cost P/Therm.</td><td>$</td>
				<td><input type="text" name="AvgCostTherm" value="<%=AvgCostTherm%>"></td>
				<td>GRT Rate</td>
				<td style="border-right:2px solid #6699cc;"><input type="hidden" name="GRTAmt" size="6" value="<%=GRTAmt%>"><input type="text" name="GRTrate" value="<%=GRTrate%>"></td>
			</tr>
			<tr> 
				<td style="border-bottom:2px solid #6699cc;">Fuel Adjustment</td>
				<td style="border-bottom:2px solid #6699cc;">$</td>
				<td align="left" style="border-bottom:2px solid #6699cc;"><input type="text" name="fuel" value="<%=fuel%>" size="6">&nbsp;%<input type="text" name="fuelrate" value="<%=fuelrate%>" size="6" onkeyup="calcFuel()"></td>
				<td style="border-bottom:2px solid #6699cc;">Conversion Factor</td>
				<td style="border-right:2px solid #6699cc;border-bottom:2px solid #6699cc;"><input type="text" name="conversion" value="<%=conversion%>" onkeyup="calcTherms()"></td>
			</tr>
		</table>
</div>

<script>
	function isValid(value){
		if (  (value==null)  ||  (  isNaN(value)  )  ||  (value=="")  ) {
			return false;
		}else{
			return true;
		}
	}
	
	function recalcSecUll(){
		var frm = document.forms.detail
		if ( ( isValid(frm.sec_ullp.value) )  &&  (isValid(frm.sec_costpertherm.value)) && (isValid(frm.sec_therm.value)) ){
			var subtotal =  roundNumber(    (parseFloat(frm.sec_costpertherm.value) * parseFloat(frm.sec_therm.value) ), 2)
			var answer = roundNumber((parseFloat(frm.sec_ullp.value) * subtotal),2)
			if (isNaN(answer)){
				answer = 0;
			}			
			frm.sec_ull.value = answer;
		}
	}
	
	function recalcSecUllp(){
		var frm = document.forms.detail
		if ( ( isValid(frm.sec_ull.value) )  &&  (isValid(frm.sec_costpertherm.value)) && (isValid(frm.sec_therm.value)) ){
			var subtotal =  roundNumber(    (parseFloat(frm.sec_costpertherm.value) * parseFloat(frm.sec_therm.value) ), 2)			
			var answer = roundNumber((parseFloat(frm.sec_ull.value) / subtotal),4)
			if (isNaN(answer)){
				answer = 0;
			}
			frm.sec_ullp.value = answer;
		}
	}
	
	function recalcTax(){
		var frm = document.forms.detail
		if ( ( isValid(frm.sec_salestaxp.value) )  &&  (isValid(frm.sec_costpertherm.value)) && (isValid(frm.sec_therm.value)) ){
			var subtotal =  roundNumber(    (parseFloat(frm.sec_costpertherm.value) * parseFloat(frm.sec_therm.value) ), 2)
			var answer = roundNumber((parseFloat(frm.sec_salestaxp.value) * subtotal),2)
			if (isNaN(answer)){
				answer=0;
			}			
			frm.sec_salestax.value = answer;
		}
	}
	
	function recalcTaxp(){
		var frm = document.forms.detail
		if ( ( isValid(frm.sec_salestax.value) )  &&  (isValid(frm.sec_costpertherm.value)) && (isValid(frm.sec_therm.value)) ){
			var subtotal =  roundNumber(    (parseFloat(frm.sec_costpertherm.value) * parseFloat(frm.sec_therm.value) ), 2)
			var answer = roundNumber((parseFloat(frm.sec_salestax.value) / subtotal),4)
			if (isNaN(answer)){
				answer=0;
			}
			frm.sec_salestaxp.value = answer;
		}
	}
	
	function updateSecTotal(){
		var frm = document.forms.detail
		if (isValid(frm.sec_costpertherm.value) == false){
			frm.sec_costpertherm.value = 0;			
		}
		if (isValid(frm.sec_therm.value) == false){
			frm.sec_therm.value = 0;
		}
		if (isValid(frm.sec_transcharge.value) == false){
			frm.sec_transcharge.value = 0;
		}
		if (isValid(frm.sec_salestax.value) == false){
			frm.sec_salestax.value = 0;
		}
		if (isValid(frm.sec_ull.value) == false){
			frm.sec_ull.value = 0;
		}
		var subtotal =  roundNumber(    (parseFloat(frm.sec_costpertherm.value) * parseFloat(frm.sec_therm.value) ), 2)          
		frm.sec_total.value =  roundNumber((subtotal + parseFloat(frm.sec_transcharge.value) + parseFloat(frm.sec_salestax.value) + parseFloat(frm.sec_ull.value)),2);
	}
</script>

<div id="secbill" style="display:none;height=205px;">
		<table cellpadding=2 cellspacing=0 width="60%">
			<tr>
				<td colspan=3 align="left" style="border-top:2px solid #6699cc;"><u>Cost</u></td>
				<td colspan=2 align="left" style="border-top:2px solid #6699cc;border-right:2px solid #6699cc;"><u>Usage</u></td>
			</tr>
			<tr>
				
        <td width="30%">Price per Therm.</td>
				<td width="3%">$</td>
				<td width="27%">
					<input type="text" name="sec_costpertherm" onChange="javascript:recalcTax();recalcSecUll();updateSecTotal();" 
						onKeyUp="javascript:recalcSecUll();recalcTax();updateSecTotal();" value="<%=formatnumber(sec_CostPerTherm,6,-2,0,0)%>">
				</td>
				<td width="10%">Therms</td>
				<td width="30%" style="border-right:2px solid #6699cc;">
					<input type="text" name="sec_therm" onChange="javascript:updatePriTherms();recalcSecUll();recalcTax();updateSecTotal();" 
						onKeyUp="javascript:updatePriTherms();recalcSecUll();recalcTax();updateSecTotal();" 
						value="<%=therms%>">
				</td>
			</tr>
			<tr>
				<td>Transportation Charge</td>
				<td>$</td>
				<td colspan=3 align="left" style="border-right:2px solid #6699cc;">
					<input type="text" onChange="javascript:updateSecTotal();" onKeyUp="javascript:updateSecTotal();" name="sec_transcharge" 
						value="<%=formatnumber(sec_TransCharge,2,-2,0,0)%>">
				</td>
			</tr>
			<tr>
				<td>Sales tax percentage</td>
				<td>%</td>
				<td colspan=1 align="left">
					<input type="text" onChange="javascript:recalcTax();updateSecTotal();" onKeyUp="javascript:recalcTax();updateSecTotal();" name="sec_salestaxp" 
						value="<%=formatnumber(sec_salestaxp,4,-2,0,0)%>">
				</td>
				<td colspan=2 style="border-right:2px solid #6699cc;"><i>(please enter as a decimal)</i></td>				
			</tr>			
			<tr>
				<td>Sales tax/comp use tax</td>
				<td>$</td>
				<td colspan=3 align="left" style="border-right:2px solid #6699cc;">
					<input type="text" onChange="javascript:recalcTaxp();updateSecTotal();" onKeyUp="javascript:recalcTaxp();updateSecTotal();" name="sec_salestax" 
						value="<%=formatnumber(sec_salestax,2,-2,0,0)%>">
				</td>				
			</tr>

			<tr>
				<td>Utility line loss percentage</td>
				<td>%</td>
				<td>
					<input type="text" name="sec_ullp" onChange="javascript:recalcSecUll();updateSecTotal();" onKeyUp="javascript:recalcSecUll();updateSecTotal();" 
						value="<%=formatnumber(sec_ullp,4,-2,0,0)%>">
				</td>
				<td colspan=2 style="border-right:2px solid #6699cc;"><i>(please enter as a decimal)</i></td>
			</tr>
			<tr>
				<td>Utility line loss</td>
				<td>$</td>
				<td colspan=3 align="left" style="border-right:2px solid #6699cc;">
					<input type="text" name="sec_ull" onChange="javascript:updateSecTotal();recalcSecUllp();" onKeyUp="javascript:updateSecTotal();recalcSecUllp();" 
						value="<%=formatnumber(sec_ull,2,-2,0,0)%>">
				</td>				
			</tr>
			<tr>
				<td style="border-bottom:2px solid #6699cc;">Total Bill Amt</td>
				<td style="border-bottom:2px solid #6699cc;">$</td>
				<td align="left" colspan="3" style="border-bottom:2px solid #6699cc;border-right:2px solid #6699cc;">
					<input type="text" name="sec_total" value="<%=formatnumber(sec_total,2,-2,0,0)%>">
				</td>
			</tr>							
		</table>
</div>	
<br><br>

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
<tr><td><!-- GL CODE  --><input name="glcode" type="hidden" size="10" maxlength="20"><!-- </font> --></td>
</table>
<!--<input type="button" name="download" value="DOWNLOAD DATA">-->
<input type="hidden" name="bldg" value="<%=bldg%>">
</form>

</body>
</html>
<%
rst1.close
set cnn1=nothing
%>
