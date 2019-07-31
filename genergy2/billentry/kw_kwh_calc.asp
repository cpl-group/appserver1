<html>
<head>
<title>Calc Fields</title>
<script>
function calccostkwh()
{	frm = document.calcfields
	onpeak = parseFloat(frm.onpeak.value)
	if(isNaN(onpeak)) onpeak=0
	offpeak = parseFloat(frm.offpeak.value)
	if(isNaN(offpeak)) offpeak=0
	frm.costkwh.value = onpeak+offpeak
	window.opener.document.forms[0].costkwh.value = frm.costkwh.value
	window.opener.calcCostKWHTotal();
	window.opener.calcTDTotalAmt();
	window.opener.calcTDgrt(window.opener.i);
	window.opener.calcUnitCostKwh();
	window.opener.calcTDunitcostkwh();
	window.opener.checkvalues();
}

function calccostkw()
{	frm = document.calcfields
	totaldemand = parseFloat(frm.totaldemand.value)
	if(isNaN(totaldemand)) totaldemand=0
	primary = parseFloat(frm.primary.value)
	if(isNaN(primary)) primary=0
	secondary = parseFloat(frm.secondary.value)
	if(isNaN(secondary)) secondary=0
	frm.costkw.value = totaldemand+primary+secondary
	window.opener.document.forms[0].costkw.value = frm.costkw.value
	window.opener.calcTDTotalAmt();
	window.opener.calcTDgrt(window.opener.i);
	window.opener.calcUnitCostKw();
	window.opener.calcTDunitcostkw();
	window.opener.checkvalues();
}
</script>
</head>
<body style="font-size:12px;font-family:arial">
<form name="calcfields">
<%if trim(request("calc"))="costkwh" then%>
Cost KWH<br>
<table style="font-size:10px;font-family:arial">
<tr><td>On&nbsp;Peak</td><td><input name="onpeak" type="text" size="10" onkeyup="calccostkwh()"></td></tr>
<tr><td>Off&nbsp;Peak</td><td><input name="offpeak" type="text" size="10" onkeyup="calccostkwh()"></td></tr>
<tr><td>Cost&nbsp;KWH</td><td><input name="costkwh" type="text" size="10" readonly></td></tr>
<tr><td></td><td><input type="button" value="Close" onclick="window.opener.focus();window.close();"></td></tr>
</table>
<%elseif trim(request("calc"))="costkw" then%>
Cost KW<br>
<table style="font-size:10px;font-family:arial">
<tr><td>Total&nbsp;Demand</td><td><input name="totaldemand" type="text" size="10" onkeyup="calccostkw()"></td></tr>
<tr><td>Primary</td><td><input name="primary" type="text" size="10" onkeyup="calccostkw()"></td></tr>
<tr><td>Secondary</td><td><input name="secondary" type="text" size="10" onkeyup="calccostkw()"></td></tr>
<tr><td>Cost&nbsp;KW</td><td><input name="costkw" type="text" size="10" readonly></td></tr>
<tr><td></td><td><input type="button" value="Close" onclick="window.opener.focus();window.close();"></td></tr>
</table>
<%end if%>
</form>
</body>
</html>
