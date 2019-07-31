<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%option explicit
dim acctid, bldg, utility, ypid, by, bp
acctid=Request.querystring("acctid")
bldg=Request.querystring("bldg")
utility=Request.querystring("utility")
ypid=Request.querystring("ypid")
by=Request.querystring("by")
bp=Request.querystring("bp")

Dim cnn1, rst1, rst2, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_genergy1")
rst1.ActiveConnection = cnn1
'session("roleid")=4

dim grandtotal, grandavg, granduckw, granduckwh
grandtotal = 0
grandavg = 0
granduckw = 0
granduckwh = 0
rst1.open "SELECT sum(totalbillamt) as totalbillamt, (case when sum(totalkwh)=0 then '0' else sum(totalbillamt)/sum(totalkwh)end) as [avg], (case when sum(totalkwh)=0 then '0' else (sum(tdcostkwh)+sum(EscoBillAmt))/sum(totalkwh)end) as unitcostkwh, (case when sum(totalkw)=0 then '0' else sum(CostKW)/sum(totalkw)end) as unitcostkw FROM utilitybill WHERE ypid="&ypid
if not(isnull(rst1("totalbillamt"))) then grandtotal = rst1("totalbillamt")
if not(isnull(rst1("avg"))) then grandavg = rst1("avg")
if not(isnull(rst1("unitcostkw"))) then granduckw = rst1("unitcostkw")
if not(isnull(rst1("unitcostkwh"))) then granduckwh = rst1("unitcostkwh")
rst1.close

sqlstr= "select * from utilitybill where ypid='"&ypid&"' and acctid='"&acctid&"'"
'response.write sqlstr
'response.end
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1

dim Escoref 'get escoref if has one
Escoref = ""
rst2.open "SELECT Escoref FROM tblacctsetup WHERE acctid='"& acctid &"'", cnn1
if not rst2.EOF then
	Escoref=rst2("Escoref")
end if
rst2.close
%>
<html>
<head>
<title>Account Detail</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function roundNumber(num,precision)
{	var i,sZeros = '';
	for(i = 0;i < precision;i++)
		sZeros += '0';
	i = Number(1 + sZeros);
	return Math.round(num * i) / i;
}

var noCalcTotalKwh = 0;
function calcTotalKwh()
{	var frm = document.forms[0];
	if(noCalcTotalKwh)
	{	if(!confirm("Recalculate Total KWH?")){return;}
		else{frm.totalkwh.focus();}
	}
	frm.totalkwh.value = (parseFloat(frm.offpeakkwh.value)+parseFloat(frm.onpeakkwh.value));
	frm.totalkwhcom.value = frm.totalkwh.value;
	calcUnitCredit();
	calcUnitCostKwh();
	calcAverageCost();
	calcTDunitcostkwh();
	calcCOunitcostkwh();
	noCalcTotalKwh = 0;
}

function calcUnitCostKwh()
{	var frm = document.forms[0];
	var comod = parseFloat(frm.totalcommodity.value);
	if(isNaN(comod)) comod=0;
	frm.unitcostkwh.value = roundNumber((parseFloat(frm.costkwh.value)+comod)/parseFloat(frm.totalkwh.value),4)
}

function calcUnitCostKw()
{	var frm = document.forms[0];
	frm.unitcostkw.value = roundNumber(parseFloat(frm.costkw.value)/parseFloat(frm.totalkw.value),4)
}

function calcTotalBillAmt()
{	var frm = document.forms[0];
	var comod = parseFloat(frm.totalcommodity.value);
	if(isNaN(comod)) comod=0;
	frm.totalbillamt.value = parseFloat(frm.TDtotalamt.value)+comod;
	calcAverageCost()
}

function calcAverageCost()
{	var frm = document.forms[0];
	frm.averagecost.value = roundNumber(parseFloat(frm.totalbillamt.value)/parseFloat(frm.totalkwh.value),4);
}

function calcTDunitcostkwh()
{	var frm = document.forms[0];
	frm.TDunitcostkwh.value = roundNumber(parseFloat(frm.costkwh.value)/parseFloat(frm.totalkwh.value),4);
}

function calcTDunitcostkw()
{	var frm = document.forms[0];
	frm.TDunitcostkw.value = roundNumber(parseFloat(frm.costkw.value)/parseFloat(frm.totalkw.value),4);
}

var noCalcTDTotalAmt = 0;
function calcTDTotalAmt()
{	var frm = document.forms[0];
	if(noCalcTDTotalAmt)
	{	if(!confirm("Recalculate Total Total T&D?")){return;}
		else{frm.TDtotalamt.focus();}
	}
	frm.TDtotalamt.value = parseFloat(frm.costkw.value)+parseFloat(frm.costkwh.value);
	noCalcTDTotalAmt = 0;
	calcTotalBillAmt();
}

function calcTDgrt(i)
{	var frm = document.forms[0];
	var TDtotalamt = parseFloat(frm.costkw.value)+parseFloat(frm.costkwh.value);
	if((isNaN(TDtotalamt))&&(!frm.TDwithtax.checked)){TDtotalamt = parseFloat(frm.TDtotalamt.value);}
	if(!isNaN(TDtotalamt))
	{	if(i.name=='TDgrtamt')
		{	frm.TDgrtpercent.value = roundNumber(parseFloat(frm.TDgrtamt.value)/TDtotalamt,6)
		}else
		{	frm.TDgrtamt.value = roundNumber(parseFloat(frm.TDgrtpercent.value) * TDtotalamt,2)
			frm.grossrecieptSub.value=frm.TDgrtamt.value;
		}
	}
}

function calcCOunitcostkwh()
{	var frm = document.forms[0];
	frm.COunitcostkwh.value = roundNumber(parseFloat(frm.totalcommodity.value)/parseFloat(frm.totalkwhcom.value),4);
}

function calcCOgrt(i)
{	var frm = document.forms[0];
	var pretaxcost = parseFloat(frm.pretaxcost.value);
	if(!isNaN(pretaxcost))
	{	if(i.name=='COgrtamt')
		{	frm.COgrtpercent.value = roundNumber(parseFloat(frm.grossreceipt.value)/parseFloat(frm.pretaxcost.value),6)
		}else
		{	frm.grossreceipt.value = roundNumber(parseFloat(frm.COgrtpercent.value) * parseFloat(frm.pretaxcost.value),2)
		}
	}
}

function calcCOsales(i,tax)
{	var frm = document.forms[0];
	var pretaxcost = parseFloat(frm.pretaxcost.value)+parseFloat(frm.grossreceipt.value);
	if(!isNaN(pretaxcost))
	{	if(i.name=='salestax')
		{	frm.COsalespercent.value = roundNumber(parseFloat(frm.salestax.value)/pretaxcost,6)
		}else
		{	frm.salestax.value = roundNumber(parseFloat(frm.COsalespercent.value) * pretaxcost,2)
		}
	}
	if(tax!='justtax')
	{	frm.totalcommodity.value = roundNumber(pretaxcost+parseFloat(frm.salestax.value),2)
		calcTotalBillAmt();
		calcUnitCostKwh();
		calcCOunitcostkwh();
	}
}

function calcfueladjSub()
{	var frm = document.forms[0];
	var MAC = parseFloat(frm.MAC.value);
	var MSC = parseFloat(frm.MSC.value);
	if(isNaN(MAC)) MAC=0;
	if(isNaN(MSC)) MSC=0;
	frm.fueladjustmentSub.value = MAC + MSC;
}

function calcUnitCredit()
{	var frm = document.forms[0];
	frm.unitcredit.value = roundNumber(parseFloat(frm.lmep.value)/parseFloat(frm.totalkwh.value),4);
}

function calcpretaxcost()
{	var frm = document.forms[0];
	frm.pretaxcost.value = roundNumber(parseFloat(frm.fixedrate.value) * parseFloat(frm.totalkwhcom.value),2);
	calcCostKWHTotal()
}

function calcfixedrate()
{	var frm = document.forms[0];
	frm.fixedrate.value = roundNumber(parseFloat(frm.pretaxcost.value) / parseFloat(frm.totalkwhcom.value),4);
}

function calcCostKWHTotal()
{	var frm = document.forms[0];
	var comod = parseFloat(frm.totalcommodity.value);
	if(isNaN(comod)) comod=0;
	frm.costkwhtotal.value = comod + parseFloat(frm.costkwh.value)
}

/////////////////////////end of calc/////////////////////////////
function checkTaxBox(cbox)
{	var frm = document.forms[0];
	if(cbox.checked)
	{	frm.TDwithtax.checked = 1
		frm.COwithtax.checked = 1
		frm.TotalIncludeTax.checked = 1
	}else
	{	frm.TDwithtax.checked = 0
		frm.COwithtax.checked = 0
		frm.TotalIncludeTax.checked = 0
	}
}

function zeroNaNs(frm)
{	for(i=0;i<frm.elements.length;i++)
	{	//alert(frm.elements[i].name+':'+(frm.elements[i].type=='text')+((isNaN(frm.elements[i].value))||(!(isFinite(frm.elements[i].value)))) )
		if((frm.elements[i].type=='text')&&(isNaN(frm.elements[i].value)||!(isFinite(frm.elements[i].value))))
		{	frm.elements[i].value=''
		}
	}
}


function checkvalues()
{	var fields = document.forms[0].elements
	for(i=0;i<fields.length;i++)
	{	if(fields[i].type=='text')
		{	var inputtext = fields[i].value;
			if((isNaN(inputtext))||(inputtext==''))
			{	fields[i].style.backgroundColor='#FFCCCC';
			}
			else
			{	fields[i].style.backgroundColor='#FFFFFF';
			}
		}
	}
}


function changeTab(hottab,showdiv)
{	document.all['tabDel'].style.backgroundColor='#CCCCCC';
	document.all['tabTot'].style.backgroundColor='#CCCCCC';
	document.all['delivery'].style.visibility='hidden';
	document.all['delivery'].style.position='absolute';
	document.all['totals'].style.visibility='hidden';
	document.all['totals'].style.position='absolute';
<%if trim(Escoref)<>"0" then%>
	document.all['tabCom'].style.backgroundColor='#CCCCCC';
	document.all['commodity'].style.visibility='hidden';
	document.all['commodity'].style.position='absolute';
<%end if
  if session("roleid")=4 then%>
	document.all['tabSub'].style.backgroundColor='#CCCCCC';
	document.all['submeter'].style.visibility='hidden';
	document.all['submeter'].style.position='absolute';
<%end if%>
	document.all[hottab].style.backgroundColor='#0099FF';
	document.all[showdiv].style.visibility='visible';
	document.all[showdiv].style.position='relative';
}
</script>
</head>
<body bgcolor="#FFFFFF" onload="checkvalues();">
<%
dim id, fueladj ,grossreceipt ,salestax ,offpeakkwh ,onpeakkwh ,totalkwh ,costkwh ,totalkw ,costkw
dim fueladjustmentSub ,saletaxSub ',grossrecieptSub 
dim escoFixedRate ,EscoPreTax ,EscoST ,EscoGR ,EscoBillAmt ,EscoAcct 
dim unitcostkwh ,unitcostkw ,avgCost ,totalbillamt
dim TDtotalamt, TDwithtax, TDunitcostkwh, TDsalestax, TDunitcostkw, TDgrtpercent, TDgrtamt, COunitcostkwh, MSC, COgrtpercent, MAC, TotalIncludeTax, TDcostKWH
dim COwithTax, lmep, unitcredit, TDsalesamt, MACdollar, EscoSTdollar, fueladjdollar
if not rst1.EOF then
	id = rst1("id")
	ypid = rst1("ypid")
	if trim(rst1("acctid"))<>"" then acctid = rst1("acctid")
	Utility = rst1("Utility")

	fueladj = rst1("fueladj")
	grossreceipt = rst1("grossreceipt")
	salestax = rst1("salestax")
	offpeakkwh = rst1("offpeakkwh")
	onpeakkwh = rst1("onpeakkwh")
	totalkwh = rst1("totalkwh")
	costkwh = rst1("costkwh")
	totalkw = rst1("totalkw")
	costkw = rst1("costkw")

	fueladjustmentSub = rst1("FuelAdj")
	saletaxSub = rst1("SalesTax")
'	grossrecieptSub = rst1("GrossReceipt")

	escoFixedRate = rst1("escoFixedRate")
	EscoPreTax = rst1("EscoPreTax")
	EscoST = rst1("EscoST")
	EscoGR = rst1("EscoGR")
	EscoBillAmt = rst1("EscoBillAmt")
	EscoAcct = rst1("EscoAcct")

	unitcostkwh = rst1("unitcostkwh")
	unitcostkw = rst1("unitcostkw")
	avgCost = rst1("avgkwh")
	totalbillamt = rst1("totalbillamt")

	TDtotalamt = rst1("TDtotalamt")
	TDwithtax = rst1("tdwithtax")
	TDunitcostkwh = rst1("TDunitcostkwh")
	TDsalestax = rst1("TDsalestax")
	TDunitcostkw = rst1("TDunitcostkw")
	TDgrtpercent = rst1("TDgrtpercent")
	TDgrtamt = rst1("grossreceipt")
	COunitcostkwh = rst1("COunitcostkwh")
	MSC = rst1("MSC")
	COgrtpercent = rst1("COgrtpercent")
	COwithtax = rst1("COwithtax")
	MAC = rst1("MAC")
	TotalIncludeTax = rst1("TotalIncludeTax")
	TDcostKWH = rst1("TDcostKWH")

	lmep = rst1("lmepcredit")
	unitcredit = rst1("unit_credit")
	TDsalesamt = rst1("TDsalesamt")
	MACdollar = rst1("MACdollar")
	EscoSTdollar = rst1("EscoSTdollar")
	fueladjdollar = rst1("fueladjdollar")
%>
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Utility Bill for Period</i></font><b><font size="4" color="#FFFFFF"> <%=Request.querystring("bp")%></font></b><font size="4" color="#FFFFFF"><i>Year</i></font><b><font size="4" color="#FFFFFF"> 
        <%=Request.querystring("by")%></font></b></font></div>
    </td>
  </tr>
</table>
<%else%>
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Enter Utility Bill for Period</i></font><b><font size="4" color="#FFFFFF"><i> 
        </i> <%=Request.querystring("bp")%> </font></b><font size="4" color="#FFFFFF"><i>Year</i></font><b><font size="4" color="#FFFFFF"> 
        </font><font face="Arial, Helvetica, sans-serif"><b><font size="4" color="#FFFFFF"><%=Request.querystring("by")%></font></b></font></b></font></div>
    </td></tr>
</table>
<%end if%>
<form name="detail" method="post" action="updutilitybill.asp" onsubmit="zeroNaNs(this)">
<table border="0" cellspacing="0" cellpadding="0">
<tr style="font-family:Arial, Helvetica, sans-serif; font-size:13">
	<td id="tabDel" style="background-color:#0099FF">&nbsp;<b><a href="javascript:changeTab('tabDel','delivery');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">T&nbsp;&amp;&nbsp;D</a></b>&nbsp;</td>
<%if Escoref<>"0" then%>
	<td id="tabCom" style="background-color:#CCCCCC">&nbsp;<b><a href="javascript:changeTab('tabCom','commodity');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">Commodity&nbsp;Charge</a></b>&nbsp;</td>
<%end if
if session("roleid")=4 then%>
	<td id="tabSub" style="background-color:#CCCCCC">&nbsp;<b><a href="javascript:changeTab('tabSub','submeter');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">Submetered</a></b>&nbsp;</td>
<%end if%>
	<td id="tabTot" style="background-color:#CCCCCC">&nbsp;<b><a href="javascript:changeTab('tabTot','totals');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">Totals</a></b>&nbsp;</td>
</tr>
</table>
<div id="delivery" style="border: 2px solid #0099FF;height:200;position:relative;visibility:visible">
<table width="100%" border="0" style="font-family:Arial, Helvetica, sans-serif; color:black">
<tr>
<td width="15%">
<input type="hidden" name="id1" value="<%=id%>">
<input type="hidden" name="by" value="<%=Request.querystring("by")%>">
<input type="hidden" name="bp" value="<%=Request.querystring("bp")%>">
<input type="hidden" name="ypid" value="<%=ypid%>">
<input type="hidden" name="acctid" value="<%=acctid%>">
<input type="hidden" name="escoref" value="<%=escoref%>">
<input type="hidden" name="utility" value="<%=utility%>">
<u>KWH</u></td><td width="1%"></td>
<td width="24%"></td>
<td width="15%"><u>KW</u></td><td width="1%"></td>
<td width="24%">&nbsp;</td>
<td width="15%"></td>
<td width="24%">&nbsp;</td>
</tr>

<tr>
<td>On Peak KWH</td><td></td>
<td> <input type="text" name="onpeakkwh" value="<%=onpeakkwh%>" onKeyUp="calcTotalKwh();checkvalues();"></td>
<td>Total KW</td><td></td>
<td><input type="tkw" name="totalkw" value="<%=totalkw%>" onKeyUp="calcUnitCostKw();calcTDunitcostkw();checkvalues();"></td>
</tr>

<tr>
<td>Off Peak KWH</td><td></td>
<td> <input type="text" name="offpeakkwh" value="<%=offpeakkwh%>" onKeyUp="calcTotalKwh();checkvalues();"></td>
<td>Total KW Cost</td><td>$</td>
<td><input type="text" name="costkw" value="<%=costkw%>" onKeyUp="calcTDTotalAmt();calcTDgrt(i);calcUnitCostKw();calcTDunitcostkw();checkvalues();">&nbsp;<a href="#" onclick="window.open('kw_kwh_calc.asp?calc=costkw','calc','width=180,height=180');" style="font-size:10;px">calc</a></td>
</tr>

<tr>
<td>Total KWH</td><td></td>
<td><input type="text" name="totalkwh" value="<%=totalkwh%>" onChange="noCalcTotalKwh=1;" onKeyUp="calcUnitCredit();totalkwhcom.value=this.value;calcUnitCostKwh();calcAverageCost();calcTDunitcostkwh();checkvalues();"></td>
<td><%if Escoref<>"0" then%>MAC<%else%>Full Service Fuel<%end if%></td><td>$</td>
<td><%if Escoref<>"0" then%><input type="text" name="MACdollar" value="<%=MACdollar%>" onKeyUp="checkvalues();">&nbsp;Rate&nbsp;$&nbsp;<input type="text" name="MAC" value="<%=MAC%>" onKeyUp="calcfueladjSub();checkvalues();"><%else%><input type="text" name="fueladjdollar" value="<%=fueladjdollar%>" onKeyUp="checkvalues();">&nbsp;&nbsp;%&nbsp;<input type="text" name="fueladjustmentSub" value="<%=fueladjustmentSub%>" onKeyUp="checkvalues();"><input type="hidden" name="MAC" value="<%=MAC%>" onKeyUp="calcfueladjSub();checkvalues();"><%end if%></td>
</tr>

<tr>
<td>Cost KWH</td><td>$</td>
<td><input type="text" name="costkwh" value="<%=TDcostkwh%>" onKeyUp="calcCostKWHTotal();calcTDTotalAmt();calcTDgrt(i);calcUnitCostKwh();calcTDunitcostkwh();checkvalues();">&nbsp;<a href="#" onclick="window.open('kw_kwh_calc.asp?calc=costkwh','calc','width=180,height=150');" style="font-size:10;px">calc</a></td>
<td>Sales Tax</td><td>$</td>
<td><input type="text" name="TDsalesamt" value="<%=TDsalesamt%>" onKeyUp="checkvalues();">&nbsp;Raw&nbsp;%&nbsp;<input type="text" name="TDsalestax" value="<%=TDsalestax%>" onKeyUp="checkvalues();"></td>
</tr>

<tr>
<td><nobr>T&nbsp;&amp;&nbsp;D Unit Cost KWH</nobr></td><td>$</td>
<td><input type="text" name="TDunitcostkwh" value="<%=TDunitcostkwh%>" onKeyUp="checkvalues();"></td>
<td><nobr>T&nbsp;&amp;&nbsp;D Total Bill</nobr></td><td>$</td>
<td><input type="text" name="TDtotalamt" value="<%=TDtotalamt%>" onChange="noCalcTDTotalAmt=1;" onKeyUp="calcTDgrt(i);calcTotalBillAmt();checkvalues();">&nbsp;<input type="checkbox" name="TDwithtax" value="1" onclick="checkTaxBox(this);"<%if trim(TDwithtax)="1" then response.write " CHECKED"%>>&nbsp;inclusive&nbsp;of&nbsp;tax</td>
</tr>

<tr>
<td>T&nbsp;&amp;&nbsp;D Unit Cost KW</td><td>$</td>
<td><input type="text" name="TDunitcostkw" value="<%=TDunitcostkw%>" onKeyUp="checkvalues();"></td>
<td></td><td></td>
<td><input type="hidden" name="TDgrtamt" value="<%=TDgrtamt%>" onKeyUp="calcTDgrt(this);grossrecieptSub.value=this.value;checkvalues();"><!-- &nbsp;Raw&nbsp;% --><input type="hidden" name="TDgrtpercent" value="<%=TDgrtpercent%>" onKeyUp="calcTDgrt(this);checkvalues();" size="10"></td>
</tr>

<tr>
<td>LMEP</td><td>$</td>
<td><input type="text" name="lmep" value="<%=lmep%>" onKeyUp="calcUnitCredit();checkvalues();"></td>
<td></td><td></td>
<td></td>
</tr>

<tr>
<td>Unit Credit</td><td>$</td>
<td><input type="text" name="unitcredit" value="<%=unitcredit%>" readonly></td>
<td></td><td></td>
<td></td>
</tr>

</table>
</div>
<div id="commodity" style="border: 2px solid #0099FF;height:200;position:absolute;visibility:hidden">
<table width="100%" border="0" style="font-family:Arial, Helvetica, sans-serif; color:black">
<tr>
<td width="15%"><u>KWH</u></td><td width="1%"></td>
<td width="24%"></td>
<td width="15%"><u>COST</u></td><td width="1%"></td>
<td width="24%">&nbsp;</td>
<td width="15%"></td>
<td width="24%">&nbsp;</td>
</tr>

<tr>
<td>Total KWH</td><td></td>
<td><input type="text" name="totalkwhcom" value="<%=totalkwh%>" onChange="noCalcTotalKwh=1;" onKeyUp="totalkwh.value=this.value;calcUnitCostKwh();calcCOunitcostkwh();calcAverageCost();checkvalues();"></td>
<td>Sub Total</td><td>$</td>
<td><input type="text" name="pretaxcost" value="<%=EscoPreTax%>" onKeyUp="checkvalues();">&nbsp;&nbsp;MSC&nbsp;$&nbsp;<input type="text" name="MSC" value="<%=MSC%>" onKeyUp="calcfueladjSub();checkvalues();"></td>
<td></td>
<td></td>
</tr>

<tr>
<td>Fixed Rate</td><td>$</td>
<td><input type="text" name="fixedrate" value="<%=escoFixedRate%>" onKeyUp="checkvalues();"></td>
<td>Gross Receipt</td><td>$</td>
<td><input type="text" name="grossreceipt" value="<%=EscoGR%>" onKeyUp="checkvalues();">&nbsp;Raw&nbsp;%&nbsp;<input type="text" name="COgrtpercent" value="<%=COgrtpercent%>" onKeyUp="checkvalues();" size="10"></td>
</tr>

<tr>
<td><nobr>Comm. Unit Cost KWH</nobr></td><td>$</td>
<td><input type="text" name="COunitcostkwh" value="<%=COunitcostkwh%>" onKeyUp="checkvalues();"></td>
<td>Sales Tax</td><td>$</td>
<td><input type="text" name="EscoSTdollar" value="<%=EscoSTdollar%>" onKeyUp="checkvalues();" size="10">&nbsp;Raw&nbsp;%&nbsp;<input type="text" name="EscoST" value="<%=EscoST%>" onKeyUp="checkvalues();"></td>
</tr>

<tr>
<td></td><td></td>
<td></td>
<td><nobr>Total commodity</nobr></td><td>$</td>
<td><input type="text" name="totalcommodity" value="<%=EscoBillAmt%>" onKeyUp="calcCostKWHTotal();calcTotalBillAmt();calcUnitCostKwh();calcCOunitcostkwh();checkvalues();">&nbsp;<input type="checkbox" name="COwithtax" value="1" onclick="checkTaxBox(this);"<%if trim(TDwithtax)="1" then response.write " CHECKED"%>>&nbsp;inclusive of tax</td>
</tr>

</table>
</div>
<div id="submeter" style="border: 2px solid #0099FF;height:200;position:absolute;visibility:hidden">
<table border="0" style="font-family:Arial, Helvetica, sans-serif; color:black">
<tr><td><u>Submeter</u></td></tr>
<tr><td width="15%"><%=utility%>&nbsp;Adjustment</td><td width="1%"></td>
	<td width="24%"><input type="text" <%if Escoref<>"0" then%>name="fueladjustmentSub"<%end if%> value="<%=fueladjustmentSub%>" readonly onKeyUp="checkvalues();"><nobr> (sum of MSC and MAC)</nobr></td><td></td></tr>
<tr><td width="15%">Sales&nbsp;Tax</td><td width="1%">%</td>
	<td width="24%"><input type="text" name="saletaxSub" value="<%=saletaxSub%>" onKeyUp="checkvalues();"></td></tr>
<tr><td width="15%">Gross&nbsp;Receipt</td><td width="1%">%</td>
	<td width="24%"><input type="text" name="grossrecieptSub" value="<%=TDgrtamt%>" onKeyUp="checkvalues();"></td></tr>
</table>
</div>
<div id="totals" style="border: 2px solid #0099FF;height:200;position:absolute;visibility:hidden">
<table border="0" style="font-family:Arial, Helvetica, sans-serif; color:black">
<tr><td>Unit Cost KWH</td><td>$</td>
	<td><input type="text" name="unitcostkwh" value="<%=unitcostkwh%>" onKeyUp="checkvalues();"></td>
	<td>Average Cost</td><td>$</td>
	<td><input type="text" name="averagecost" value="<%=avgCost%>" onKeyUp="checkvalues();"></td></tr>
<tr valign="top"><td>Unit Cost KW</td><td>$</td>
	<td><input type="text" name="unitcostkw" value="<%=unitcostkw%>" onKeyUp="checkvalues();"></td>
	<td>Total Bill Amt Cost</td><td>$</td>
	<td><input type="text" name="totalbillamt" value="<%=totalbillamt%>" onKeyUp="checkvalues();"></td>
	<td><input type="checkbox" name="TotalIncludeTax" value="1" onclick="checkTaxBox(this);"<%if trim(TDwithtax)="1" then response.write " CHECKED"%>>&nbsp;inclusive of tax</td>
</tr>
<tr><td>&nbsp;</td>
<tr><td>Building Totals</td>
<tr><td>Unit Cost KWH</td><td>$</td>
	<td><input type="text" name="" value="<%=formatnumber(granduckwh,5)%>" onKeyUp="checkvalues();"></td>
	<td>Average Cost</td><td>$</td>
	<td><input type="text" name="" value="<%=formatnumber(grandavg,2)%>" onKeyUp="checkvalues();"></td></tr>
<tr valign="top"><td>Unit Cost KW</td><td>$</td>
	<td><input type="text" name="" value="<%=formatnumber(granduckw,5)%>" onKeyUp="checkvalues();"></td>
	<td>Total Bill Amt Cost</td><td>$</td>
	<td><input type="text" name="" value="<%=grandtotal%>" onKeyUp="checkvalues();"></td>
</tr>
</table>
</div>&nbsp;<br><input type="hidden" name="costkwhtotal" value="<%=costkwh%>">
<font size="2" face="Arial, Helvetica, sans-serif"> GL CODE <input name="glcode" type="text" size="10" maxlength="20"></font> 

<%
'if session("roleid")=4 and locked<>"True" then
if not rst1.EOF then%>
  <input type="submit" name="action" value="UPDATE">
  	  <input type="button" name="download" value="DOWNLOAD DATA">
<%else%>
  <input type="submit" name="action" value="SAVE">
<%end if%>
<%'end if%>
</form>
<%
rst1.close
set cnn1=nothing
%>
</body>
</html>