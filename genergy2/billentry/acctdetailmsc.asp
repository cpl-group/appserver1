<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

dim acctid, bldg, utility, ypid, by, bp, pid,timestamp
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

cnn1.Open getLocalConnect(bldg)

rst1.ActiveConnection = cnn1
'session("roleid")=4

dim grandtotal, grandavg, granduckw, granduckwh
grandtotal = 0
grandavg = 0
granduckw = 0
granduckwh = 0
'rst1.open "SELECT sum(totalbillamt) as totalbillamt, (case when sum(totalkwh)=0 then '0' else sum(totalbillamt)/sum(totalkwh)end) as [avg], (case when sum(totalkwh)=0 then '0' else (sum(tdcostkwh)+sum(EscoBillAmt))/sum(totalkwh)end) as unitcostkwh, (case when sum(totalkw)=0 then '0' else sum(CostKW)/sum(totalkw)end) as unitcostkw FROM utilitybill WHERE ypid="&ypid
rst1.open "SELECT sum(totalbillamt) as totalbillamt, (case when sum(totalkwh)=0 then '0' else sum(totalbillamt)/sum(totalkwh)end) as [avg], (case when sum(totalkwh)=0 then '0' else (sum(tdcostkwh)+sum(EscoBillAmt)+ (select case when sum(isnull(tdkwhsalesamt,0)) > 0 then sum(isnull(tdkwhsalesamt,0)) else 0 end from utilitybill where tdwithtax = 1 and ypid ="&ypid&"))/sum(totalkwh)end) as unitcostkwh,(case when sum(totalkw)=0 then '0' else (sum(CostKW)+ (select case when sum(isnull(tdsalesamt,0))>0 then sum(isnull(tdsalesamt,0)) else 0 end from utilitybill where tdwithtax = 1 and ypid = "&ypid&"))/sum(totalkw)end) as unitcostkw  FROM utilitybill WHERE ypid ="&ypid
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
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		
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
	calcfixedrate();
	noCalcTotalKwh = 0;
}

function calcUnitCostKwh()
{	var frm = document.forms[0];
	var comod = parseFloat(frm.totalcommodity.value);
	if(isNaN(comod)) comod=0;
    var kwhtax = (frm.TDwithtax.checked ? parseFloat(frm.TDkwhsalesamt.value) : 0) 
    if(isNaN(kwhtax)) kwhtax=0;
	frm.unitcostkwh.value = roundNumber((parseFloat(frm.costkwh.value)+comod+kwhtax)/parseFloat(frm.totalkwh.value),6)
}

function checksalestax()
{	var frm = document.forms[0];
	var salestax = parseFloat(frm.saletaxSub.value);
    var bldg = frm.bldg.value;
    if(salestax < .07 && bldg != "30E60" )
    { alert("Sales tax cannot be lower than .07")
      setTimeout(function(){frm.saletaxSub.focus();}, 1);
    }
    if(salestax > .09)
    { alert("Sales tax cannot be higher than .09")
      setTimeout(function(){frm.saletaxSub.focus();}, 1);
    }
	
}

function checkgrossreceipt()
{	var frm = document.forms[0];
    var grossreceipt = parseFloat(frm.grossrecieptSub.value);
    if(grossreceipt < .02)
    { alert("gross receipt cannot be lower than .02")
      setTimeout(function(){frm.grossrecieptSub.focus();}, 1);
    }
    if(grossreceipt > .06)
    { alert("gross receipt cannot be higher than .06")
      setTimeout(function(){frm.grossrecieptSub.focus();}, 1);
    }
	
}

function checkgrossreceiptsupply()
{	
    var frm = document.forms[0];
    
    var grossreceipt = parseFloat(frm.grossreceiptsupply.value);
    if(grossreceipt < .02)
    { alert("gross receipt supply cannot be lower than .02")
      setTimeout(function(){frm.grossreceiptsupply.focus();}, 1);
    }
    if(grossreceipt > .06)
    { alert("gross receipt supply cannot be higher than .06")
      setTimeout(function(){frm.grossreceiptsupply.focus();}, 1);
    }
	
}

function calcUnitCostKw()
{	var frm = document.forms[0];
    var kwtax = (frm.TDwithtax.checked ? parseFloat(frm.TDsalesamt.value) : 0) 
    if(isNaN(kwtax)) kwtax=0;
	frm.unitcostkw.value = roundNumber((parseFloat(frm.costkw.value)+kwtax)/parseFloat(frm.totalkw.value),6)
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
	frm.averagecost.value = roundNumber(parseFloat(frm.totalbillamt.value)/parseFloat(frm.totalkwh.value),6);
}

function calcTDunitcostkwh()
{	var frm = document.forms[0];
    var kwhtax = (frm.TDwithtax.checked ? parseFloat(frm.TDkwhsalesamt.value) : 0) 
    if(isNaN(kwhtax)) kwhtax=0;
	frm.TDunitcostkwh.value = roundNumber((parseFloat(frm.costkwh.value)+kwhtax)/parseFloat(frm.totalkwh.value),6);
}

function calcTDunitcostkw()
{	var frm = document.forms[0];
    var kwtax = (frm.TDwithtax.checked ? parseFloat(frm.TDsalesamt.value) : 0) 
    if(isNaN(kwtax)) kwtax=0;
	frm.TDunitcostkw.value = roundNumber((parseFloat(frm.costkw.value)+kwtax)/parseFloat(frm.totalkw.value),6);
}

var noCalcTDTotalAmt = 0;
function calcTDTotalAmt()
{	var frm = document.forms[0];
	if(noCalcTDTotalAmt)
	{	if(!confirm("Recalculate Total Total T&D?")){return;}
		else{frm.TDtotalamt.focus();}
	}
	if(frm.TDwithmac!=undefined)
	{	var mac = (frm.TDwithmac.checked ? parseFloat(frm.MACdollar.value) : 0)
		if(isNaN(mac)) mac=0;
	}else
	{	var mac = (frm.TDwithfuel.checked ? parseFloat(frm.fueladjdollar.value) : 0)
		if(isNaN(mac)) mac=0;
	}
    var kwtax = (frm.TDwithtax.checked ? parseFloat(frm.TDsalesamt.value) : 0) 
    var kwhtax = (frm.TDwithtax.checked ? parseFloat(frm.TDkwhsalesamt.value) : 0) 
	if(isNaN(kwtax)) kwtax=0;
    if(isNaN(kwhtax)) kwhtax=0;
	var ckw = parseFloat(frm.costkw.value);
	if(isNaN(ckw)) ckw=0;
	var ckwh = parseFloat(frm.costkwh.value);
	if(isNaN(ckwh)) ckwh=0;
	if(frm.TDtotalcalc[0].checked) frm.TDtotalamt.value = ckw+ckwh+mac+kwtax+kwhtax;
	noCalcTDTotalAmt = 0;
	calcTotalBillAmt();
    calcTDunitcostkwh();
    calcTDunitcostkw();
    calcUnitCostKwh();
    calcUnitCostKw();
    calcCostKWHTotal();
    calcCostKWTotal();
}

function calcCOTotalcommodity()
{	var frm = document.forms[0];
	var subtotal = parseFloat(frm.pretaxcost.value)
	if(isNaN(subtotal)) subtotal=0;
	var grt = parseFloat(frm.grossreceipt.value)
	if(isNaN(grt)) grt=0;
	var tax = (frm.COwithtax.checked ? parseFloat(frm.EscoSTdollar.value) : 0)
	if(isNaN(tax)) tax=0;
	frm.totalcommodity.value = subtotal+grt+tax
	calcCostKWHTotal();
	calcTotalBillAmt();
	calcUnitCostKwh();
	calcCOunitcostkwh();
    calcCostKWHTotal();
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
			//frm.grossrecieptSub.value=frm.TDgrtamt.value;
		}
	}
}







function calcCOunitcostkwh()
{	var frm = document.forms[0];
	frm.COunitcostkwh.value = roundNumber(parseFloat(frm.totalcommodity.value)/parseFloat(frm.totalkwhcom.value),6);
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
	frm.unitcredit.value = roundNumber(parseFloat(frm.lmep.value)/parseFloat(frm.totalkwh.value),6);
}

function calcpretaxcost()
{	var frm = document.forms[0];
	frm.pretaxcost.value = roundNumber(parseFloat(frm.fixedrate.value) * parseFloat(frm.totalkwhcom.value),2);
	calcCostKWHTotal()
}

function calcfixedrate()
{	var frm = document.forms[0];
	frm.fixedrate.value = roundNumber(parseFloat(frm.pretaxcost.value) / parseFloat(frm.totalkwhcom.value),6);
}

function calcCostKWHTotal()
{	var frm = document.forms[0];
	var comod = parseFloat(frm.totalcommodity.value);
	if(isNaN(comod)) comod=0;
    var kwhtax = (frm.TDwithtax.checked ? parseFloat(frm.TDkwhsalesamt.value) : 0) 
	if(isNaN(kwhtax)) kwhtax=0;
	frm.costkwhtotal.value = comod + parseFloat(frm.costkwh.value) + kwhtax;
}

function calcCostKWTotal()
{	var frm = document.forms[0];
	var kwtax = (frm.TDwithtax.checked ? parseFloat(frm.TDsalesamt.value) : 0) 
	if(isNaN(kwtax)) kwtax=0;
	frm.totalkwcost.value = parseFloat(frm.costkw.value) + kwtax;
}

/////////////////////////end of calc/////////////////////////////
function checkTaxBox(cbox)
{	var frm = document.forms[0];
	if(cbox.checked)
	{	frm.TDwithtax.checked = 1
		frm.COwithtax.checked = 1
		//frm.TotalIncludeTax.checked = 1
	}else
	{	frm.TDwithtax.checked = 0
		frm.COwithtax.checked = 0
		//frm.TotalIncludeTax.checked = 0
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

function validateDecimal() {
    var frm = document.forms[0];
    var SC4R2Sub = parseFloat(frm.SC4R2Sub.value);
    var SC4R1Sub = parseFloat(frm.SC4R1Sub.value);
    var SC9R1Sub = parseFloat(frm.SC9R1Sub.value);
    var SC2R1Sub = parseFloat(frm.SC2R1Sub.value);
    var SC9RA1Sub = parseFloat(frm.SC9RA1Sub.value);
    var SC9RA2Sub = parseFloat(frm.SC9RA2Sub.value);
    var SC9HTR1Sub = parseFloat(frm.SC9HTR1Sub.value);
    var SC4RA1Sub = parseFloat(frm.SC4RA1Sub.value);
    var SC4RA2Sub = parseFloat(frm.SC4RA2Sub.value);
    var RiderMSub = parseFloat(frm.RiderMSub.value);
    
    
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC4R1Sub))
		{
			
			frm.SC4R1Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC4R1Sub.style.backgroundColor='#FFCCCC';
		}
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC4R2Sub))
		{
			
			frm.SC4R2Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC4R2Sub.style.backgroundColor='#FFCCCC';
		}
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC9R1Sub))
		{
			
			frm.SC9R1Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC9R1Sub.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC2R1Sub))
		{
			
			frm.SC2R1Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC2R1Sub.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC9RA1Sub))
		{
			
			frm.SC9RA1Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC9RA1Sub.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC9RA2Sub))
		{
			
			frm.SC9RA2Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC9RA2Sub.style.backgroundColor='#FFCCCC';
		}
		
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC9HTR1Sub))
		{
			
			frm.SC9HTR1Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC9HTR1Sub.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC4RA1Sub))
		{
			
			frm.SC4RA1Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC4RA1Sub.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC4RA2Sub))
		{
			
			frm.SC4RA2Sub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC4RA2Sub.style.backgroundColor='#FFCCCC';
		}

	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(RiderMSub))
		{
			
			frm.RiderMSub.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.RiderMSub.style.backgroundColor='#FFCCCC';
		}		
	
	
    
}

function validateDecimalOff() {
    var frm = document.forms[0];
    
    
    var SC4R2SubOFF = parseFloat(frm.SC4R2SubOFF.value);
    var SC4R1SubOFF = parseFloat(frm.SC4R1SubOFF.value);
    var SC9R1SubOFF = parseFloat(frm.SC9R1SubOFF.value);
    var SC2R1SubOFF = parseFloat(frm.SC2R1SubOFF.value);
    var SC9RA1SubOFF = parseFloat(frm.SC9RA1SubOFF.value);
    var SC9RA2SubOFF = parseFloat(frm.SC9RA2SubOFF.value);
    var SC9HTR1SubOFF = parseFloat(frm.SC9HTR1SubOFF.value);
    var SC4RA1SubOFF = parseFloat(frm.SC4RA1SubOFF.value);
    var SC4RA2SubOFF = parseFloat(frm.SC4RA2SubOFF.value);
    
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC4R1SubOFF))
		{
			
			frm.SC4R1SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC4R1SubOFF.style.backgroundColor='#FFCCCC';
		}
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC4R2SubOFF))
		{
			
			frm.SC4R2SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC4R2SubOFF.style.backgroundColor='#FFCCCC';
		}
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC9R1SubOFF))
		{
			
			frm.SC9R1SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC9R1SubOFF.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC2R1SubOFF))
		{
			
			frm.SC2R1SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC2R1SubOFF.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC9RA1SubOFF))
		{
			
			frm.SC9RA1SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC9RA1SubOFF.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC9RA2SubOFF))
		{
			
			frm.SC9RA2SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC9RA2SubOFF.style.backgroundColor='#FFCCCC';
		}
		
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC9HTR1SubOFF))
		{
			
			frm.SC9HTR1SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC9HTR1SubOFF.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC4RA1SubOFF))
		{
			
			frm.SC4RA1SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC4RA1SubOFF.style.backgroundColor='#FFCCCC';
		}
		
	if(/^[-+]?\d{0,9}(\.\d{0,6})?$/.test(SC4RA2SubOFF))
		{
			
			frm.SC4RA2SubOFF.style.backgroundColor='#FFFFFF';
	    }
	else
		{
			
			frm.SC4RA2SubOFF.style.backgroundColor='#FFCCCC';
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
	document.all['tabCap'].style.backgroundColor='#CCCCCC';
	document.all['cap'].style.visibility='hidden';
	document.all['cap'].style.position='absolute';
<%if trim(Escoref)<>"0" then%>
	document.all['tabCom'].style.backgroundColor='#CCCCCC';
	document.all['commodity'].style.visibility='hidden';
	document.all['commodity'].style.position='absolute';
<%end if
  if allowGroups("Genergy_Corp,gReadingandBilling")  then%>
	document.all['tabSub'].style.backgroundColor='#CCCCCC';
	document.all['submeter'].style.visibility='hidden';
	document.all['submeter'].style.position='absolute';
<%end if%>
	document.all[hottab].style.backgroundColor='#6699cc';
	document.all[showdiv].style.visibility='visible';
	document.all[showdiv].style.position='relative';
}
</script>
</head>
<body bgcolor="#eeeeee" onload="checkvalues();">
<%
dim id, fueladj ,grossreceipt ,salestax ,offpeakkwh ,onpeakkwh ,totalkwh ,costkwh ,totalkw ,costkw, grossreceiptsupply
dim fueladjustmentSub ,saletaxSub ',grossrecieptSub 
dim escoFixedRate ,EscoPreTax ,EscoST ,EscoGR ,EscoBillAmt ,EscoAcct 
dim unitcostkwh ,unitcostkw ,avgCost ,totalbillamt, totalkwcost, Estimated,count, MSCDetailList, MSCDeatailOffList
dim TDtotalamt, TDwithtax, COwithtax, TDwithmac, TDunitcostkwh, TDsalestax, TDunitcostkw, TDgrtpercent, TDgrtamt, COunitcostkwh, MSC, COgrtpercent, MAC, TotalIncludeTax, TDcostKWH
dim lmep, unitcredit, TDsalesamt, TDkwhsalesamt,MACdollar, EscoSTdollar, fueladjdollar, TDwithfuel, TDtotalcalc, hasdata, note, MSCDetail, MSCDetailOFF
dim MSCSC4R1, MSCSC4R2, MSCSC9R1, MSCSC2R1, MSCSC9RA1, MSCSC9RA2, MSCSC9HTR1,MSCSC4RA1, MSCSC4RA2,MSCRiderM
dim MSCSC4R1OFF, MSCSC4R2OFF, MSCSC9R1OFF, MSCSC2R1OFF, MSCSC9RA1OFF, MSCSC9RA2OFF, MSCSC9HTR1OFF,MSCSC4RA1OFF, MSCSC4RA2OFF
dim search1, search2, search3,search4,search5,search6,search7,search8, search9, search10, search11,search12,search13,search14,search15,search16,search17,search18,search19,search20
dim search1OFF, search2OFF, search3OFF,search4OFF,search5OFF,search6OFF,search7OFF,search8OFF, search9OFF, search10OFF, search11OFF,search12OFF,search13OFF,search14OFF,search15OFF,search16OFF,search17OFF,search18OFF
dim hold1, hold2, hold3, hold4, hold5, hold6, hold7, hold8, hold9, length, diff1, diff2, diff3, diff4, diff5, diff6, diff7, diff8, diff9
dim hold1OFF, hold2OFF, hold3OFF, hold4OFF, hold5OFF, hold6OFF, hold7OFF, hold8OFF, lengthOFF, diff1OFF, diff2OFF, diff3OFF, diff4OFF, diff5OFF, diff6OFF, diff7OFF, diff8OFF
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
    totalkwcost = rst1("totalkwcost")
	
	MSCDetailOFF = rst1("MSCDetailOffPk")
	
	MSCDetail = rst1("MSCDetail")
		
	length = Len(MSCDetail)
	
	lengthOFF = Len(MSCDetailOFF)


	
	if length > 0 then
	
		mscdetaillist = Split( MSCDetail, "|" )
		if IsArray(MSCDetailList) then
			count = ubound(MSCDetailList)
		end if

    	
	    MSCSC4R1   = Split(MSCDetailList(0),"=")(1)
 
        MSCSC4R2   = Split(MSCDetailList(1),"=")(1)

        MSCSC9R1   = Split(MSCDetailList(2),"=")(1)
response.write MSCSC9R1 & "</br>"
        MSCSC2R1   = Split(MSCDetailList(3),"=")(1)

	    MSCSC9RA1  = Split(MSCDetailList(4),"=")(1)

	    MSCSC9RA2  = Split(MSCDetailList(5),"=")(1)

	    MSCSC9HTR1 = Split(MSCDetailList(6),"=")(1)

	    MSCSC4RA1  = Split(MSCDetailList(7),"=")(1)

	    MSCSC4RA2  = Split(MSCDetailList(8),"=")(1)
		if count > 9 then
			MSCRiderM  = Split(MSCDetailList(9),"=")(1)
		end if
    end if
    
    
    if lengthOFF > 0 then
	
	    search1OFF = InStr(MSCDetailOFF,"=")
	    search2OFF = InStr(MSCDetailOFF,"|")
    	
	    diff1OFF = search2OFF - search1OFF
    	
	    MSCSC4R1OFF = Mid(MSCDetailOFF,search1OFF + 1,diff1OFF - 1)

        hold1OFF = Mid(MSCDetailOFF, search2OFF + 1, lengthOFF - search2OFF)
        
        search3OFF = InStr(hold1OFF,"=")
	    search4OFF = InStr(hold1OFF,"|")
    	
	    diff2OFF = search4OFF - search3OFF
        
        
        MSCSC4R2OFF = Mid(hold1OFF,search3OFF + 1,diff2OFF - 1)
        
        hold2OFF = Mid(hold1OFF, search4OFF + 1, lengthOFF - search4OFF)
        
        search5OFF = InStr(hold2OFF,"=")
	    search6OFF = InStr(hold2OFF,"|")
    	
	    diff3OFF = search6OFF - search5OFF
        
        MSCSC9R1OFF = Mid(hold2OFF,search5OFF + 1,diff3OFF - 1)
        
        hold3OFF = Mid(hold2OFF, search6OFF + 1, lengthOFF - search6OFF)
        
        search7OFF = InStr(hold3OFF,"=")
	    search8OFF = InStr(hold3OFF,"|")
    	
	    diff3OFF = search8OFF - search7OFF
        
        MSCSC2R1OFF = Mid(hold3OFF,search7OFF + 1,diff3OFF - 1)
        
        hold4OFF = Mid(hold3OFF, search8OFF + 1, lengthOFF - search8OFF)
        
        search9OFF = InStr(hold4OFF,"=")
	    search10OFF = InStr(hold4OFF,"|")
    	
	    diff4OFF = search10OFF - search9OFF
	    
	    MSCSC9RA1OFF = Mid(hold4OFF,search9OFF + 1,diff4OFF - 1)
	    
	    hold5OFF = Mid(hold4OFF, search10OFF + 1, lengthOFF - search10OFF)
        
        search11OFF = InStr(hold5OFF,"=")
	    search12OFF = InStr(hold5OFF,"|")
    	
	    diff5OFF = search12OFF - search11OFF
	    
	    MSCSC9RA2OFF = Mid(hold5OFF,search11OFF + 1,diff5OFF - 1)
	    
	    hold6OFF = Mid(hold5OFF, search12OFF + 1, lengthOFF - search12OFF)
        
        search13OFF = InStr(hold6OFF,"=")
	    search14OFF = InStr(hold6OFF,"|")
    	
	    diff6OFF = search14OFF - search13OFF
	    
	    MSCSC9HTR1OFF = Mid(hold6OFF,search13OFF + 1,diff6OFF - 1)
	    
	    hold7OFF = Mid(hold6OFF, search14OFF + 1, lengthOFF - search14OFF)
        
        search15OFF = InStr(hold7OFF,"=")
	    search16OFF = InStr(hold7OFF,"|")
    	
	    diff7OFF = search16OFF - search15OFF
	    
	    MSCSC4RA1OFF = Mid(hold7OFF,search15OFF + 1,diff7OFF - 1)
	    
	    hold8OFF = Mid(hold7OFF, search16OFF + 1, lengthOFF - search16OFF)
        
        search17OFF = InStr(hold8OFF,"=")
	    search18OFF = InStr(hold8OFF,"|")
    	
	    diff8OFF = search18OFF - search17OFF
	    
	    MSCSC4RA2OFF = Mid(hold8OFF,search17OFF + 1,diff8OFF - 1)
	    
    end if

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
	COwithtax = rst1("COwithtax")
	TDwithmac = rst1("TDwithmac")
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

    Estimated = rst1("Estimated")

    TDkwhsalesamt = rst1("TDkwhsalesamt")

    grossreceiptsupply = rst1("grossreceiptsupply")
	MACdollar = rst1("MACdollar")
	EscoSTdollar = rst1("EscoSTdollar")
	fueladjdollar = rst1("fueladjdollar")
	TDwithfuel = rst1("TDwithfuel")
  TDtotalcalc = rst1("TDtotalcalc")
	note = rst1("note")
	timestamp = rst1("entry_timestamp")
  rst1.close
  hasdata = true
%>
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#6699cc"> 
      <div align="left"> <span class="standardheader"><font size="2">Utility Bill for Period <%=Request.querystring("bp")%> Year <%=Request.querystring("by")%> (Timestamped: <%=timestamp%>)</font></span></div>
    </td>
  </tr>
</table>
<%else
  rst1.close
hasdata = false
%>
<body bgcolor="#eeeeee">
<table width="100%" border="0">
<td bgcolor="#6699cc"><span class="standardheader"><font size="2">Utility Bill for Period <%=request("bp")%>, Year <%=request("by")%></font></span></td>
</table>
<%
end if
%>
<form name="detail" method="post" action="updutilitybill.asp" onsubmit="zeroNaNs(this)">
<table border="0" cellspacing="1" cellpadding="0">
<tr style="font-family:Arial, Helvetica, sans-serif; font-size:13">
	<td id="tabDel" style="background-color:#6699cc">&nbsp;<b><a href="javascript:changeTab('tabDel','delivery');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">T&nbsp;&amp;&nbsp;D</a></b>&nbsp;</td>

	<td id="tabCap" style="background-color:#CCCCCC">&nbsp;<b><a href="javascript:changeTab('tabCap','cap');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">MSC</a></b>&nbsp;</td>
<%if Escoref<>"0" then%>
	<td id="tabCom" style="background-color:#CCCCCC">&nbsp;<b><a href="javascript:changeTab('tabCom','commodity');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">Commodity&nbsp;Charge</a></b>&nbsp;</td>
<%end if
if allowGroups("Genergy_Corp,gReadingandBilling,IT Services")  then%>
	<td id="tabSub" style="background-color:#CCCCCC">&nbsp;<b><a href="javascript:changeTab('tabSub','submeter');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">Submetered</a></b>&nbsp;</td>
<%end if%>
	<td id="tabTot" style="background-color:#CCCCCC">&nbsp;<b><a href="javascript:changeTab('tabTot','totals');" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">Totals</a></b>&nbsp;</td>
    <td>****FILL IN T&D TAB VALUES BEFORE ALL OTHER TABS****</td>
</tr>
</table>
    <div id="delivery" style="border: 2px solid #6699cc;height:260;position:relative;visibility:visible"> 
    <table width="100%" border="0" style="font-family:Arial, Helvetica, sans-serif; color:black">
      <tr> 
        <td width="15%"> <input type="hidden" name="id1" value="<%=id%>"> <input type="hidden" name="by" value="<%=Request.querystring("by")%>"> 
          <input type="hidden" name="bp" value="<%=Request.querystring("bp")%>"> 
          <input type="hidden" name="ypid" value="<%=ypid%>"> <input type="hidden" name="acctid" value="<%=acctid%>"> 
          <input type="hidden" name="escoref" value="<%=escoref%>"> <input type="hidden" name="utility" value="<%=utility%>"> 
          <u>KWH</u></td>
        <td width="1%"></td>
        <td width="24%"></td>
        <td width="15%"><u>KW</u></td>
        <td width="1%"></td>
        <td width="24%">&nbsp;</td>
        <td width="15%"></td>
        <td width="24%">&nbsp;</td>
      </tr>
      <tr> 
        <td>On Peak KWH</td>
        <td></td>
        <td> <input type="text" name="onpeakkwh" value="<%=onpeakkwh%>" onKeyUp="calcTotalKwh();checkvalues();"> 
        </td>
        <td>Total KW</td>
        <td></td>
        <td> <input type="tkw" name="totalkw" value="<%=totalkw%>" onKeyUp="calcUnitCostKw();calcTDunitcostkw();checkvalues();"> 
        </td>
      </tr>
      <tr> 
        <td>Off Peak KWH</td>
        <td></td>
        <td> <input type="text" name="offpeakkwh" value="<%=offpeakkwh%>" onKeyUp="calcTotalKwh();checkvalues();"> 
        </td>
        <td>Total KW Cost</td>
        <td>$</td>
        <td> <input type="text" name="costkw" value="<%=costkw%>" onKeyUp="calcTDTotalAmt();calcTDgrt(i);calcUnitCostKw();calcTDunitcostkw();checkvalues();"> 
          &nbsp;<a href="#" onclick="window.open('kw_kwh_calc.asp?calc=costkw','calc','width=180,height=180');" style="font-size:10;px">calc</a></td>
      </tr>
      <tr> 
        <td>Total KWH</td>
        <td></td>
        <td> <input type="text" name="totalkwh" value="<%=totalkwh%>" onChange="noCalcTotalKwh=1;" onKeyUp="calcUnitCredit();totalkwhcom.value=this.value;calcUnitCostKwh();calcAverageCost();calcTDunitcostkwh();checkvalues();"> 
        </td>
        <td>Fuel Adj. Factor</td>
        <td>MAC&nbsp;$</td>
        <td> 
         <input type="text" name="MAC" value="<%=MAC%>" onKeyUp="calcfueladjSub();checkvalues();" size="5"> 
         MSC&nbsp;$<input type="text" name="MSC" value="<%=MSC%>" onKeyUp="calcfueladjSub();checkvalues();"  size="5">
         &nbsp;2&nbsp;ADJ&nbsp; <input type="text" name="fueladjustmentSub" value="<%=formatnumber(fueladjustmentSub,6)%>" readonly="true" size="5">(MAC + MSC)
          
          
        </td>
      </tr>
      <tr> 
        <td>Cost KWH</td>
        <td>$</td>
        <td> <input type="text" name="costkwh" value="<%=TDcostkwh%>" onKeyUp="calcCostKWHTotal();calcTDTotalAmt();calcTDgrt(i);calcUnitCostKwh();calcTDunitcostkwh();checkvalues();"> 
          &nbsp;<a href="#" onclick="window.open('kw_kwh_calc.asp?calc=costkwh','calc','width=180,height=150');" style="font-size:10;px">calc</a></td>
        <td>Sales&nbsp;Tax&nbsp; <input type="checkbox" name="TDwithtax" value="1" onclick="calcTDTotalAmt();"<%if trim(TDwithtax)="1" then response.write " CHECKED"%>> 
          <span style="font-size:10">&nbsp;include</span></td>
        <td>$(KWH)</td>
        <td> <input type="text" name="TDkwhsalesamt" value="<%=TDkwhsalesamt%>" onKeyUp="calcTDTotalAmt();checkvalues();" size="5"> 
            &nbsp;$(KW)&nbsp; 
            <input type="text" name="TDsalesamt" value="<%=TDsalesamt%>" onKeyUp="calcTDTotalAmt();checkvalues();" size="5"> 
            &nbsp;Raw&nbsp;%&nbsp; <input type="text" name="TDsalestax" value="<%=TDsalestax%>" onKeyUp="checkvalues();" size="5"> 
        </td>
      </tr>
      <tr> 
        <td><nobr>T&nbsp;&amp;&nbsp;D Unit Cost KWH</nobr></td>
        <td>$</td>
        <td> <input type="text" name="TDunitcostkwh" value="<%=TDunitcostkwh%>" onKeyUp="checkvalues();"> 
        </td>
        <td><nobr>T&nbsp;&amp;&nbsp;D Total Bill   ( Delivery )</nobr></td>
        <td>$</td>
        <td> <input type="text" name="TDtotalamt" value="<%=TDtotalamt%>" onChange="noCalcTDTotalAmt=1;" onKeyUp="calcTDgrt(i);calcTotalBillAmt();checkvalues();"> 
        </td>
      </tr>
      <tr> 
        <td>T&nbsp;&amp;&nbsp;D Unit Cost KW</td>
        <td>$</td>
        <td> <input type="text" name="TDunitcostkw" value="<%=TDunitcostkw%>" onKeyUp="checkvalues();"> 
        </td>
        <td>Calculate Total</td>
        <td></td>
        <td> <input name="TDtotalcalc" value="1" onclick="calcTDTotalAmt()" type="radio"<%if trim(TDtotalcalc)="True" or trim(TDtotalcalc)="" then response.write " CHECKED"%>>
          Yes&nbsp; <input name="TDtotalcalc" value="0" onclick="calcTDTotalAmt();" type="radio"<%if trim(TDtotalcalc)="False" then response.write " CHECKED"%>>
          No</td>
      </tr>
      <tr> 
        <td>LMEP</td>
        <td>$</td>
        <td> <input type="text" name="lmep" value="<%=lmep%>" onKeyUp="calcUnitCredit();checkvalues();"> 
        </td>
        <td>Estimated<input type="checkbox" name="Estimated" value="1" <%if trim(Estimated)="1" then response.write " CHECKED"%>> </td>
        <td> <input type="hidden" name="TDgrtamt" value="<%=TDgrtamt%>" onKeyUp="calcTDgrt(this);grossrecieptSub.value=this.value;checkvalues();"> 
          <!-- &nbsp;Raw&nbsp;% -->
          <input type="hidden" name="TDgrtpercent" value="<%=TDgrtpercent%>" onKeyUp="calcTDgrt(this);checkvalues();" size="10">
          
        </td>
      </tr>
      <tr> 
        <td>Unit Credit</td>
        <td>$</td>
        <td> <input type="text" name="unitcredit" value="<%=unitcredit%>" readonly> 
        </td>
        <td></td>
        <td></td>
        <td></td>
        <td> 
          <input type="hidden" name="TDwithmac" value="0" > 
          <input type="hidden" name="TDwithfuel" value="0"> 
          <input type="hidden" name="MACdollar" value="<%=MACdollar%>"> 
          <input type="hidden" name="fueladjdollar" value="<%=fueladjdollar%>"> 
          
        </td>
      </tr>
    </table>
  </div>
<div id="commodity" style="border: 2px solid #6699cc;height:260;position:absolute;visibility:hidden">
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
<td><input type="text" name="totalkwhcom" value="<%=totalkwh%>" onChange="noCalcTotalKwh=1;" onKeyUp="" readonly></td>
<td>Sub Total</td><td>$</td>
<td><input type="text" name="pretaxcost" value="<%=EscoPreTax%>" onKeyUp="calcfixedrate();calcCOTotalcommodity();checkvalues();"></td>
<td></td>
<td></td>
</tr>

<tr>
<td>Fixed Rate</td><td>$</td>
<td><input type="text" name="fixedrate" value="<%=escoFixedRate%>" onKeyUp="calcpretaxcost();calcCOTotalcommodity();checkvalues();"></td>
<td>Gross Receipt</td><td>$</td>
<td><input type="text" name="grossreceipt" value="<%=EscoGR%>" onKeyUp="calcCOTotalcommodity();checkvalues();">&nbsp;Raw&nbsp;%&nbsp;<input type="text" name="COgrtpercent" value="<%=COgrtpercent%>" onKeyUp="checkvalues();" size="10"></td>
</tr>

<tr>
<td><nobr>Comm. Unit Cost KWH</nobr></td><td>$</td>
<td><input type="text" name="COunitcostkwh" value="<%=COunitcostkwh%>" onKeyUp="checkvalues();" readonly></td>
<td>Sales&nbsp;Tax&nbsp;<input type="checkbox" name="COwithtax" value="1" onclick="calcCOTotalcommodity();"<%if trim(COwithtax)="1" then response.write " CHECKED"%>><span style="font-size:10">&nbsp;include</td><td>$</td>
<td><input type="text" name="EscoSTdollar" value="<%=EscoSTdollar%>" onKeyUp="calcCOTotalcommodity();checkvalues();" size="10">&nbsp;Raw&nbsp;%&nbsp;<input type="text" name="EscoST" value="<%=EscoST%>" onKeyUp="checkvalues();"></td>
</tr>

<tr>
<td></td><td></td>
<td></td>
<td><nobr>Total commodity    ( Supply )</nobr></td><td>$</td>
<td><input type="text" name="totalcommodity" value="<%=EscoBillAmt%>" onKeyUp="checkvalues();" readonly></td>
</tr>

</table>
</div>
<div id="cap" style="border: 2px solid #6699cc;height:260;position:absolute;visibility:hidden">
<table border="0" style="font-family:Arial, Helvetica, sans-serif; color:black">
<tr><td><u>MSC</u></td><td></td><td><u>OFF-PEAK</u></td></tr>

<tr><td width="15%">SC4R1</td>
	<td width="24%"><input type="text" name="SC4R1Sub" value="<%=MSCSC4R1%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC4R1SubOFF" value="<%=MSCSC4R1OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">SC4R2-SC9R2</td>
	<td width="24%"><input type="text" name="SC4R2Sub" value="<%=MSCSC4R2%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC4R2SubOFF" value="<%=MSCSC4R2OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">SC9R1</td>
	<td width="24%"><input type="text" name="SC9R1Sub" value="<%=MSCSC9R1%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC9R1SubOFF" value="<%=MSCSC9R1OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">SC2R1</td>
	<td width="24%"><input type="text" name="SC2R1Sub" value="<%=MSCSC2R1%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC2R1SubOFF" value="<%=MSCSC2R1OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">SC9RA1</td>
	<td width="24%"><input type="text" name="SC9RA1Sub" value="<%=MSCSC9RA1%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC9RA1SubOFF" value="<%=MSCSC9RA1OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">SC9RA2</td>
	<td width="24%"><input type="text" name="SC9RA2Sub" value="<%=MSCSC9RA2%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC9RA2SubOFF" value="<%=MSCSC9RA2OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">SC9HTR1</td>
	<td width="24%"><input type="text" name="SC9HTR1Sub" value="<%=MSCSC9HTR1%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC9HTR1SubOFF" value="<%=MSCSC9HTR1OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">SC4RA1</td>
	<td width="24%"><input type="text" name="SC4RA1Sub" value="<%=MSCSC4RA1%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC4RA1SubOFF" value="<%=MSCSC4RA1OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">SC4RA2</td>
	<td width="24%"><input type="text" name="SC4RA2Sub" value="<%=MSCSC4RA2%>" onKeyUp="validateDecimal();"></td>
	<td width="24%"><input type="text" name="SC4RA2SubOFF" value="<%=MSCSC4RA2OFF%>" onKeyUp="validateDecimalOff();"></td></tr>
<tr><td width="15%">Rider M ONLY</td>
	<td width="24%"><input type="text" name="RiderMSub" value="<%=MSCRiderM%>" onKeyUp="validateDecimal();">** READ - Utility Bill pg1. : Message Center : Adjustment Information</td>
	<td width="24%">&nbsp;</td></tr>	
</table>
</div>
<div id="submeter" style="border: 2px solid #6699cc;height:260;position:absolute;visibility:hidden">
<table border="0" style="font-family:Arial, Helvetica, sans-serif; color:black">
<tr><td><u>Submeter</u></td></tr>
<tr><td width="15%">Sales&nbsp;Tax</td><td width="1%">%</td>
	<td width="24%"><input type="text" name="saletaxSub" value="<%=saletaxSub%>" onKeyUp="checkvalues();" onblur="checksalestax();"></td></tr>
<tr><td width="15%">Gross&nbsp;Receipt&nbsp;Delivery</td><td width="1%">%</td>
	<td width="24%"><input type="text" name="grossrecieptSub" value="<%=TDgrtamt%>" onKeyUp="checkvalues();" onblur="checkgrossreceipt();"></td></tr>
<tr><td width="15%">Gross&nbsp;Receipt&nbsp;Supply</td><td width="1%">%</td>
	<td width="24%"><input type="text" name="grossreceiptsupply" value="<%=grossreceiptsupply%>" onKeyUp="checkvalues();" onblur="checkgrossreceiptsupply();"></td></tr>
</table>

<%
rst1.open "SELECT count(distinct FuelAdj) as fueladj FROM utilitybill u, billyrperiod b WHERE b.ypid=u.ypid and b.billyear="&by&" and b.billperiod="&bp&" and BldgNum='"&bldg&"' and totalbillamt <> 0", cnn1
if not rst1.eof then 
  if cint(rst1("fueladj"))>1 then response.write "<span style=""color:red;font-face:arial;"">The fuel adjustments do not match between utility accounts</span>"
end if
rst1.close
%>
</div>
<div id="totals" style="border: 2px solid #6699cc;height:260;position:absolute;visibility:hidden">
<table border="0" style="font-family:Arial, Helvetica, sans-serif; color:black">
<tr><td>Unit Cost KWH</td><td>$</td>
	<td><input type="text" name="unitcostkwh" value="<%=unitcostkwh%>" onKeyUp="checkvalues();"></td>
	<td>Average Cost</td><td>$</td>
	<td><input type="text" name="averagecost" value="<%=avgCost%>" onKeyUp="checkvalues();"></td></tr>
<tr valign="top"><td>Unit Cost KW</td><td>$</td>
	<td><input type="text" name="unitcostkw" value="<%=unitcostkw%>" onKeyUp="checkvalues();"></td>
	<td>Total Bill Amt Cost</td><td>$</td>
	<td><input type="text" name="totalbillamt" value="<%=totalbillamt%>" onKeyUp="checkvalues();"></td>
	<td></td>
</tr>
<tr><td>&nbsp;</td>
<tr><td>Building Totals</td>
<tr><td>Unit Cost KWH</td><td>$</td>
	<td><input type="text" name="" value="<%=formatnumber(granduckwh,6)%>" onKeyUp="checkvalues();"></td>
	<td>Average Cost</td><td>$</td>
	<td><input type="text" name="" value="<%=formatnumber(grandavg,6)%>" onKeyUp="checkvalues();"></td></tr>
<tr valign="top"><td>Unit Cost KW</td><td>$</td>
	<td><input type="text" name="" value="<%=formatnumber(granduckw,6)%>" onKeyUp="checkvalues();"></td>
	<td>Total Bill Amt Cost</td><td>$</td>
	<td><input type="text" name="" value="<%=grandtotal%>" onKeyUp="checkvalues();"></td>
</tr>
</table>
</div>&nbsp;<br><input type="hidden" name="costkwhtotal" value="<%=costkwh%>"><input type="hidden" name="bldg" value="<%=bldg%>"><input type="hidden" name="totalkwcost" value="<%=totalkwcost%>">
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr><td rowspan="2" valign="top">
		<%
		if not(isBuildingOff(bldg)) then
			if hasdata then%>
			  <input type="submit" name="action" value="UPDATE">
			 <input type="button" value="Email R&B" onClick="window.open('RB_email.asp?action=Email&utility=<%=utility%>&building=<%=bldg%>&byear=<%=by%>&bperiod=<%=bp%>&pid=<%=getpid(bldg)%>&utilityid='+this.form.utility.value,'Email','width=200,height=100')">
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
'TK: 04/28/2006 
on error resume next
set rst1 = nothing
set rst2 = nothing	
'#TK: 04/28/2006
set cnn1 = nothing
%>
</body>
</html>