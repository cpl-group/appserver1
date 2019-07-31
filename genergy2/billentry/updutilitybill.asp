<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, id, by, bp, ypid, acctid, offpeakkwh, totalkw, onpeakkwh, costkw, totalkwh, costkwh, totalkwhcom, pretaxcost, fixedrate, salestax, grossreceipt, totalcommodity, unitcostkwh, unitcostkw, averagecost, totalbillamt, action, escoref, utility, fueladjustmentSub, saletaxSub, grossrecieptSub, TDtotalamt, TDwithtax, TDunitcostkwh, TDsalestax, TDunitcostkw, TDgrtpercent, TDgrtamt, COunitcostkwh, MSC, COgrtpercent, MAC, TotalIncludeTax, COwithtax, lmep, unitcredit, TDcostKWH, TDsalesamt, TDkwhsalesamt, MACdollar, EscoSTdollar, EscoST, fueladjdollar, TDwithmac, TDwithfuel, TDtotalcalc, note
dim MSCDetail, MSCDetailOFF, totalkwcost,grossreceiptsupply,Estimated
bldg = request("bldg")

by = Request.form("by")
bp = Request.form("bp")
id=Request.form("ypid")
ypid = Request.form("ypid")
acctid = Request.form("acctid")
'delivery
offpeakkwh = Request.form("offpeakkwh")
totalkw = Request.form("totalkw")
onpeakkwh = Request.form("onpeakkwh")
costkw = Request.form("costkw")
totalkwh = Request.form("totalkwh")
costkwh = Request.form("costkwhtotal")
if not(isnumeric(costkwh)) then costkwh=0
'submeter
fueladjustmentSub = Request.form("fueladjustmentSub")
saletaxSub = Request.form("saletaxSub")
grossrecieptSub = Request.form("grossrecieptSub")
grossreceiptsupply = Request.form("grossreceiptsupply")
'totals
unitcostkwh = Request.form("unitcostkwh")
unitcostkw = Request.form("unitcostkw")
averagecost = Request.form("averagecost")
totalbillamt = Request.form("totalbillamt")
'added
TDtotalamt = Request.form("TDtotalamt")
TDwithtax = Request.form("TDwithtax")
Estimated = Request.form("Estimated")

TDunitcostkwh = Request.form("TDunitcostkwh")
TDsalestax = Request.form("TDsalestax")
TDunitcostkw = Request.form("TDunitcostkw")
TDgrtpercent = Request.form("TDgrtpercent")
'TDgrtamt = Request.form("TDgrtamt")
MSC = Request.form("MSC")
TotalIncludeTax = Request.form("TotalIncludeTax")
lmep = Request.form("lmep")
unitcredit = Request.form("unitcredit")
'TDtotalamt, TDwithtax, TDunitcostkwh, TDsalestax, TDunitcostkw, TDgrtpercent, TDgrtamt, COunitcostkwh, MSC, COgrtpercent, MAC, TotalIncludeTax
TDcostKWH = Request("costKWH")
TDwithfuel = Request("TDwithfuel")
TDtotalcalc = Request("TDtotalcalc")
totalkwcost = Request("totalkwcost")
if not(isnumeric(totalkwcost)) then totalkwcost=0

MSCDetail = "SC4R1=" + Request.form("SC4R1Sub") + "|SC4R2=" + Request.form("SC4R2Sub") + "|SC9R1=" + Request.form("SC9R1Sub") + "|SC2R1=" + Request.form("SC2R1Sub") + "|SC9RA1=" + Request.form("SC9RA1Sub") + "|SC9RA2=" + Request.form("SC9RA2Sub") + "|SC9HTR1=" + Request.form("SC9HTR1Sub") + "|SC4RA1=" + Request.form("SC4RA1Sub") + "|SC4RA2=" + Request.form("SC4RA2Sub") + "|RiderM=" + Request.form("RiderMSub") + "|"     

MSCDetailOFF = "SC4R1=" + Request.form("SC4R1SubOFF") + "|SC4R2=" + Request.form("SC4R2SubOFF") + "|SC9R1=" + Request.form("SC9R1SubOFF") + "|SC2R1=" + Request.form("SC2R1SubOFF") + "|SC9RA1=" + Request.form("SC9RA1SubOFF") + "|SC9RA2=" + Request.form("SC9RA2SubOFF") + "|SC9HTR1=" + Request.form("SC9HTR1SubOFF") + "|SC4RA1=" + Request.form("SC4RA1SubOFF") + "|SC4RA2=" + Request.form("SC4RA2SubOFF") + "|"   

action = Request.form("action")
escoref = Request.form("escoref")
utility = Request.form("utility")
TDsalesamt = request("TDsalesamt")
TDkwhsalesamt = request("TDkwhsalesamt")
MACdollar = request("MACdollar")
fueladjdollar = request("fueladjdollar")
note = replace(left(request("note"),500),"'","''")


MAC = Request.form("MAC")
if escoref<>"0" then 'has esco so fill esco fields
	pretaxcost = Request.form("pretaxcost")
	fixedrate = Request.form("fixedrate")
	EscoST = Request.form("EscoST")
	grossreceipt = Request.form("grossreceipt")
    totalcommodity = Request.form("totalcommodity")
	COunitcostkwh = Request.form("COunitcostkwh")
	COwithtax = Request.form("COwithtax")
	'MAC = Request.form("MAC")
	COgrtpercent = Request.form("COgrtpercent")
	EscoSTdollar = request("EscoSTdollar")
	TDwithmac = request("TDwithmac")
else 'has no esco so fields should be blank
	pretaxcost = ""
	fixedrate = ""
	salestax = ""
	grossreceipt = ""
    totalcommodity = ""
	COunitcostkwh = ""
	'MAC = ""
	TDwithmac = ""
	EscoST = ""
	COgrtpercent = ""
	COwithtax = ""
	EscoSTdollar = ""
end if

dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getLocalConnect(bldg)
response.write totalcommodity
if action="UPDATE" then
	strsql = "UPDATE utilitybill SET "&_
	"Utility='"& utility &"', TotalBillAmt='"& totalbillamt &"', TotalKWH='"& totalkwh &"', avgkwh='"& averagecost &"', OnPeakKWH='"& onpeakkwh &"', OffPeakKWH='"& offpeakkwh &"', TotalKW='"& totalkw &"', ypId='"& ypid &"', CostKWH='"& costkwh &"', UnitCostKWH='"& unitcostkwh &"', CostKW='"& costkw &"', UnitCostKW='"& unitcostkw &"', AcctID='"& acctid &"', EscoST='"& EscoST &"', EscoGR='"& grossreceipt &"', grossreceiptsupply='"& grossreceiptsupply &"', EscoBillAmt='"& totalcommodity &"', EscoPreTax='"& pretaxcost &"', EscoAcct='"& escoref &"', escoFixedRate='"& fixedrate &"', FuelAdj='"& fueladjustmentSub &"', GrossReceipt='"& grossrecieptSub &"', SalesTax='"& saletaxSub &"',    TDtotalamt='"&TDtotalamt&"', TDwithtax='"&TDwithtax&"', TDunitcostkwh='"&TDunitcostkwh&"', TDsalestax='"&TDsalestax&"', TDunitcostkw='"&TDunitcostkw&"', TDgrtpercent='"&TDgrtpercent&"', TDgrtamt='"&TDgrtamt&"', COunitcostkwh='"&COunitcostkwh&"', MSC='"&MSC&"', COgrtpercent='"&COgrtpercent&"', MAC='"&MAC&"', TotalIncludeTax='"&TotalIncludeTax&"', COwithtax='"&COwithtax&"', lmepcredit='"&lmep&"', unit_credit='"&unitcredit&"', TDcostKWH='"&TDcostKWH&"', TDsalesamt='"&TDsalesamt&"', TDkwhsalesamt='"&TDkwhsalesamt&"',MACdollar='"&MACdollar&"', EscoSTdollar='"&EscoSTdollar&"', fueladjdollar='"&fueladjdollar&"', TDwithmac='"&TDwithmac&"', TDwithfuel='"&TDwithfuel&"', TDtotalcalc='"&TDtotalcalc&"', note='"&note&"', MSCDetail='"&MSCDetail&"', MSCDetailOffPk='"&MSCDetailOFF&"', Estimated='"&Estimated&"' , totalkwcost='"&totalkwcost&"'" &_
	" where ypid="& ypid &" and acctid='"&acctid&"'"
'response.write cnn1 & "<br>"
'response.write strsql
'response.end
else

	strsql = "INSERT utilitybill "&_
	"(Utility, 			TotalBillAmt, 			TotalKWH, 		avgkwh, 			OnPeakKWH, 			OffPeakKWH, 		TotalKW, 		ypId, 			CostKWH, 		UnitCostKWH, 		CostKW, 		UnitCostKW, 			AcctID, 	EscoST, 		EscoGR, 				EscoBillAmt, 			EscoPreTax, 		EscoAcct, 		escoFixedRate, 		FuelAdj,					GrossReceipt,			SalesTax, 			TDtotalamt, 	TDwithtax, 			TDunitcostkwh, 		TDsalestax, 		TDunitcostkw, 		TDgrtpercent, 		TDgrtamt, 		COunitcostkwh, 		MSC, 		COgrtpercent, 		MAC, 		TotalIncludeTax,	COwithtax,		lmepcredit,	  unit_credit,		TDcostKWH,		TDsalesamt,	TDkwhsalesamt,			MACdollar,		EscoSTdollar,		fueladjdollar,			TDwithmac,			TDwithfuel,     TDtotalcalc, note, MSCDetail, MSCDetailOffPk, totalkwcost, grossreceiptsupply,Estimated) VALUES  "&_
	"('"& utility &"', '"& totalbillamt &"', '"& totalkwh &"', '"& averagecost &"', '"& onpeakkwh &"', '"& offpeakkwh &"', '"& totalkw &"', '"& ypid &"', '"& costkwh &"', '"& unitcostkwh &"', '"& costkw &"', '"& unitcostkw &"', '"& acctid &"', '"& salestax &"', '"& grossreceipt &"', '"& totalcommodity &"', '"& pretaxcost &"', '"& escoref &"', '"& fixedrate &"', '"& fueladjustmentSub &"', '"& grossrecieptSub &"', '"& saletaxSub &"', '"&TDtotalamt&"', '"&TDwithtax&"', '"&TDunitcostkwh&"', '"&TDsalestax&"', '"&TDunitcostkw&"', '"&TDgrtpercent&"', '"&TDgrtamt&"', '"&COunitcostkwh&"', '"&MSC&"', '"&COgrtpercent&"', '"&MAC&"', '"&TotalIncludeTax&"', '"&COwithtax&"',	'"&lmep&"', '"&unitcredit&"', '"&TDcostKWH&"', '"&TDsalesamt&"', '"&TDkwhsalesamt&"','"&MACdollar&"', 	'"&EscoSTdollar&"',	'"&fueladjdollar&"',	'"&TDwithmac&"',	'"&TDwithfuel&"', '"&TDtotalcalc&"', '"&note&"', '"&MSCDetail&"', '"&MSCDetailOFF&"', '"&totalkwcost&"', '"&grossreceiptsupply&"', '"&Estimated&"')"

end if
'response.write strsql
'response.end
cnn1.execute strsql
logger(strsql)

set cnn1=nothing
dim tmpMoveFrame
tmpMoveFrame =  "document.location = ""acctdetail.asp?bldg="&bldg&"&acctid="&acctid&"&ypid="&ypid&"&bp="&bp&"&by="&by& """" & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>