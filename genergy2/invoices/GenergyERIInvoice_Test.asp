<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Server.ScriptTimeout = 60*15
dim leaseid, ypid, building, pid, byear, bperiod, utilityid, detailed, meterbreakdown, calcintpeak, SJPproperties, ShowUsageDetails, ShowDemandDetails,onlinereview,totdeliverychrgs, maxmeters, textheader, masterTotal,demo, reject
dim pdfsession,startCount,endCount
leaseid = trim(Request("l"))
ypid = trim(request("y"))
utilityid = trim(request("utilityid"))
building = trim(request("building"))
pid = trim(request("pid"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))
logo = trim(request("logo"))
detailed = trim(request("detailed"))
textheader = trim(request("textheader"))
demo = request("demo")
pdfsession = request("fdp")
startCount = request("s")
endCount = request("e")

if startCount = "" then startCount = 0
if (endCount = "" or endCount = "-1") then endCount = 1000


if demo = "" then demo = false end if 
if trim(request("reject"))="" then reject = 0 else reject = 1

if trim(request("SJPproperties"))="true" then SJPproperties = true else SJPproperties = false
if trim(request("meterbreakdown"))="no" then meterbreakdown = false else meterbreakdown = true
if trim(request("summaryusage"))="true" then showusagedetails= true else showusagedetails = false
if trim(request("summarydemand"))="true" then showdemanddetails = true else showdemanddetails = false

'This Code is Written for Pdf View
if ((request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj"))) OR (pdfsession = "pdffdp"))  then 'this is for pdf sessions
 loadNewXML("activepdf")
loadIps(0)
end if

dim cnn1, rst1, rst2, rst3, bldgrs, usagelabel, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(building)
usagelabel = "used"
dim DBmainIP
DBmainIP = ""

'Getting UtilityId If it is "" 
if utilityid="" and leaseid<>"" then
  rst1.open "SELECT utility FROM tblleasesutilityprices WHERE leaseutilityid="&leaseid, cnn1
  if not rst1.eof then utilityid = cint(rst1("utility"))
  rst1.close
  set rst1 = nothing
end if


dim templid, tempypid, temputility, logo, extusage
set bldgrs = Server.CreateObject("ADODB.Recordset")

if logo = "" then logo = "invoice_logo_1.jpg"

if trim(request("billid"))<>"" and trim(leaseid)<>"" then
	showtenantbill leaseid, trim(request("billid")), utilityid,extusage
elseif leaseid<>"" and ypid<>"" then
	sql = "SELECT Eb.EriBillId as billid FROM tblEriBills Eb WHERE Flag='C' and Reject=0 and Eb.YpId="&ypid&" and Eb.LeaseUtilityId="&leaseid

	'Response.End 
	bldgrs.open sql, cnn1
	if not bldgrs.eof then showtenantbill leaseid, bldgrs("billid"), utilityid,extusage
	bldgrs.close
elseif building<>"" then
	if ypid<>"" then
		bldgrs.open "SELECT Eb.EriBillId as billid, Eb.LeaseUtilityId, Eb.Utility, Eb.Extusg FROM tblEriBills Eb WHERE Flag='B' and Reject=0 and BuildingNumber='"&building&"' and YpId="&ypid&" ORDER BY TenantName", cnn1
		do until bldgrs.eof
			templid = trim(bldgrs("LeaseUtilityId"))
		    temputility = trim(bldgrs("Utility"))
			Extusage = trim(bldgrs("Extusg"))
			showtenantbill templid, bldgrs("billid"), temputility,extusage
			bldgrs.movenext
		loop
	elseif byear<>"" and bperiod<>"" then
    sql = "SELECT Eb.EriBillId as billid, Eb.LeaseUtilityId, YpId, Eb.Utility, Extusg FROM tblEriBills Eb WHERE Flag='B' and Reject=0 and BuildingNumber='"&building&"' and BillYear="&byear&" and BillPeriod="&bperiod
    if isnumeric(utilityid) then sql = sql & " and Utility="&utilityid
    sql = sql & "  ORDER BY TenantName"
		
		bldgrs.open sql, cnn1
		dim count
		dim tempCount
		tempCount = 0
		count = startCount
			if count > 0 then 'skip to the proper bill
			bldgrs.movefirst
			for tempCount = 0 to count-1 Step 1
			bldgrs.movenext
			next
			end if
		do until bldgrs.eof
		'on Error resume next
			templid = trim(bldgrs("LeaseUtilityId"))
			tempypid = trim(bldgrs("YpId"))
			temputility = trim(bldgrs("Utility"))
			Extusage = trim(bldgrs("Extusg"))
			showtenantbill templid, bldgrs("billid"), temputility, extusage
			count = count + 1
				if count >= endCount+1 then 
				response.write "</body></html>"
				response.End()
				end if		
			bldgrs.movenext
		loop
	end if
end if
set cnn1 = nothing


'###begin of showtenantbill , is rest of file ###

sub showtenantbill(leaseid, billid, utilityid, extusage)

dim loopgroup, metercount, tot_onpeak, tot_offpeak, tot_kwhused,tot_kwhusedoff,tot_kwhusedon,tot_kwhusedint, tot_demand_p, tot_demand_c, coincidentflag, usagedivisor, unittype, tot_intpeak, tot_demandoff_p, tot_demandint_p,currentdemandP,currentUsage, serviceclass,serviceclassid, utilityname, paymentterm
Dim ArrDetail1,BaseTotalConsumtion,TotalConsumtion,DateStart,DateEnd,Sqft,BaseTotalAmount
Dim TotalAmount,AdminFee,ServiceFee,SalesTax

'Getting TenantDetails
Dim TenantName,TenantNumber,TenantStreet,TenantCity,TenantState,TenantZip

'Getting Billing Details
Dim BillingBuildingNumber,BillTo,BillingCity,BillingState,BillingZip

usagedivisor = 1
Set rst2 = Server.CreateObject("ADODB.recordset")
'get utility name
rst2.open "SELECT utility FROM tblutility WHERE utilityid="&utilityid, getConnect(pid,building,"billing")
if not rst2.eof then
	utilityname = rst2("utility")
end if
rst2.close
if extusage = "" then 
	sql = "SELECT Extusg FROM tblEriBills WHERE EriBillId = " & billid
	rst2.open sql, cnn1, 2
	
	if not rst2.eof and trim(rst2("Extusg")) <> "" then 
		extusage = rst2("Extusg")
	else 
		extusage = false
	end if 
	rst2.close	
end if

select case utilityid
case 3, 10,6,1,4
	select case utilityid
	case 6,1,4
		usagedivisor = 1
	case else
		usagedivisor = 100
  end select 
	sql = "SELECT Eb.EriDetails,Eb.DateStart,Eb.DateEnd,Eb.Sqft,Eb.TenantName,Eb.TenantNumber,Eb.TenantStreet,Eb.TenantCity,Eb.TenantState,Eb.TenantZip,Eb.BillingTenantBuildingName,Eb.BillingTenantCity,Eb.BillingTenantState,isnull(b.btzip,'') as billzip, b.portfolioid, isnull(b.btstrt,'') as billto,Eb.AddOnFee as myaddonfee, isnull(Eb.EnergyDetail,'0') as energydetail, isnull(Eb.DemandDetail,'0') as demanddetail, Eb.Utility as unittype, isnull(Eb.TotalAmount,0) as Totalamtnonull, isnull(Eb.Tax,0) as taxnonull, isnull(Eb.Energy,0) as energynonull, isnull(Eb.TenantStreet,'') as billingaddress,isnull(Eb.AdminFee,0) as adminfee,isnull(Eb.ServiceFee,0) as servicefeenonull, Eb.FuelAdjustment as fadj, isnull(Eb.Demand,0) as demandnonull, isnull(Eb.Credit,0) as creditnonull, isnull(Eb.Adjustment, 0) as adjustmentnonull, isnull(Eb.Credit,0) as credit, rt.[type] as rt,rt.[id] as rtid, datediff(day, DateStart,DateEnd)+1 as days, Eb.Invoice_Note as invoiceNote, isnull(Eb.Rate_servicefee_dollar,0) as rateservicefee_dollar, Eb.*, l.onlinebill FROM tblEriBills Eb, tblleases l, buildings b, "&DBmainIP&"ratetypes rt WHERE Eb.RateTenant=rt.id AND b.bldgnum=Eb.BuildingNumber and l.billingid = (SELECT billingid FROM tblleasesutilityprices lup WHERE lup.leaseutilityid=Eb.leaseutilityid) and Eb.EriBillId="&billid
	
case else
	if SJPproperties then
  		sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill FROM rpt_Bill_summary_NoBill r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup WHERE r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid
	else
  		sql = "SELECT isnull(b.btzip,'') as billzip, b.portfolioid,	isnull(r.BillingTenantStreet,'') as billto, " & _
  			  "isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, " & _
  			  "r.utility as unittype, isnull(Totalamount,0) as Totalamtnonull, isnull(tax,0) as taxnonull, " & _
			  "isnull(energy,0) as energynonull, isnull(r.tenantstreet,'') as billingaddress, " & _
			  "r.fueladjustment as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, " & _
			  "isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, " & _
			  "rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, " & _
			  "isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, " & _
			  "r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull, " & _
			  "r.addonfee as myaddonfee,  isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, " & _
			  "r.*, l.onlinebill " & _
			  " FROM tblEriBills r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup " & _
			  " WHERE r.ratetenant=rt.id AND b.bldgnum=r.BuildingNumber and lup.leaseutilityid=r.leaseutilityid " & _
			  " and l.billingid=lup.billingid and r.Eribillid="&billid
	end if
end select

rst2.open sql, cnn1, 2
if not rst2.eof then
	pid = rst2("portfolioid")
	unittype = rst2("unittype")
	serviceclass = rst2("rt")
	serviceclassid = rst2("rtid")
	
	TotalConsumtion=rst2("energynonull")			'Added By Rahul getting TotalConsumption
	DateStart = trim(rst2("DateStart"))					'Getting Bill's DateStart
	DateEnd = trim(rst2("DateEnd"))						'Getting Bill's DateEnd
	Sqft=trim(rst2("Sqft"))								'Getting Sqft Area of the Tenant
	TotalAmount=rst2("Totalamtnonull")					'Getting TotalAmount

	AdminFee=rst2("adminfee")							'Getting AdminFee
	ServiceFee=rst2("servicefeenonull")					'Getting Service Fee
	SalesTax=rst2("taxnonull")							'Getting Sales Tax
	ArrDetail2=Split(rst2("EriDetails"),"|")			'Added By RAhul Spliting ERI details
	
	'Getting Tenant Billing Details 
	TenantName=Trim(rst2("TenantName"))
	TenantNumber=Trim(rst2("TenantNumber"))	
	TenantStreet=rst2("TenantStreet")
	TenantCity=rst2("TenantCity")
	TenantState=rst2("TenantState")
	TenantZip=rst2("TenantZip")
	
	'Getting Billing Details
	BillingBuildingNumber=rst2("BillingTenantBuildingName")
	BillTo=rst2("billto")
	BillingCity=rst2("BillingTenantCity")
	BillingState=rst2("BillingTenantState")
	BillingZip=rst2("billzip")
	
	select case serviceclassid 
		case 36
			maxmeters = 3
		case else	
			maxmeters = 9
	end select

	'get payment text
	rst3.open "SELECT isnull(paymentterm,'') as paymentterm FROM portfolio WHERE id="&pid, getConnect(pid,building,"billing")
	if not rst3.eof then
		paymentterm = rst3("paymentterm")
		if trim(paymentterm)="" then 
			paymentterm = "*Payment due upon receipt"
		end if
	end if
	rst3.close
End If
rst2.close

'Code Added by Rahul for getting Details of Normal Calculation
Dim rst4,rst5,ArrDetail2 , sSql
Set rst4 = Server.CreateObject("ADODB.Recordset")
Set rst5 = Server.CreateObject("ADODB.Recordset")
sSql = "SELECT Eb.EriBillId as billid FROM tblEriBills Eb WHERE Flag='B' and Reject=0 and Eb.LeaseUtilityId="&leaseid
rst4.open sSql, cnn1, 2

if not rst4.eof then
	rst5.open "SELECT Energy,TotalAmount,EriDetails FROM tblEriBills WHERE EriBillId="& rst4("billid"), cnn1, 2
	If not rst5.eof then
		BaseTotalConsumtion=rst5("Energy")
		BaseTotalAmount=rst5("TotalAmount")
		ArrDetail1=Split(rst5("EriDetails"),"|")
	End If
	rst5.Close
end if
rst4.close


' Pick up Building KWH and KW  from the UtilityBill
Dim BuildingKWH, BuildingKW
BuildingKWH = 0.0
BuildingKW = 0.0
sSql = "Select TotalKwh, TotalKW from UtilityBill Where ypid = " & ypid
rst4.open sSql, cnn1, 2
if not rst4.eof then
	BuildingKWH = rst4("TotalKwh")
	BuildingKW = rst4("TotalKW")
end if
rst4.Close 

'Making Static ERIF Value as $3.00
Dim ERIF,SubTotal,MonthlyAdjustment,PercentageIncrease,NewSubTotal,TotalDue,BilledAmount,BalanceDue
NewSubTotal = 0.0
TotalDue = 0.0
ERIF=3.00
SubTotal=(ERIF*Sqft)/12
 

IF ((Cdbl(TotalAmount)-Cdbl(BaseTotalAmount))/Cdbl(BaseTotalAmount)) <=0 Then 
	PercentageIncrease= 0
Else
	PercentageIncrease= (Cdbl(TotalAmount)-Cdbl(BaseTotalAmount))/Cdbl(BaseTotalAmount)
End If
 MonthlyAdjustment=Cdbl(SubTotal) * Cdbl(PercentageIncrease)
 
 NewSubTotal=Cdbl(SubTotal)+ Cdbl(MonthlyAdjustment)+ Cdbl(AdminFee)+ Cdbl(ServiceFee)
 TotalDue =Cdbl(NewSubTotal)+ Cdbl(SalesTax)
 BilledAmount=SubTotal
 BalanceDue=TotalDue-BilledAmount
%>

<!-- header -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD><BODY>
<P>
<TABLE cellSpacing=1 cellPadding=1 width="100%" align=left border=0>
  <TR>
    <TD colspan=9><strong><%=utilityname%> Bill</strong></TD>
  </TR>
  <TR align="middle">
    <TD>&nbsp;&nbsp;&nbsp;</TD>
    <TD>From</TD>
    <TD>TO</TD>
    <TD>&nbsp;&nbsp;&nbsp;</TD>
    <TD>Surveyed KWH</TD>
    <TD>Surveyed KW</TD>
    <TD>&nbsp;&nbsp;&nbsp;</TD>
    <TD>Building Kwh</TD>
    <TD>Building KW</TD>
  </TR>
  <TR align="middle">
    <TD></TD>
    <TD><%=DateStart%></TD>
    <TD><%=DateEnd%></TD>
    <TD></TD>
    <TD>N/A</TD>
    <TD>N/A</TD>
    <TD></TD>
    <TD><%=BuildingKWH%></TD>
    <TD><%=BuildingKW%></TD>
  </TR>
  </TABLE></P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P> </P>
<TABLE cellSpacing=1 cellPadding=1 width="40%" align=left border=0>
  
  <TR>
    <TD><b>Service Class</b></TD>
    <TD width="60%"><%=serviceclass%></TD>
  </TR>
  <TR>
    <TD><b>ERI No.</b></TD>
    <TD>N/A</TD>
  </TR>
  <TR>
    <TD><b>ERIF</b></TD>
    <TD><%=ERIF%></TD>
  </TR>
  <TR>
    <TD><b>SSQFT</b></TD>
    <TD><%=Sqft%></TD>
  </TR>
</TABLE>
<P>&nbsp;</P>
<P><br>&nbsp;</P>
<P>&nbsp;</P>
<P>
<TABLE cellSpacing=1 cellPadding=1 width="100%" align=left border=0>
  <TR>
    <TD><B>Base Consumption Details:</B></TD>
    <TD><B>Monthly Consumption Details:</B></TD>
  </TR>

  <TR>
    <TD><%=ArrDetail1(0)%></TD>
    <TD><%=ArrDetail2(0)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(1)%></TD>
    <TD><%=ArrDetail2(1)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(2)%></TD>
    <TD><%=ArrDetail2(2)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(3)%></TD>
    <TD><%=ArrDetail2(3)%></TD>
  </TR>
  <TR>
    <TD><b>Total Consumption &nbsp;= 
      &nbsp;<%=BaseTotalConsumtion%></b></TD>
    <TD><b>Total Consumption &nbsp;= 
      &nbsp;<%=TotalConsumtion%></b></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(5)%></TD>
    <TD><%=ArrDetail2(5)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(6)%></TD>
    <TD><%=ArrDetail2(6)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(7)%></TD>
    <TD><%=ArrDetail2(7)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(8)%></TD>
    <TD><%=ArrDetail2(8)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(9)%></TD>
    <TD><%=ArrDetail2(9)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(10)%></TD>
    <TD><%=ArrDetail2(10)%></TD>
  </TR>
  <TR>
    <TD><b>Total Demand<%=ArrDetail1(11)%></b></TD>
    <TD><b>Total Demand<%=ArrDetail2(11)%></b></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(12)%></TD>
    <TD><%=ArrDetail2(12)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(13)%></TD>
    <TD><%=ArrDetail2(13)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(14)%></TD>
    <TD><%=ArrDetail2(14)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(15)%></TD>
    <TD><%=ArrDetail2(15)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(16)%></TD>
    <TD><%=ArrDetail2(16)%></TD>
  </TR>
  <TR>
    <TD><%=ArrDetail1(17)%></TD>
    <TD><%=ArrDetail2(17)%></TD>
  </TR>
</TABLE>
<br>
</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P>&nbsp;</P>
<P><STRONG>Monthly Adjustment Factor:</STRONG></P>
<P>
<TABLE cellSpacing=1 cellPadding=1 width="100%" align=left border=0>  
  <TR>
    <TD>Base Charge</TD>
    <TD>Current Monthly Charge</TD>
    <TD>Percentage Increase</TD>
    <TD></TD>
    <TD></TD>
  </TR>
  <TR>
    <TD><%=BaseTotalAmount%></TD>
    <TD><%=TotalAmount%></TD>
    <TD><%=PercentageIncrease%>%</TD>
    <TD></TD>
    <TD></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Sub Total</TD>
    <TD align="right">$<%=FormatCurrency(SubTotal,2)%></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Mo. Adj.</TD>
    <TD align="right">$<%=FormatCurrency(MonthlyAdjustment,2)%></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Admin Fee</TD>
    <TD align="right"><%=FormatPercent(AdminFee,2)%></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Service Fee</TD>
    <TD align="right">$<%=FormatCurrency(ServiceFee,2)%></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Sub Total</TD>
    <TD align="right">$<%=FormatCurrency(NewSubTotal,2)%></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Sales Tax</TD>
    <TD align="right">$<%=FormatCurrency(SalesTax,2)%></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Total Due</TD>
    <TD align="right">$<%=FormatCurrency(TotalDue,2)%></TD>
  </TR>
  <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Billed Amount</TD>
    <TD align="right">$<%=FormatCurrency(BilledAmount,2)%></TD>
  </TR>
   <TR>
    <TD></TD>
    <TD></TD>
    <TD></TD>
    <TD>Balance Due</TD>
    <TD align="right">$<%=FormatCurrency(BalanceDue,2)%></TD>
  </TR>
 <TR>
    <TD>&nbsp;</TD>
    <TD>&nbsp;</TD>
    <TD>&nbsp;</TD>
    <TD>&nbsp;</TD>
    <TD align="right">&nbsp;</TD>
   
  <TR>
    <TD>
		Tenant Name and Address:<br>
		<b><%=TenantName%> (<%=TenantNumber%>)<br>
		<%=replace(TenantStreet,vbNewLine,"<br>")%><br>
		<%=TenantCity%>, <%=TenantState%>&nbsp;<%=TenantZip%></b><br>
		&nbsp;
    </TD>
    <TD></TD>
    <TD></TD>
    <TD colspan="2">
		Make Check Payable To:<br>
		<%=BillingBuildingNumber%><br>
		<%=replace(BillTo,vbNewLine,"<br>")%><br>
		<%=BillingCity%>, <%=BillingState%>&nbsp;<%=BillingZip%><br>
		<%=paymentterm%></b><br>
    </TD>
  </TR>
  <TR>
	<TD ColSpan="5" Align="Center">
		If you have any questions about your bill, Please call Genergy @ (845) 228-5200
	</TD>
  </TR>  
</TABLE>
</P>
</BODY>
</HTML>
<% end sub%>
