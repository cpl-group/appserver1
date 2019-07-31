<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'12.14.2007 N.Ambo made change for caclulating admin/fee. Adminfee value in the line detail part of the bill was being calulated differntly from the right-hand side. 
'Admin fee in line detail was calculated as : (Energy + demand –credit) * Admin Fee percentage; it has now been changed to remove the -credit in the calculation
'It now uses value admindollar from the query which represents (energy+demand)*admin fee
function getNumber(number)
'	response.write "|"&number&"|"
	if not(isNumeric(number)) then number = 0
	getNumber = number
end function

dim bperiod, building, byear, pid,rpt, pdf, Genergy_Users, utilityid,demo, sql


bperiod = request("bperiod")
building = request("building")
byear = request("byear")
pid = request("pid")
'if pid = "" then pid = getpid(building) end if
Genergy_Users = request("Genergy_Users")
utilityid = trim(request("utilityid"))
demo = request("demo")
if demo = "" then demo = false end if

if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
	if Genergy_Users="True" then setGroup("Genergy Users")
	pdf = true
end if

dim rst1, rst2,rst3, cnn1
set rst1 = server.createobject("ADODB.Recordset")
set rst2 = server.createobject("ADODB.Recordset")
set rst3 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
cnn1.open getLocalConnect(building)


'get utility display labels
dim usage, demand, utilityname
rst1.open "SELECT umeasure as usage, dmeasure as demand, utilitydisplay as utility FROM tblutility WHERE UtilityId="&utilityid, getConnect(pid,building,"Billing")

if not rst1.eof then 
	usage = rst1("usage")
	demand = rst1("demand")
	utilityname = rst1("utility")
end if
rst1.close

' # Added by Tarun  07/10/2006
	Dim lngTenantCount,PAlngTenantCount,MAlngTenantCount, AMlngTenantCount
	lngTenantCount = 0
	PAlngTenantCount = 0 ''added by Andy 9/19/06 'val initialize below before loop
	MAlngTenantCount = 0  ''added by Andy 9/29/06 'val initialize below before loop
	AMlngTenantCount = 0
 ' #

if trim(bperiod)="" or trim(byear)="" then
	rst1.open "select top 1 BillYear, BillPeriod from tblmetersbyperiod WHERE bldgnum='"&building&"' and utility="&utility&" ORDER BY billyear desc, billperiod desc", cnn1
	if rst1.eof then
		response.write "No information for this building"
		response.end
	else
		byear = cint(rst1("billyear"))
		bperiod = cint(rst1("billperiod"))
	end if
	rst1.close
end if
dim DBlocalIP
if trim(building)<>"" then DBlocalIP = ""

dim summarylogo
rst1.open "SELECT logo FROM portfolio p, billtemplates bt WHERE p.templateid=bt.id and p.id=(SELECT portfolioid FROM "&DBlocalIP&"buildings WHERE bldgnum='"&building&"')", getConnect(pid,building,"Billing")

if not rst1.eof then summarylogo = rst1("logo")
rst1.close
if allowGroups("Genergy Users") then rpt = "rpt_bill_summary" else rpt = "rpt_Bill_summary_client"

sql = "SELECT isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.ibsexempt, " & _
			 " r.unit_credit, b.strt, r.sqft, r.adminfee, isnull(r.RateModify,0) as RateModify, " & _
			 " isnull(r.fuelAdj,0) as fuelAdjnum, rt.[type], AvgKWH, tenantname, datestart-1 as datestart, " & _
			 " datediff(day, datestart-1, dateend) as days, dateend, ypid, r.leaseutilityid, r.billingname, " & _
			 " r.tenantnum, isnull(r.adjustment,0)-isnull(r.credit,0) as credit, " & _
			  " isnull(energy,0) + isnull(demand,0) + (isnull(r.adjustment,0)-isnull(r.credit,0)) as subtotal, " & _
			 " isnull(energy,0) as energy, isnull(demand,0) as demand, isnull(serviceFee,0) as serviceFee, " & _
			 " isnull(tax,0) as tax, isnull(totalamt,0) as totalamt, b.btstrt, lup.calcintpeak, " & _
			 " isnull(adminfeedollar,0) as admindollar, isnull(r.extusg,0) as extusg, r.rate_servicefee, r.shadow " & _
	  " FROM "&rpt&" r, dbo.ratetypes rt, buildings b, tblleasesutilityprices lup, tblleases l " & _
	  " WHERE r.reject=0 and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid " & _
				" and b.bldgnum=r.bldgnum AND r.[type]=rt.id and billyear="&byear & _
				" and billperiod="&bperiod&" and r.bldgnum='"&building & _
				"' and l.billsummaryexempt = 0 and r.utility="&utilityid & _
	 " ORDER BY TenantName"
'rateservicefee_dollar
'response.write sql
'response.end

rst1.open sql, cnn1
%>
<html><head><title>Bill Summary</title></head><body bgcolor="#FFFFFF">
<basefont size="1" face="Arial,Helvetica,sans-serif">
<%
if rst1.eof then
	rst1.close
	if allowGroups("Genergy Users") then rpt = "rpt_bill_summary_nobill" else rpt = "rpt_Bill_summary_nobill_client"
	'response.write rpt
	'response.end
	sql = "SELECT isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.ibsexempt, r.unit_credit, " & _
				" b.strt, r.sqft, isnull(r.adminfee,0) as adminfee, isnull(r.RateModify,0) as RateModify, " & _
				" isnull(r.fuelAdj,0) as fuelAdjnum, rt.[type], AvgKWH, tenantname, datestart-1 as datestart, " & _
				" datediff(day, datestart-1, dateend) as days, dateend, ypid, r.leaseutilityid, r.billingname, " & _
				" r.tenantnum, isnull(r.adjustment,0)-isnull(r.credit,0) as credit, isnull(subtotal,0) as subtotal, " & _
				" isnull(energy,0) as energy, isnull(demand,0) as demand, isnull(serviceFee,0) as serviceFee, " & _
				" isnull(tax,0) as tax, isnull(totalamt,0) as totalamt, b.btstrt, lup.calcintpeak, " & _
				" isnull(adminfeedollar,0) as admindollar, isnull(r.extusg,0) as extusg, lup.calcintpeak, " & _
				" r.rate_servicefee, r.shadow " & _
		 " FROM "&rpt&" r, dbo.ratetypes rt, buildings b, tblleasesutilityprices lup, tblleases l " & _
		 " WHERE r.reject=0 and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid " & _
				" and b.bldgnum=r.bldgnum AND r.[type]=rt.id and billyear="&byear & _
				" and billperiod="&bperiod&" and r.bldgnum='"&building&"' and l.billsummaryexempt = 0 " & _
				" and r.utility="&utilityid&_
		" ORDER BY TenantName"
	rst1.open sql, cnn1
'response.write sql
'response.end
end if

dim salestaxsupply, salestaxtd, newtotalamount, bldgtotaltaxsupply, bldgtotaltaxtd, bldgnewtotalamount  'rsm
dim bldgOnPeak, bldgOffPeak,bldgintpeak, bldgTotalPeak, bldgTotalKWon, bldgTotalKWoff, bldgTotalKWint, bldgAdmin, bldgService, bldgCredit, bldgSubtotal, bldgTax, bldgTotalAmt, ExmpOnPeak, ExmpIntPeak, ExmpOffPeak, ExmpTotalPeak, ExmpTotalKW, ExmpTotalKWon, ExmpTotalKWoff, ExmpTotalKWint, ExmpTotalAmt, ExmpData, ExmpAdmin, ExmpService, ExmpCredit, ExmpSubtotal, ExmpTax, Exmpsubsubtotal
dim totaldemand_PC, totaldemand_PCint, totaldemand_PCoff, totalOnpeak, totalOffPeak, totalIntPeak, totalKWH, meterdemandtemp, subsubtotal, meterdemandtempint, meterdemandtempoff,totalkwhoff, totalkwhint

dim usagedivisor
'Port Authority vars
dim PAtotalOnPeak,PAtotalOffPeak,PAtotalKWH,PAtotaldemand_PC,PAtotaldemand_PCoff,PAtotaldemand_PCint
dim PAbldgOnPeak,PAbldgOffPeak,PAbldgintpeak,PAtotalIntPeak,PAbldgTotalPeak ,PAbldgTotalKWon ,PAbldgTotalKWoff,PAbldgTotalKWint,PAbldgAdmin,PAtotalKWHoff,PAtotalKWHint 
dim PAbldgService,PAbldgCredit,PAbldgSubtotal,PAbldgTax,PAsubsubtotal,PAbldgTotalAmt,PACondition

'Master PA
dim MAtotalOnPeak,MAtotalOffPeak,MAtotalKWH,MAtotaldemand_PC,MAtotaldemand_PCoff,MAtotaldemand_PCint
dim MAbldgOnPeak,MAbldgOffPeak,MAbldgintpeak,MAtotalIntPeak,MAbldgTotalPeak ,MAbldgTotalKWon ,MAbldgTotalKWoff,MAbldgTotalKWint,MAbldgAdmin,MAtotalKWHoff,MAtotalKWHint 
dim MAbldgService,MAbldgCredit,MAbldgSubtotal,MAbldgTax,MAsubsubtotal,MAbldgTotalAmt,MACondition

'Audit Meters
dim AMtotalOnPeak,AMtotalOffPeak,AMtotalKWH,AMtotaldemand_PC,AMtotaldemand_PCoff,AMtotaldemand_PCint
dim AMbldgOnPeak,AMbldgOffPeak,AMbldgintpeak,AMtotalIntPeak,AMbldgTotalPeak ,AMbldgTotalKWon ,AMbldgTotalKWoff,AMbldgTotalKWint,AMbldgAdmin,AMtotalKWHoff,AMtotalKWHint 
dim AMbldgService,AMbldgCredit,AMbldgSubtotal,AMbldgTax,AMsubsubtotal,AMbldgTotalAmt,AMCondition



select case utilityid
case 3
	usagedivisor = 100
case else
	usagedivisor = 1
end select 

'Zero out building peak numbers
bldgonpeak 	= 0
bldgoffpeak = 0
bldgintpeak = 0
'PA var setup
PAbldgonpeak  = 0
PAbldgoffpeak = 0
PAbldgintpeak = 0

'MA var setup
MAbldgonpeak  = 0
MAbldgoffpeak = 0
MAbldgintpeak = 0

'AM var setup
AMbldgonpeak  = 0
AMbldgoffpeak = 0
AMbldgintpeak = 0

'rsm
bldgtotaltaxsupply = 0.00
bldgtotaltaxtd = 0.00
bldgnewtotalamount = 0.00


if not rst1.eof and trim(request("noheader"))="" then%>
<a href="http://pdfmaker.genergyonline.com/pdfMaker/pdfBillSummary20EX.asp?pid=<%=pid%>&genergy2=true&demo=<%=demo%>&building=<%=building%>&byear=<%=byear%>&bperiod=<%=bperiod%>&devIP=<%=request.servervariables("SERVER_NAME")%>&strt=<%=server.urlencode(rst1("Strt"))%>&utilitydisplay=<%=utilityname%>&logo=<%=summarylogo%>&utilityid=<%=utilityid%>&Genergy_Users=<%=allowGroups("Genergy Users")%>" target="_blank">Download printable PDF of Bill Summary</a>
<table width="100%" border="0" bgcolor="#FFFFFF"><tr>
<td height="68"><img src="http://pdfmaker.genergyonline.com/pdfMaker/<%=summarylogo%>" hspace="0" width="202" height="143"></td>
<td width="90%" valign="top" align="center"><b><%if demo then%>Demo Property<%else%><%=rst1("Strt")%><%end if%></b><br>Submetering Summary Report</td>
<td width="10%" valign="bottom">&nbsp;<br>&nbsp;<table border="0" cellspacing="2" cellpadding="3" bgcolor="#000000">
<tr bgcolor="white"><td align="center"><font size="1">Bill&nbsp;Year</font></td><td align="center"><font size="1">Bill&nbsp;Period</font></td><td align="center"><font size="1">Utility</font></td></tr>
<tr bgcolor="white"><td align="center"><font size="1"><%=byear%></font></td><td align="center"><font size="1"><%=bperiod%></font></td><td align="center" nowrap><font size="1"><%=utilityname%></font></td></tr>
</table></td></tr></table>
<%end if

dim pagepart, calcintpeak, sql2, sInvoiceNo, extusageflag
'response.write "nothing"'rst1("calcintpeak")
'response.end
if rst1("calcintpeak")="True" then calcintpeak = true else calcintpeak = false
extusageflag = false
pagepart = 1
do until rst1.eof
	totaldemand_PC 		= 0
	totaldemand_PCint 	= 0
	totaldemand_PCoff 	= 0
	totalOnpeak 		= 0
	totalOffPeak 		= 0
  	totalIntPeak 		= 0
	totalKWH 			= 0
	totalKWHoff 		= 0
	totalKWHint 		= 0
	'rsm
	salestaxsupply      = 0
	salestaxtd          = 0
	extusage 			= rst1("extusg")
	
	'rsm
	salestaxsupply = cdbl(rst1("energy")) * cDbl(.08875)
	salestaxtd = (cdbl(rst1("demand")) + cdbl(rst1("credit"))) * cDbl(.045)
	newtotalamount = cdbl(rst1("subtotal")) + formatnumber(cdbl(salestaxsupply),2) + formatnumber(cdbl(salestaxtd),2)
	bldgtotaltaxsupply = bldgtotaltaxsupply + formatnumber(cdbl(salestaxsupply),2)
	bldgtotaltaxtd = bldgtotaltaxtd + formatnumber(cdbl(salestaxtd),2)
	bldgnewtotalamount = bldgnewtotalamount + formatnumber(cdbl(newtotalamount),2)
	'rsm end
	
	
	if  extusage  then
		extusageflag = true
	end if
    'PA vars
    PAtotaldemand_PC 	= 0
	PAtotaldemand_PCint = 0
	PAtotaldemand_PCoff = 0
	PAtotalOnpeak 		= 0
	PAtotalOffPeak 		= 0
  	PAtotalIntPeak 		= 0
	PAtotalKWH 			= 0
	PAtotalKWHoff 		= 0
	PAtotalKWHint 		= 0
  
	'MA vars
    MAtotaldemand_PC 	= 0
	MAtotaldemand_PCint = 0
	MAtotaldemand_PCoff = 0
	MAtotalOnpeak 		= 0
	MAtotalOffPeak 		= 0
  	MAtotalIntPeak 		= 0
	MAtotalKWH 			= 0
	MAtotalKWHoff 		= 0
	MAtotalKWHint 		= 0
	
	' AM Vars
    AMtotaldemand_PC 	= 0
	AMtotaldemand_PCint = 0
	AMtotaldemand_PCoff = 0
	AMtotalOnpeak 		= 0
	AMtotalOffPeak 		= 0
  	AMtotalIntPeak 		= 0
	AMtotalKWH 			= 0
	AMtotalKWHoff 		= 0
	AMtotalKWHint 		= 0


	'sql = "SELECT m.*, b.*, C.Used, C.Estimated as UEstimated,P.Demand, P.estimated as DEstimated FROM tblmetersbyperiod m, tblbillbyperiod b,  billyrperiod bp, Consumption C, PeakDemand P " & _
	'			" WHERE b.id=m.bill_id and b.reject=0 and m.leaseutilityid="&cdbl(rst1("leaseutilityid"))& _
	'			" and m.ypid="&cdbl(rst1("ypid"))& _
	'			" and bp.YpId="&cdbl(rst1("ypid"))& _
	'			" and C.MeterId = m.MeterId	and C.billYear = bp.BillYear and C.BillPeriod = bp.BillPeriod " & _
	'			" and P.MeterId = m.MeterId	and P.billYear = bp.BillYear and P.BillPeriod = bp.BillPeriod " & _
	'	" ORDER BY meternum"
	sql = "SELECT m.*, b.*, C.Estimated as UEstimated,P.Demand, P.estimated as DEstimated FROM tblmetersbyperiod m, tblbillbyperiod b,  billyrperiod bp, Consumption C, PeakDemand P " & _
				" WHERE b.id=m.bill_id and b.reject=0 and m.leaseutilityid="&cdbl(rst1("leaseutilityid"))& _
				" and m.ypid="&cdbl(rst1("ypid"))& _
				" and bp.YpId="&cdbl(rst1("ypid"))& _
				" and C.MeterId = m.MeterId	and C.billYear = bp.BillYear and C.BillPeriod = bp.BillPeriod " & _
				" and P.MeterId = m.MeterId	and P.billYear = bp.BillYear and P.BillPeriod = bp.BillPeriod " & _
		" ORDER BY meternum"
	
	'response.write sql
	'response.end

	rst2.open sql, cnn1
	

	if pid = "108"  then   
		sInvoiceNo = ""
		if cint(utilityid) = 2 then
			sql2 = " SELECT b.id,  RIGHT(REPLACE(RTRIM(SPACE(5) + STR(BN.InvoiceSeqNo)),' ','0'),5) AS InvoiceSeqNo, BN.BillType  " & _
				   " FROM tblbillbyperiod b, tblPAInvoiceBillNumbers BN"  & _	
				   " WHERE " & _
						" b.id = BN.BillId " & _
						" AND b.reject=0 " & _
						" AND b.leaseutilityId = " & rst1("leaseutilityid") & _
						" AND b.ypid = " & rst1("ypid")
		else
			sql2 = " SELECT b.id,  RIGHT(REPLACE(RTRIM(SPACE(5) + STR(BN.InvoiceSeqNo)),' ','0'),5) AS InvoiceSeqNo, BN.BillType  " & _
				   " FROM tblbillbyperiod b, tblPAWaterBillNumbers BN"  & _	
				   " WHERE " & _
						" b.id = BN.BillId " & _
						" AND b.reject=0 " & _
						" AND b.leaseutilityId = " & rst1("leaseutilityid") & _
						" AND b.ypid = " & rst1("ypid")
		
		end if
		rst3.Open sql2, cnn1
		if not rst3.EOF then
			If not isNull(rst3("BillType")) Then 	
				sInvoiceNo = rst3("BillType") & rst3("InvoiceSeqNo")
			Else
				sInvoiceNo = "0000000" ' Indicates an Error While Generating The Invoice Number
			End If
		end if	
		rst3.Close 
	end if

	%>
	
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<tr><td colspan="2" bgcolor="#3399CC">
			<% if pid <> "108" then %>
			<font size="2" color="#FFFFFF">
				<b><%if rst1("ibsexempt")="True" then response.write "*"%>
				<%if demo then%>Demo Tenant
				<%else%><%=rst1("billingname")%>
				<%end if%> (<%=rst1("TenantNum")%>)</b></font>
			<% else %>
			<font size="2" color="#FFFFFF">
				<b><%if rst1("ibsexempt")="True" then response.write "*"%>
				<%if demo then%>Demo Tenant
				<%else%><%=rst1("billingname")%>
				<%end if%> (<%=rst1("TenantNum")%>) Invoice No: <%=sInvoiceNo%></b></font>
			<% end if %>
		</td>
	</tr>
	<tr><td width="80%" valign="top"><!-- outer table -->
	
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<% if extusage then %>
		<tr bgcolor="#CCCCCC">
			<td colspan="2" align="center"><font size="1">Tenant</font></td>
			<td colspan="3" align="center"><font size="1">Readings</font></td>
			<td colspan="1" align="center"><font size="1">Consumption</font></td>
			<%if utilityid=2 or utilityid=1 or utilityid=6 then%><td align="center"><font size="1">Demand</font></td><%end if%>
		</tr>
		<tr bgcolor="#CCCCCC" align="center">
			<td><font size="1">Meter No.</font></td>
			<td><font size="1">Multi.</font></td>
			<td>&nbsp;</td>
			<td><font size="1">Previous</font></td>
			<td><font size="1">Current</font></td>
			<td><font size="1">Total Usage (<%if utilityid=3 then%>C<%end if%><%=usage%>)</font></td>
			<%if utilityid=2 or utilityid=1 or utilityid=6 then%><td><font size="1"><%=demand%></font></td><%end if%>
		</tr>
	<% else %>
		<tr bgcolor="#CCCCCC">
			<td colspan="2" align="center"><font size="1">Tenant</font></td>
			<td colspan="2" align="center"><font size="1">Readings</font></td>
			<td colspan="<%if calcintpeak then%>4<%else%>3<%end if%>" align="center"><font size="1">Consumption</font></td>
			<%if utilityid=2 or utilityid=1 or utilityid=6 then%>
				<td align="center" <%if calcintpeak then%>colspan="3"<%end if%>><font size="1">Demand</font></td>
			<%end if%>
		</tr>
		<tr bgcolor="#CCCCCC" align="center">
			<td><font size="1">Meter No.</font></td>
			<td><font size="1">Multi.</font></td>
			<td><font size="1">Previous</font></td>
			<td><font size="1">Current</font></td>
			<td><font size="1">On Peak</font></td>
			<%if calcintpeak then%><td><font size="1">Int Peak</font></td><%end if%>
			<td><font size="1">Off Peak</font></td>
			<td><font size="1">Total Usage (<%if utilityid=3 then%>C<%end if%><%=usage%>)</font></td>
			<%if utilityid=2 or utilityid=1 or utilityid=6 then%>
				<td><font size="1"><%if calcintpeak then%>On<%else%><%=demand%><%end if%></font></td>
				<%if calcintpeak then%>
					<td><font size="1">Off</font></td>
					<td><font size="1">Int</font></td>
				<%end if%>
			<%end if%>
		</tr>
	<% end if %>
	<% 
	dim metercount, intpeak, extusage,tester,PAFlag,flag,MAflag',PAlngTenantCount
 	dim AMFlag
 	
	extusage = rst1("extusg")
	metercount = 0
	'PAlngTenantCount=0
	do until rst2.eof
		metercount = metercount+1
		meterdemandtemp = rst2("Demand_P")
		meterdemandtempint = rst2("Demand_int")
		meterdemandtempoff = rst2("Demand_off")
	    tester = InStr(rst2("tenantnum"),"MA")
		PAFlag = InStr(rst2("tenantnum"),"PA") 
		MAFlag = InStr(rst2("tenantnum"),"MA")
		AMFlag = InStr(rst2("tenantnum"),"AM") 		
      if pid = "108" then 
		if PAFlag > 0  then flag = true else flag = false end if
		if tester > 0  then MAflag = true else MAflag = false end if
        
		if flag or MAFlag then PACondition = true
		
		if AMFlag > 0 then 
			AMFlag=true 
		else
			AMFlag=false
		end if
	  end if
       'response.write PACondition
	   'response.end 
		intpeak = 0
		if isnumeric(rst2("IntPeak")) then intpeak = cdbl(rst2("IntPeak"))
		if rst2("coincident")="True" then
			meterdemandtemp = 0
			if not flag and not MAflag and not AMFlag  then
				totaldemand_PC = cdbl(rst2("Demand_C")) + cdbl(rst2("Demand_int")) + cdbl(rst2("Demand_off"))
			end if
			if flag then
           		PAtotaldemand_PC = cdbl(rst2("Demand_C")) + cdbl(rst2("Demand_int")) + cdbl(rst2("Demand_off"))
		    end if
			if MAflag then
				MAtotaldemand_PC = cdbl(rst2("Demand_C")) + cdbl(rst2("Demand_int")) + cdbl(rst2("Demand_off"))
			end if
			if AMFlag then 
				AMtotaldemand_PC = cdbl(rst2("Demand_C")) + cdbl(rst2("Demand_int")) + cdbl(rst2("Demand_off"))
			end if
			
		else
			if not flag and not MAflag and not AMFlag then
				totaldemand_PC = totaldemand_PC + cdbl(meterdemandtemp)
			end if
			
			if flag then
				PAtotaldemand_PC = PAtotaldemand_PC + cdbl(meterdemandtemp)
			end if
			
			if MAflag then
				MAtotaldemand_PC = MAtotaldemand_PC + cdbl(meterdemandtemp)
			end if

			if AMflag then
				AMtotaldemand_PC = AMtotaldemand_PC + cdbl(meterdemandtemp)
			end if
			
			if calcintpeak then
				if not flag and not MAflag and not AMFlag then
					totaldemand_PCoff = totaldemand_PCoff + cdbl(meterdemandtempoff)
					totaldemand_PCint = totaldemand_PCint + cdbl(meterdemandtempint)
				end if
				'condition might not be in use for PA
				if flag then
					PAtotaldemand_PCoff = PAtotaldemand_PCoff + cdbl(meterdemandtempoff)
					PAtotaldemand_PCint = PAtotaldemand_PCint + cdbl(meterdemandtempint)
				end if
				
				if MAflag then
			 		MAtotaldemand_PCoff = MAtotaldemand_PCoff + cdbl(meterdemandtempoff)
					MAtotaldemand_PCint = MAtotaldemand_PCint + cdbl(meterdemandtempint)
				end if
				
				if AMflag then
			 		AMtotaldemand_PCoff = AMtotaldemand_PCoff + cdbl(meterdemandtempoff)
					AMtotaldemand_PCint = AMtotaldemand_PCint + cdbl(meterdemandtempint)
				end if

			end if
		end if
		
		if not flag and not MAflag and not AMFlag then
    		totalOffPeak = totalOffPeak + formatnumber(rst2("OffPeak"),2)
		end if
		'PA condition goes here
		
		if flag then
			PAtotalOffPeak = PAtotalOffPeak + formatnumber(rst2("OffPeak"),2)
		end if
		
		if MAflag then
			MAtotalOffPeak = MAtotalOffPeak + formatnumber(rst2("OffPeak"),2)
		end if
		
		if AMflag then
			AMtotalOffPeak = AMtotalOffPeak + formatnumber(rst2("OffPeak"),2)
		end if

		'PAlngTenantCount =PAlngTenantCount + 1
		
    if calcintpeak then 
		totalIntPeak 	= totalIntPeak + formatnumber(rst2("IntPeak"),2)
		totalOnPeak 	= totalOnpeak + formatnumber(rst2("OnPeak"),2)
	else
		if not flag and not MAflag and not AMFlag then
			totalOnpeak = totalOnpeak + formatnumber(cdbl(rst2("OnPeak"))+IntPeak,2)
		end if 
		'PA condition goes here
		if flag then
			PAtotalOnpeak = PAtotalOnpeak + formatnumber(cdbl(rst2("OnPeak"))+IntPeak,2)
		end if
		
		if MAflag then
			MAtotalOnpeak = MAtotalOnpeak + formatnumber(cdbl(rst2("OnPeak"))+IntPeak,2)
		end if

		if AMFlag then 
			AMtotalOnpeak = AMtotalOnpeak + formatnumber(cdbl(rst2("OnPeak"))+IntPeak,2)
		end if
			
	end if
		if extusage then 
			if rst2("mextusg") then 
				if not flag and not MAflag and not AMFlag then
				totalKWH 	= totalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2)
				
				totalkwhoff = totalKWHoff + formatnumber(cdbl(rst2("Usedoff"))/usagedivisor,2)
				totalkwhint = totalKWHint + formatnumber(cdbl(rst2("Usedint"))/usagedivisor,2)
				end if
				'PA condition goes here
				if flag then
					PAtotalKWH 	= PAtotalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2)
					PAtotalkwhoff = PAtotalKWHoff + formatnumber(cdbl(rst2("Usedoff"))/usagedivisor,2)
					PAtotalkwhint = PAtotalKWHint + formatnumber(cdbl(rst2("Usedint"))/usagedivisor,2)
				end if
				
				if MAflag then
					MAtotalKWH 	= MAtotalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2)
					MAtotalkwhoff = MAtotalKWHoff + formatnumber(cdbl(rst2("Usedoff"))/usagedivisor,2)
					MAtotalkwhint = MAtotalKWHint + formatnumber(cdbl(rst2("Usedint"))/usagedivisor,2)				
				end if

				if AMflag then
					AMtotalKWH 	= AMtotalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2)
					AMtotalkwhoff = AMtotalKWHoff + formatnumber(cdbl(rst2("Usedoff"))/usagedivisor,2)
					AMtotalkwhint = AMtotalKWHint + formatnumber(cdbl(rst2("Usedint"))/usagedivisor,2)				
				end if

				
			else
				if not flag and not MAflag and not AMFlag then
					totalKWH 	= totalKWH + formatnumber(cdbl(rst2("onpeak"))/usagedivisor,2)
					totalkwhoff = totalKWHoff + formatnumber(cdbl(rst2("offpeak"))/usagedivisor,2)
					totalkwhint = totalKWHint + formatnumber(cdbl(rst2("intpeak"))/usagedivisor,2)
					
				end if
				
				'PA condition goes here
				if flag then
					PAtotalKWH 	= PAtotalKWH + formatnumber(cdbl(rst2("onpeak"))/usagedivisor,2)
					PAtotalkwhoff = PAtotalKWHoff + formatnumber(cdbl(rst2("offpeak"))/usagedivisor,2)
					PAtotalkwhint = PAtotalKWHint + formatnumber(cdbl(rst2("intpeak"))/usagedivisor,2)
				end if
				
				if MAflag then
					MAtotalKWH 	= MAtotalKWH + formatnumber(cdbl(rst2("onpeak"))/usagedivisor,2)
					MAtotalkwhoff = MAtotalkwhoff + formatnumber(cdbl(rst2("offpeak"))/usagedivisor,2)
					MAtotalkwhint = MAtotalkwhint + formatnumber(cdbl(rst2("intpeak"))/usagedivisor,2)
				end if
				
				if AMflag then
					AMtotalKWH 	= AMtotalKWH + formatnumber(cdbl(rst2("onpeak"))/usagedivisor,2)
					AMtotalkwhoff = AMtotalkwhoff + formatnumber(cdbl(rst2("offpeak"))/usagedivisor,2)
					AMtotalkwhint = AMtotalkwhint + formatnumber(cdbl(rst2("intpeak"))/usagedivisor,2)
				end if

			end if
		else
			if not flag and not MAflag and not AMFlag then
				'if rst2("mextusg") then
				
				'else
					totalKWH 	= totalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2) 
				'end if
			end if
			
			'if condition to distinguish between PA tenants and regular tenants     
			if flag then
				PAtotalKWH 	= PAtotalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2) 
			end if
			
			if MAflag then
				MAtotalKWH 	= MAtotalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2)
			end if
			
			if AMflag then
				AMtotalKWH 	= AMtotalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2)
			end if


		end if 
		
		if extusage and rst2("mextusg") then 
			metercount = metercount + 2
		%>
		<tr>
			<td><font size="1"><%=rst2("Meternum")%></font></td>
			<td><font size="1"><%=rst2("ManualMultiplier")%>&nbsp;&nbsp;</font></td>
			
          <td><font size="1">On Peak</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawPrevious"),2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawCurrent"),2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(cdbl(rst2("Used"))/usagedivisor,2)%>&nbsp;&nbsp;</font></td>
			<%if utilityid=2 or utilityid=1 or utilityid=6 then%><td><font size="1"><%=formatnumber(meterdemandtemp,2)%>&nbsp;&nbsp;</font></td><%end if%>
		</tr>
		<tr>
			<td colspan=2>&nbsp;</td>
			
          <td><font size="1">Off Peak</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawPreviousoff"),2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawCurrentoff"),2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(cdbl(rst2("Usedoff"))/usagedivisor,2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=Formatnumber(meterdemandtempoff,2)%></font></td>
		</tr>
		<tr>
			<td colspan=2>&nbsp;</td>
            <td><font size="1">Mid Peak</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawPreviousint"),2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawCurrentint"),2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(cdbl(rst2("Usedint"))/usagedivisor,2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=Formatnumber(meterdemandtempint,2)%></font></td>
		</tr>
		<% else %>
		<tr>
			<td><font size="1"><%=rst2("Meternum")%></font></td>
			<td><font size="1"><%=rst2("ManualMultiplier")%>&nbsp;&nbsp;</font></td>
			<%if extusage then%><td></td><%end if%>
			<td><font size="1"><%=formatnumber(rst2("rawPrevious"),2)%>&nbsp;&nbsp;</font></td>
			<td><font size="1"><%=formatnumber(rst2("rawCurrent"),2)%>&nbsp;&nbsp;</font></td>
			<%if not extusage then%>
				<%if calcintpeak then%>
					<td><font size="1"><%=formatnumber(cdbl(rst2("OnPeak"))/usagedivisor,2)%>&nbsp;&nbsp;</font></td>
					<td><font size="1"><%=formatnumber(cdbl(rst2("IntPeak"))/usagedivisor,2)%>&nbsp;&nbsp;</font></td>			
				<%else%>
					<% if utilityid<>4 then%>
					<td><font size="1"><%=formatnumber((cdbl(rst2("OnPeak"))+IntPeak)/usagedivisor,2)%>&nbsp;&nbsp;</font></td>
					<% else %>
					<td><font size="1">&nbsp;&nbsp;&nbsp;</font></td>
					<% end if %>
		        <%end if%>
					<% if utilityid<>4 then%>
					<td><font size="1"><%=formatnumber(cdbl(rst2("OffPeak"))/usagedivisor,2)%>&nbsp;&nbsp;</font></td>
					<% else %>
					<td><font size="1">&nbsp;&nbsp;&nbsp;</font></td>
					<% end if %>
					
			<%end if%>
			<td><font size="1"><%=formatnumber(cdbl(rst2("Used"))/usagedivisor,2)%>&nbsp;&nbsp;
							   <% If  rst2("UEstimated") = "True" then 
									Response.Write "*"
								  End If%> </font></td>
			<%if utilityid=2 or utilityid=1 or utilityid=6 then%>
				<td><font size="1"><%if rst2("coincident")="True" then%>0<%else%><%=formatnumber(meterdemandtemp,2)%><%end if%>&nbsp;&nbsp;
								  <% If  rst2("DEstimated") = "True" then 
									Response.Write "*"
								  End If%></font></td>
		    	  <%if calcintpeak then%>
	  					<td><font size="1"><%if rst2("coincident")="True" then%>0<%else%><%=formatnumber(meterdemandtempint,2)%><%end if%>&nbsp;&nbsp;</font></td>
			  			<td><font size="1"><%if rst2("coincident")="True" then%>0<%else%><%=formatnumber(meterdemandtempoff,2)%><%end if%>&nbsp;&nbsp;</font></td>
			      <%end if%>
			<%end if%>
		</tr>
	
<%	
		end if 
		
	if (metercount>24 or (metercount>11 and pagepart>1)) and trim(request("noheader"))<>"" then
			pagepart = 1
			metercount=1
			%>
			<tr><td colspan="8" align="right"><font size="1">Continued on next page...</font></td></tr>
			</table>
			</tr></table>
			<WxPrinter PageBreak>
			<table cellspacing="1" cellpadding="2" border="0" width="100%">
			<tr><td colspan="2" bgcolor="#3399CC"><font size="2" color="#FFFFFF"><b><%if demo then%>Demo Tenant<%else%><%=rst1("billingname")%><%end if%> (<%=rst1("TenantNum")%>)</b></font></td></tr>
			<tr><td width="80%" valign="top">
			<table cellspacing="1" cellpadding="2" border="0" width="100%">
			<%if extusage then %>
				<tr bgcolor="#CCCCCC">
					<td colspan="2" align="center"><font size="1">Tenant</font></td>
					<td colspan="3" align="center"><font size="1">Readings</font></td>
					<td colspan="1" align="center"><font size="1">Consumption</font></td>
					<td align="center"><font size="1">Demand</font></td>
				</tr>
				<tr bgcolor="#CCCCCC" align="center">
					<td><font size="1">Meter No.</font></td>
					<td><font size="1">Multi.</font></td>
					<td>&nbsp;</td>
					<td><font size="1">Previous</font></td>
					<td><font size="1">Current</font></td>
					<td><font size="1">Total Usage</font></td>
					<%if utilityid=2 or utilityid=1 or utilityid=6 then%><td><font size="1"><%=demand%></font></td><%end if%>
				</tr>
			<%else%>
				<tr bgcolor="#CCCCCC">
					<td colspan="2" align="center"><font size="1">Tenant</font></td>
					<td colspan="2" align="center"><font size="1">Readings</font></td>
					<td colspan="3" align="center"><font size="1">Consumption</font></td>
					<td align="center"><font size="1">Demand</font></td>
				</tr>
				<tr bgcolor="#CCCCCC" align="center">
					<td><font size="1">Meter No.</font></td>
					<td><font size="1">Multi.</font></td>
					<td><font size="1">Previous</font></td>
					<td><font size="1">Current</font></td>
					<td><font size="1">On Peak</font></td>
					<td><font size="1">Off Peak</font></td>
					<td><font size="1">KWHR</font></td>
					<%if utilityid=2 or utilityid=1 or utilityid=6 then%><td><font size="1"><%=demand%></font></td><%end if%>
				</tr>
			<%end if%>
<%		end if
		rst2.movenext
	loop
		
	%>
	<% if extusage then %>
	<tr>
		<td></td>
		<td></td>
		<td><font size="1"><b>Meter Totals</b></font></td>
		<td><font size="1"><b>On&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b>Off&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b>Mid&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b>Total&nbsp;&nbsp;</b></font></td>
	</tr>
	<tr>
		<td></td>
		<td></td>
		<td><font size="1"><b><%=usage%></b></font></td>
		<td><font size="1"><b><%=formatnumber(totalKWH,2)%>&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b><%=formatnumber(totalKWHoff,2)%>&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b><%=formatnumber(totalKWHint,2)%>&nbsp;&nbsp;</b></font></td>
		<td><font size="1"><b><%=formatnumber(totalKWH+totalkwhoff+totalkwhint,2)%>&nbsp;&nbsp;</b></font></td>
	</tr>
	<tr>
		<td></td>
		<td></td>
		<td><font size="1"><b><%=demand%></b></font></td>
		<td><font size="1"><b><%=formatnumber(totaldemand_PC)%></b></font></td>
		<td><font size="1"><b><%=formatnumber(totaldemand_PCoff)%></b></font></td>
		<td><font size="1"><b><%=formatnumber(totaldemand_PCint)%></b></font></td>
		<td><font size="1"><b><%=formatnumber(totaldemand_PC+totaldemand_PCint+totaldemand_PCoff,2)%>&nbsp;&nbsp;</b></font></td>
	</tr>
	<%else %>
	<tr>
		<td></td>
		<td></td>
		<td></td>
		<td><font size="1"><b>Meter Totals</b></font></td>
		<% if utilityid <> 4 then%>
		<td><font size="1"><b><%=formatnumber(totalOnPeak,2)%>&nbsp;&nbsp;</b></font></td>
		<% else %>
		<td><font size="1">&nbsp;&nbsp;&nbsp;</font></td>
		<% end if %>
	
    <%if calcintpeak then%><td><font size="1"><b><%=formatnumber(totalIntPeak,0)%>&nbsp;&nbsp;</b></font></td><%end if%>
		<% if utilityid <> 4  then%>
		<td><font size="1"><b><%=formatnumber(totalOffPeak,2)%>&nbsp;&nbsp;</b></font></td>
		<% else %>
		<td><font size="1">&nbsp;&nbsp;&nbsp;</font></td>
		<% end if %>
	
		<td><font size="1"><b><%=formatnumber(totalKWH,2)%>&nbsp;&nbsp;</b></font></td>
		<%if utilityid=2 or utilityid=1 or utilityid=6 then%>
			<td><font size="1"><b><%=formatnumber(totaldemand_PC,2)%>&nbsp;&nbsp;</b></font></td>
		    <%if calcintpeak then%>
		  		<td><font size="1"><b><%=formatnumber(totaldemand_PCint,2)%>&nbsp;&nbsp;</b></font></td>
		  		<td><font size="1"><b><%=formatnumber(totaldemand_PCoff,2)%>&nbsp;&nbsp;</b></font></td>
		    <%end if%>
	    <%end if%>
	</tr>
	<%end if%>
	</table>
	<%'response.write bldgOnPeak
	if not rst1("ibsexempt") then
		if extusage then 
			if not flag and not MAflag and not AMFlag  then
				bldgOnPeak = bldgOnPeak + totalKWH
				bldgOffPeak = bldgOffPeak + totalkwhoff
				bldgIntPeak = bldgIntPeak + totalkwhint 
				bldgTotalPeak = bldgTotalPeak + totalKWH + totalkwhoff + totalkwhint
			end if
			
			if flag then
				PAbldgOnPeak = PAbldgOnPeak + PAtotalKWH
				PAbldgOffPeak = PAbldgOffPeak + PAtotalkwhoff
				PAbldgIntPeak = PAbldgIntPeak + PAtotalkwhint 
				PAbldgTotalPeak = PAbldgTotalPeak + PAtotalKWH + PAtotalkwhoff + PAtotalkwhint
			end if
			
		    if MAflag then
				MAbldgOnPeak = MAbldgOnPeak + MAtotalKWH
				MAbldgOffPeak = MAbldgOffPeak + MAtotalkwhoff
				MAbldgIntPeak = MAbldgIntPeak + MAtotalkwhint 
				MAbldgTotalPeak = MAbldgTotalPeak + MAtotalKWH + MAtotalkwhoff + MAtotalkwhint
	    	end if
 
		    if AMflag then
				AMbldgOnPeak = AMbldgOnPeak + AMtotalKWH
				AMbldgOffPeak = AMbldgOffPeak + AMtotalkwhoff
				AMbldgIntPeak = AMbldgIntPeak + AMtotalkwhint 
				AMbldgTotalPeak = AMbldgTotalPeak + AMtotalKWH + AMtotalkwhoff + AMtotalkwhint
	    	end if
			

		else
			if not flag and not MAflag and not AMFlag then
				bldgOnPeak = bldgOnPeak + totalOnPeak
			end if
			
			if flag then
				PAbldgOnPeak = PAbldgOnPeak + PAtotalOnPeak
			end if
			
			if MAflag then
				MAbldgOnPeak = MAbldgOnPeak + MAtotalOnPeak 
			end if
				
			if AMflag then
				AMbldgOnPeak = AMbldgOnPeak + AMtotalOnPeak 
			end if


			if calcintpeak then 
				if not flag and not MAflag and not AMFlag then
					bldgIntPeak = bldgIntPeak + totalIntPeak 
				end if
				
				if  flag then
				    PAbldgIntPeak = PAbldgIntPeak + PAtotalIntPeak 
		      	end if
			      	 
				if MAflag then
					MAbldgIntPeak = MAbldgIntPeak + MAbldgIntPeak 
				end if

				if AMflag then
					AMbldgIntPeak = AMbldgIntPeak + AMbldgIntPeak 
				end if				
					
			end if
			
			if not flag and not MAflag and not AMFlag then
				bldgOffPeak = bldgOffPeak + totalOffPeak		
				bldgTotalPeak = bldgTotalPeak + totalKWH 
			end if
			
			if flag then
				PAbldgOffPeak = PAbldgOffPeak + PAtotalOffPeak		
				PAbldgTotalPeak = PAbldgTotalPeak + PAtotalKWH
			end if 
			
			if MAflag then
				MAbldgOffPeak = MAbldgOffPeak + MAtotalOffPeak		
				MAbldgTotalPeak = MAbldgTotalPeak + MAtotalKWH
			end if
			
			if AMflag then
				AMbldgOffPeak = AMbldgOffPeak + AMtotalOffPeak		
				AMbldgTotalPeak = AMbldgTotalPeak + AMtotalKWH
			end if			
				
		end if
		
		if not flag and not MAflag and not AMFlag then
	 		bldgTotalKWon = bldgTotalKWon + totaldemand_PC
			bldgTotalKWoff = bldgTotalKWoff + totaldemand_PCoff
			bldgTotalKWint = bldgTotalKWint + totaldemand_PCint
		end if

		if flag then
			PAbldgTotalKWon =  PAbldgTotalKWon +  PAtotaldemand_PC
			PAbldgTotalKWoff = PAbldgTotalKWoff + PAtotaldemand_PCoff
			PAbldgTotalKWint = PAbldgTotalKWint + PAtotaldemand_PCint
		end if

		if MAflag then
			MAbldgTotalKWon =  MAbldgTotalKWon +  MAtotaldemand_PC
			MAbldgTotalKWoff = MAbldgTotalKWoff + MAtotaldemand_PCoff
			MAbldgTotalKWint = MAbldgTotalKWint + MAtotaldemand_PCint
		end if

		if AMflag then
			AMbldgTotalKWon =  AMbldgTotalKWon +  AMtotaldemand_PC
			AMbldgTotalKWoff = AMbldgTotalKWoff + AMtotaldemand_PCoff
			AMbldgTotalKWint = AMbldgTotalKWint + AMtotaldemand_PCint
		end if
		


			'bldgAdmin = bldgAdmin + (cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDbl(rst1("credit")))*cdbl(rst1("adminfee")) 'removed 12/14/2006 N.Ambo		
		bldgAdmin = bldgAdmin + (cDbl(rst1("energy"))+cDBL(rst1("demand")))*cdbl(rst1("adminfee")) 'added 12/14/2006 N.Ambo to remove the LMP creit in the figure
		if ucase(trim(rst1("type"))) = "LPLS2" then bldgAdmin = bldgAdmin + cdbl(rst1("rate_servicefee"))
		if not flag and not MAflag and not AMFlag then
			bldgService = bldgService + cDbl(rst1("serviceFee"))
			bldgCredit = bldgCredit + cDbl(rst1("credit"))
			bldgSubtotal = bldgSubtotal + cDbl(rst1("subtotal"))
			bldgTax = bldgTax + cDbl(rst1("tax"))
	    end if
		'PA condition goes here
		
		if flag then
			PAlngTenantCount =PAlngTenantCount + 1
			PAbldgTax = PAbldgTax + cDbl(rst1("tax"))
		end if
		
		if MAflag then
			MAlngTenantCount = MAlngTenantCount + 1
			MAbldgTax = MAbldgTax + cDbl(rst1("tax"))
		end if

		if AMflag then
			AMlngTenantCount = AMlngTenantCount + 1
			AMbldgTax = AMbldgTax + cDbl(rst1("tax"))
		end if
		
		bldgTotalAmt = bldgTotalAmt + cDbl(rst1("TotalAmt"))
		
		if not flag and not MAflag and not AMFlag then
			subsubtotal = subsubtotal + cDbl(rst1("energy"))+cDBL(rst1("demand"))+cDBL(rst1("rateservicefee_dollar"))
	    end if

		if flag then
			PAsubsubtotal = PAsubsubtotal + cDbl(rst1("energy"))+cDBL(rst1("demand"))+cDBL(rst1("rateservicefee_dollar"))
		end if

		if MAflag then
			MAsubsubtotal = MAsubsubtotal + cDbl(rst1("energy"))+cDBL(rst1("demand"))+cDBL(rst1("rateservicefee_dollar"))
		end if
		
		if AMflag then
			AMsubsubtotal = AMsubsubtotal + cDbl(rst1("energy"))+cDBL(rst1("demand"))+cDBL(rst1("rateservicefee_dollar"))
		end if
		
	elseif rst1("shadow")="False" then
		if extusage then 
			ExmpOnPeak = ExmpOnPeak + totalKWH
			ExmpOffPeak = ExmpOffPeak + totalkwhoff
			ExmpIntPeak = ExmpIntPeak + totalkwhint 
			ExmpTotalPeak = ExmpTotalPeak + totalKWH + totalkwhoff + totalkwhint
		else
			ExmpOnPeak = ExmpOnPeak + totalOnPeak
			if calcintpeak then 
				ExmpIntPeak = ExmpIntPeak + totalIntPeak 
			end if
			ExmpOffPeak = ExmpOffPeak + totalOffPeak		
			ExmpTotalPeak = ExmpTotalPeak + totalKWH 
		end if
		ExmpTotalKWon = ExmpTotalKWon + totaldemand_PC
		ExmpTotalKWoff = ExmpTotalKWoff + totaldemand_PCoff
		ExmpTotalKWint = ExmpTotalKWint + totaldemand_PCint
		ExmpAdmin = ExmpAdmin + (cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDbl(rst1("credit")))*cdbl(rst1("adminfee"))
		if ucase(trim(rst1("type"))) = "LPLS2" then ExmpAdmin = ExmpAdmin + cdbl(rst1("rate_servicefee"))
		ExmpService = ExmpService + cDbl(rst1("serviceFee"))
		ExmpCredit = ExmpCredit + cDbl(rst1("credit"))
		ExmpSubtotal = ExmpSubtotal + cDbl(rst1("subtotal"))
		ExmpTax = ExmpTax + cDbl(rst1("tax"))
		Exmpsubsubtotal = Exmpsubsubtotal + cDbl(rst1("energy"))+cDBL(rst1("demand"))+cDBL(rst1("rateservicefee_dollar"))
		ExmpTotalAmt = ExmpTotalAmt + cDbl(rst1("TotalAmt"))
		ExmpData = true
	end if
	rst2.close
	%>
	&nbsp;
	<table cellspacing="1" cellpadding="2" border="0" width="100%">
	<tr bgcolor="#CCCCCC" align="center">
		<td align="center"><font size="1">Service Class</font></td>
		<td align="center"><font size="1">Admin Fee</font></td>
		<%if utilityid=2 then%><td align="center"><font size="1">El. Adj. Factor</font></td><%end if%>
		<td align="center"><font size="1"><%if utilityid=3 then%>Sewer<%else%>Demand<%end if%> Charge</font></td>
		<td align="center"><font size="1"><%if utilityid=3 then%>Water<%else%>Consumption<%end if%> Charge</font></td>
		<td align="center"><font size="1"><%if isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" then%>LMEP Rate<%else%>Modify Rate<%end if%></font></td>
		<td align="center"><font size="1"><%if ucase(trim(rst1("type"))) = "LPLS2" then %>Tenant<% end if%>&nbsp;Service Fee</font></td>
		<%if rst1("rate_servicefee")<>"0" then%><td align="center"><font size="1">Utility Service Fee</font></td><%end if%>
		<%if rst1("adminfee")<>"0" then%><td align="center"><font size="1">Admin Fee</font></td><%end if%>
		<td align="center"><font size="1">SqFt</font></td>
		<%if utilityid="2" then%><td align="center"><font size="1">Watts/SqFt</font></td><%end if%>
	</tr>
	<tr>
		<td><nobr><font size="1"><%=rst1("type")%><%if ucase(trim(rst1("type")))="AVG COST 1" then response.write "&nbsp;&nbsp;"&formatnumber(rst1("AvgKWH"),6)%></font></nobr></td>
		<td align="center"><font size="1"><%=formatpercent(rst1("AdminFee"),2)%></font></td>
		<%if utilityid=2 then%><td align="center"><font size="1"><%=formatnumber(rst1("fuelAdjnum"),5)%></font></td><%end if%>
		<td align="center"><font size="1"><%=formatcurrency(rst1("demand"),2)%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(rst1("energy"),2)%></font></td>
		<td align="center"><font size="1"><%if isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" then%><%=formatcurrency(rst1("unit_credit"),6)%><%else%><%=formatcurrency(rst1("RateModify"),6)%><%end if%></font></td>
		<td align="center"><font size="1"><%=formatcurrency(cDbl(rst1("serviceFee")))%></font></td>
		<%if rst1("rate_servicefee")<>"0" then%><td align="center"><font size="1"><%=formatcurrency(cDbl(rst1("rate_servicefee")))%></font></td><%end if%>
		<%if rst1("adminfee")<>"0" then%><td align="center"><font size="1"><%=formatcurrency(cdbl(rst1("admindollar")))%></font></td><%end if%><!--[demand]-[text210])*[adminfee] -->
		<td align="center"><font size="1"><%=getNumber(rst1("sqft"))%></font></td>
		<%if utilityid="2" then%><td align="center"><font size="1"><%if getNumber(rst1("sqft"))=0 then%>0<%else response.write formatnumber((totaldemand_PC*1000)/cDbl(rst1("sqft"))) end if%></font></td><%end if%>
	</tr>
<%if utilityid="2" then%>
	<tr  bgcolor="#CCCCCC"><td align="left" colspan="10"><font size="1">Annualized Cost Per Square Foot For This Bill</font></td></tr>
	<tr><td align="left" colspan="10"><font size="1"><%if getNumber(rst1("sqft"))=0 then%>0<%else response.write formatcurrency((cdbl(rst1("TotalAmt"))*12)/cDbl(rst1("sqft"))) end if%></font></td></tr>
<%end if%>
</table>

	</td><td width="20%" valign="top">
<table cellspacing="2" cellpadding="3" border="0" bgcolor="#000000">
	<tr bgcolor="white">
		<td><font size="1">From</font></td>
		<td><font size="1">To</font></td>
		<td><font size="1">No. Days</font></td>
	</tr><tr bgcolor="white">
		<td><font size="1"><%=rst1("DateStart")%></font></td>
		<td><font size="1"><%=rst1("DateEnd")%></font></td>
		<td><font size="1"><%=rst1("days")%></font></td>
	</tr>
	</table>
	<%

	if pid = "" then pid = 0
	%>
		<table cellspacing="1" cellpadding="2" border="0">
	<tr align="right"><td><font size="1">Sub Total:</font></td><td><font size="1"><%=formatcurrency(cDbl(rst1("energy"))+cDBL(rst1("demand")),2)%></font></td></tr>

	<%if cdbl(rst1("rateservicefee_dollar"))>0 then%><tr align="right"><td><font size="1">Rate&nbsp;Service&nbsp;Fee:</font></td><td><font size="1"><%=formatcurrency(rst1("rateservicefee_dollar"),2)%></font></td></tr><%end if%>
	<%if pid=49 then%><tr align="right"><td><font size="1">Admin/Service Fee:</font></td><td><font size="1"><%=formatcurrency(cdbl(rst1("servicefee"))+cdbl(rst1("admindollar")),2)%></font></td></tr><%end if%>
	<%if pid<>49 then'not sjp properties%><tr align="right"><td><font size="1">Admin/Service Fee:</font></td><td><font size="1"><%=formatcurrency(cdbl(rst1("servicefee"))+cdbl(rst1("admindollar")),2)%></font></td></tr><%end if%>
	<tr align="right"><td><font size="1"><%if isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" then%>LMEP Credit:<%elseif pid=49 then%>Restructuring Rate Reduction:<%else%>Credit/Adjustment:<%end if%></font></td><td><font size="1"><%=formatcurrency(rst1("credit"),2)%></font></td></tr>
	<tr align="right"><td><font size="1">Sub Total:</font></td><td><font size="1"><%=formatcurrency(rst1("subtotal"),2)%></font></td></tr>
	<tr align="right"><td><font size="1">Sales Tax(Supply):</font></td><td><font size="1"><%=formatcurrency(cDbl(salestaxsupply),2)%></font></td></tr>
	<tr align="right"><td><font size="1">Sales Tax(T&D):</font></td><td><font size="1"><%=formatcurrency(cDbl(salestaxtd),2)%></font></td></tr>
	<tr align="right"><td><font size="1"><b>Total Charges:</b></font></td><td><font size="1"><b><%=formatcurrency(cDbl(newtotalamount))%></b></font></td></tr>
	</table>



<%
	If Err then
	response.Write("<br>" & Err.Description & " " & rst1("tenantname"))
	response.End()
	end if
%>
	
	</td></tr></table>
	&nbsp;
	<%
	rst1.movenext
	pagepart = pagepart+1
	if pagepart>2 or metercount>11 then
		if not rst1.eof then response.write "<WxPrinter PageBreak>&nbsp;"
		pagepart = 1
	elseif trim(request("noheader"))<>"" then
		response.write "&nbsp;<br>&nbsp;<br>"
	end if
	'# Added by Tarun 07/10/2006
		lngTenantCount = lngTenantCount + 1
        
	'# 
loop
rst1.close

select case utilityid
case 1 'steam
	sql = "SELECT SalesTax, (case when isnull(taxincluded,0)=1 then TotalBillAmt-isnull(SalesTax,0) else TotalBillAmt end) as subtotal, isnull(MLbUsage,0) as MLbUsage, (case when isnull(taxincluded,0)=0 then isnull(TotalBillAmt,0)+isnull(SalesTax,0) else isnull(TotalBillAmt,0) end) as TotalBillAmt, isnull(AvgCost,0) as AvgCost FROM Utilitybill_steam u, billyrperiod bp WHERE u.ypid=bp.ypid and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
case 3 'water
	sql = "SELECT salestax, TotalBillAmt-salestax as subtotal, isnull(totalccf,0) as totalccf, isnull(watercharge,0) as watercharge, isnull(SewerCharge,0) as SewerCharge, isnull(TotalbillAmt,0) as TotalbillAmt, isnull(avgcost,0) as avgcost FROM Utilitybill_coldwater u, billyrperiod bp WHERE u.ypid=bp.ypid and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
case 4 'gas
	sql = "SELECT salestax, (case when isnull(taxincluded,0)=1 then TotalBillAmt-isnull(SalesTax,0) else TotalBillAmt end) as subtotal, isnull(ThermUsage,0) as ThermUsage, isnull(ccfUsage,0) as ccfUsage, (case when isnull(taxincluded,0)=0 then isnull(TotalBillAmt,0)+isnull(SalesTax,0) else isnull(TotalBillAmt,0) end) as TotalBillAmt, isnull(avgcosttherm,0) as avgcosttherm FROM Utilitybill_gas u, billyrperiod bp WHERE u.ypid=bp.ypid and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
case else 'electricity?
	sql = "select distinct ub.salestax, TotalBillAmt-ub.intot as subtotal, OnPeakKWH as OnPeakKWH,OffPeakKWH as OffPeakKWH, TotalKWH as TotalKWH, CostKWH as CostKWH,(case when TotalKWH=0 then 0 else CostKWH/TotalKWH end) as UnitCostKWH, TotalKW as TotalKW,CostKW as CostKW, (case when TotalKW=0 then 0 else CostKW/TotalKW end) as UnitCostKW, isnull(TotalBillAmt,0)+extot as TotalBillAmt,(case when [totalkw]*24*DateDiff(day,[ypiddatestart],[ypidDateEnd])=0 then 0 else [totalkwh]/([totalkw]*24*(DateDiff(day,[ypiddatestart],[ypidDateEnd])+1)) end) as loadfactor from "&rpt&" ub  WHERE ub.reject=0 and ub.bldgnum='"&building&"' and ub.billyear="&byear&" and ub.billperiod="&bperiod&" and ub.utility="&utilityid
end select
'response.write sql
'response.end
rst1.open sql, cnn1
%>
<WxPrinter PageBreak>
<table cellspacing="1" cellpadding="2" border="0" width="100%">
<tr><td colspan="5" bgcolor="#3399CC"><font size="2" color="#FFFFFF"><b>Building Totals</b></font></td></tr>
<tr><td colspan="5"><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p></td></tr>
<tr><td align="center" valign="top">
<%if not rst1.eof then%>
	<%if not pid = 108 then%>
	<table border="0" cellspacing="2" cellpadding="5" <%if not utilityid=6 then%>bgcolor="black"<%end if%> width="250"><tr bgcolor="white"><td>
	<%if not utilityid=6 then%>
		<div align="center">Utility Expenses</div>
		<table width="100%">
		<%select case utilityid%>
		<%case 1%>
			<tr><td><font size="1"><%=usage%> Used</font></td><td><font size="1"><%=formatnumber(rst1("MLbUsage"),0)%></font></td></tr>
			<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%=formatcurrency(rst1("subtotal"),2)%></font></td></tr>
			<tr><td><font size="1">Sales Tax</font></td><td><font size="1"><%=formatcurrency(rst1("salestax"),2)%></font></td></tr>
			<tr><td><font size="1">Utility Bill Amount</font></td><td><font size="1"><%=formatcurrency(rst1("TotalBillAmt"),2)%></font></td></tr>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<tr><td><font size="1">Average Cost per <%=usage%></font></td><td><font size="1"><%=formatcurrency(rst1("AvgCost"),2)%></font></td></tr>
		<%case 3%>
			<tr><td><font size="1">CCF Used</font></td><td><font size="1"><%=formatnumber(rst1("totalccf"),0)%></font></td></tr>
			<tr><td><font size="1">Water Charge</font></td><td><font size="1"><%=formatcurrency(rst1("watercharge"),2)%></font></td></tr>
			<tr><td><font size="1">Sewer Charge</font></td><td><font size="1"><%=formatcurrency(rst1("sewercharge"),2)%></font></td></tr>
			<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%=formatcurrency(rst1("subtotal"),2)%></font></td></tr>
			<tr><td><font size="1">Sales Tax</font></td><td><font size="1"><%=formatcurrency(rst1("salestax"),2)%></font></td></tr>
			<tr><td><font size="1">Utility Bill Amount</font></td><td><font size="1"><%=formatcurrency(rst1("TotalBillAmt"),2)%></font></td></tr>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<tr><td><font size="1">Average Cost per CCF</font></td><td><font size="1"><%=formatcurrency(rst1("avgcost"),2)%></font></td></tr>
		<%case 4%>
			<tr><td><font size="1">CCF Used</font></td><td><font size="1"><%=formatnumber(rst1("ccfUsage"),0)%></font></td></tr>
			<tr><td><font size="1">Therms Used</font></td><td><font size="1"><%=formatnumber(rst1("ThermUsage"),0)%></font></td></tr>
			<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%=formatcurrency(rst1("subtotal"),2)%></font></td></tr>
			<tr><td><font size="1">Sales Tax</font></td><td><font size="1"><%=formatcurrency(rst1("salestax"),2)%></font></td></tr>
			<tr><td><font size="1">Utility Bill Amount</font></td><td><font size="1"><%=formatcurrency(rst1("TotalBillAmt"),2)%></font></td></tr>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<tr><td><font size="1">Average Cost per Therm</font></td><td><font size="1"><%=formatcurrency(rst1("avgcosttherm"),2)%></font></td></tr>
		<%case 6%>
			<!-- <tr><td><font size="1">Utility Bill Amount</font></td><td><font size="1"><%=formatcurrency(rst1("TotalBillAmt"),2)%></font></td></tr> -->
           
		<%case else%>
			<tr><td><font size="1">On Peak <%=usage%></font></td><td><font size="1"><%if isnumeric(rst1("OnPeakKWH")) then response.write formatnumber(rst1("OnPeakKWH"),0)%></font></td></tr>
			<tr><td><font size="1">Off Peak <%=usage%></font></td><td><font size="1"><%if isnumeric(rst1("OffPeakKWH")) then response.write formatnumber(rst1("OffPeakKWH"),0)%></font></td></tr>
			<tr><td><font size="1">Total <%=usage%></font></td><td><font size="1"><%if isnumeric(rst1("TotalKWH")) then response.write formatnumber(rst1("TotalKWH"),0)%></font></td></tr>
			<tr><td><font size="1">Cost <%=usage%></font></td><td><font size="1"><%if isnumeric(rst1("CostKWH")) then response.write formatcurrency(rst1("CostKWH"))%></font></td></tr>
			<tr><td><font size="1">Unit Cost <%=usage%></font></td><td><font size="1"><%if isnumeric(rst1("UnitCostKWH")) then response.write formatcurrency(rst1("UnitCostKWH"),4)%></font></td></tr>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<!--<tr><td><font size="1">Total <%=demand%></font></td><td><font size="1"><%if isnumeric(rst1("TotalKW")) then response.write formatnumber(rst1("TotalKW"),0)%></font></td></tr>-->
			<!--reformatted to display two places after decimal on bill summary total kw.  Michelle T. 5/4/2009--> 
			<tr><td><font size="1">Total <%=demand%></font></td><td><font size="1"><%if isnumeric(rst1("TotalKW")) then response.write formatnumber(rst1("TotalKW"),2)%></font></td></tr>
			
			<tr><td><font size="1">Cost <%=demand%></font></td><td><font size="1"><%if isnumeric(rst1("CostKW")) then response.write formatcurrency(rst1("CostKW"))%></font></td></tr>
			<tr><td><font size="1">Unit Cost <%=demand%></font></td><td><font size="1"><%if isnumeric(rst1("UnitCostKW")) then response.write formatcurrency(rst1("UnitCostKW"),2)%></font></td></tr>
			<tr><td><font size="1">Load Factor</font></td><td><font size="1"><%if isnumeric(rst1("loadfactor")) then response.write formatpercent(rst1("loadfactor"),2)%></font></td></tr>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<!-- <tr><td colspan="2">&nbsp;</td></tr> -->
			<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%if isnumeric(rst1("subtotal")) then response.write formatcurrency(rst1("subtotal"),2)%></font></td></tr>
			<tr><td><font size="1">Sales Tax</font></td><td><font size="1"><%if isnumeric(rst1("salestax")) then response.write formatcurrency(rst1("salestax"),2)%></font></td></tr>
			<tr><td><font size="1">Utility Bill Amount</font></td><td><font size="1"><%if isnumeric(rst1("TotalBillAmt")) then response.write formatcurrency(rst1("TotalBillAmt"),2)%></font></td></tr>
			<tr><td><font size="1">Average Cost</font></td><td><font size="1"><%if isnumeric(rst1("TotalBillAmt")) and clng(rst1("TotalKWH"))<>0 then response.write formatcurrency(rst1("TotalBillAmt")/rst1("TotalKWH"),6)%></font></td></tr>
		<%end select%>
		</table>
		
	<%end if%>
</td></tr></table>
<%end if%>
<%end if%>
</td><td>&nbsp;</td><td align="center" valign="top">
	<table border="0" cellspacing="2" cellpadding="5" bgcolor="black" width="250"><tr bgcolor="white"><td><div align="center">Sub-Meter Revenue</div>
		<table width="100%">
  		 <%
			
			'dim ttlOnPeak,ttlIntPeak,ttlOffPeak,comboOnPeak,comboIntPeak,comboOffPeak,comboTotalPeak,comboTotalKWon,comboTotalKWint,comboTotalKWoff, ttlTotalKWon,ttlTotalKWint,ttlTotalKWoff,ttlTotalPeak
			'comboOnPeak = PAbldgOnPeak+MAbldgOnPeak
			'comboIntPeak = PAbldgIntPeak+MAbldgIntPeak
			'comboOffPeak = PAbldgOffPeak+MAbldgOffPeak
			'comboTotalPeak = PAbldgTotalPeak+MAbldgTotalPeak        
			'comboTotalKWint= PAbldgTotalKWint+MAbldgTotalKWint 
    		'comboTotalKWon = PAbldgTotalKWon+MAbldgTotalKWon
			'comboTotalKWint = PAbldgTotalKWint+MAbldgTotalKWint
			'comboTotalKWoff = PAbldgTotalKWoff+MAbldgTotalKWoff

			'if bldgOnPeak > comboOnPeak then
			'ttlOnPeak = bldgOnPeak - comboOnPeak
			'else
			'ttlOnPeak = comboOnPeak-bldgOnPeak  
			'end if

			'if bldgIntPeak > comboIntPeak then
			'ttlIntPeak = bldgIntPeak - comboIntPeak
			'else
			'ttlIntPeak = comboIntPeak-bldgIntPeak  
			'end if
			
			'if bldgOffPeak > comboOffPeak then
			'ttlOffPeak = bldgOffPeak - comboOffPeak
			'else
			'ttlOffPeak = comboOffPeak-bldgOffPeak  
			'end if
			
			'if bldgTotalPeak > comboTotalPeak then
			'ttlTotalPeak = bldgTotalPeak - comboTotalPeak
			'else
            'ttlTotalPeak=comboTotalPeak-bldgTotalPeak
			'end if
						
			'if bldgTotalKWon > comboTotalKWon then
			'ttlTotalKWon = bldgTotalKWon - comboTotalKWon
			'else
           ' ttlTotalKWon=comboTotalPeak-bldgTotalKWon
			'end if

			'if bldgTotalKWint > comboTotalKWint then
			'ttlTotalKWint = bldgTotalKWint - comboTotalKWint
			'else
            'ttlTotalKWint = comboTotalKWint-bldgTotalKWint
			'end if

			'if bldgTotalKWoff > comboTotalKWoff then
			'ttlTotalKWoff = bldgTotalKWoff - comboTotalKWoff
			'else
           ' ttlTotalKWoff=comboTotalKWoff-bldgTotalKWoff
			'end if



		%>	


		<%if utilityid=2 then'on/off break downs%>
			<tr><td><font size="1">On Peak</font></td><td><font size="1"><%if isnumeric(bldgOnPeak) then response.write formatnumber(bldgOnPeak,2)%></font></td></tr>
			<%if calcintpeak OR extusageflag then %>
			<tr><td><font size="1">Int Peak</font></td><td><font size="1"><%if isnumeric(bldgIntPeak) then response.write formatnumber(bldgIntPeak,2)%></font></td></tr>
			<%end if%>
			<tr><td><font size="1">Off Peak</font></td><td><font size="1"><%if isnumeric(bldgOffPeak) then response.write formatnumber(bldgOffPeak,2)%></font></td></tr>
		<%end if%>
		<tr><td><font size="1">Total <%if utilityid=3 or utilityid=4 then%>C<%end if%><%=usage%></font></td><td><font size="1"><%if isnumeric(bldgTotalPeak) then response.write formatnumber(bldgTotalPeak,2)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<%if utilityid=2 then'demand total'or utilityid=6 %>
			<%if calcintpeak OR extusageflag then %>
				<tr><td><font size="1"><%=demand%> On</font></td><td><font size="1"><%if isnumeric(bldgTotalKWon) then response.write formatnumber(bldgTotalKWon,0)%></font></td></tr>
				<tr><td><font size="1"><%=demand%> Int</font></td><td><font size="1"><%if isnumeric(bldgTotalKWint) then response.write formatnumber(bldgTotalKWint,0)%></font></td></tr>
				<tr><td><font size="1"><%=demand%> Off</font></td><td><font size="1"><%if isnumeric(bldgTotalKWoff) then response.write formatnumber(bldgTotalKWoff,0)%></font></td></tr>
			<%end if
			'dim combototal,ttl,ttlfinal
			'ttl = bldgTotalKWon+bldgTotalKWoff+bldgTotalKWint 'reg ttl
			'combototal=PAbldgTotalKWon+PAbldgTotalKWoff+PAbldgTotalKWint + MAbldgTotalKWon+MAbldgTotalKWoff+MAbldgTotalKWint 'pa ttl plus MA ttl
			'if ttl > combototal then 
			'ttlfinal = ttl-combototal
			'else 
			'ttlfinal =  combototal - ttl
			'end if
			%>
			<tr><td><font size="1">Total <%=demand%></font></td><td><font size="1"><%if isnumeric(bldgTotalKWon+bldgTotalKWoff+bldgTotalKWint) then response.write formatnumber(bldgTotalKWon+bldgTotalKWoff+bldgTotalKWint,0)%></font></td></tr>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<%end if%>
		<%


        'RSM
		%>
		<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%if isnumeric(subsubtotal) then response.write formatcurrency(subsubtotal)%></font></td></tr>
		<tr><td><font size="1">Admin Fee</font></td><td><font size="1"><%if isnumeric(bldgAdmin) then response.write formatcurrency(bldgAdmin)%></font></td></tr>
		<tr><td><font size="1">Service Fee</font></td><td><font size="1"><%if isnumeric(bldgService) then response.write formatcurrency(bldgService)%></font></td></tr>
		<tr><td><font size="1">Credit</font></td><td><font size="1"><%if isnumeric(bldgCredit) then response.write formatcurrency(bldgCredit)%></font></td></tr>
		<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%if isnumeric(bldgSubtotal) then response.write formatcurrency(bldgSubtotal)%></font></td></tr>
		<tr><td><font size="1">Tax(Supply)</font></td><td><font size="1"><%if isnumeric(bldgtotaltaxsupply) then response.write formatcurrency(bldgtotaltaxsupply,2)%></font></td></tr>
		<tr><td><font size="1">Tax(T&D)</font></td><td><font size="1"><%if isnumeric(bldgtotaltaxtd) then response.write formatcurrency(bldgtotaltaxtd,2)%></font></td></tr>
		<tr><td><font size="1">Total</font></td><td><font size="1"><%if isnumeric(bldgnewtotalamount) then response.write formatcurrency(bldgnewtotalamount,2)%></font></td></tr>
		<%if utilityid<>10 and utilityid<>6 and not(rst1.eof) then%>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<% if pid = "108" then
			    dim decide,bothPAnMA ' to eliminate negative  count on regular sub meter bill
				bothPAnMA = PAlngTenantCount + MAlngTenantCount + AMlngTenantCount
				if bothPAnMA > lngTenantCount then 
					decide = bothPAnMA-lngTenantCount
				else
					decide = lngTenantCount-bothPAnMA
				end if 
			%>
			<tr><td><font size="1">Total Bill Count</font></td><td><font size="1"><% response.Write FormatNumber(decide,0)%></font></td></tr>		
			<% else %>
			<tr><td><font size="1">% Recoup1</font></td><td><font size="1"><%if isnumeric(bldgSubtotal) and trim(rst1("TotalBillAmt"))<>"0" then response.write formatpercent(bldgSubtotal/cdbl(rst1("TotalBillAmt")),2)%></font></td></tr><!-- a bit iffy test this out-->
				<%if utilityid=2 then 
					if ExmpData then %>
						<tr><td><font size="1">% Recoup (KWH)</font></td><td><font size="1">
							<%if isnumeric(bldgTotalPeak) and trim(rst1("TotalKWH"))<>"0" then 
									if  isnumeric(ExmpTotalPeak) then
										response.write formatpercent((bldgTotalPeak +ExmpTotalPeak) /cdbl(rst1("TotalKWH")),2) 
									else
										response.write formatpercent(bldgTotalPeak  /cdbl(rst1("TotalKWH")),2) 
									end if
							end if%>
						</font></td></tr>				
					<%else%>
						<tr><td><font size="1">% Recoup (KWH)</font></td><td><font size="1"><%if isnumeric(bldgTotalPeak) and trim(rst1("TotalKWH"))<>"0" then response.write formatpercent(bldgTotalPeak/cdbl(rst1("TotalKWH")),2)%></font></td></tr>
					<%end If%>
				<% end if%>
			<% end if %>
		<%end if%>
		</table>
	</td></tr></table>
   <% if pid = "108" then 'find another condition to make this work
		if PAFlag then %>
  <td align="center" valign="top">
<table border="0" cellspacing="2" cellpadding="5" bgcolor="black" width="250"><tr bgcolor="white"><td><div align="center">PA Sub-Meters </div>
		<table width="100%">
		<%if utilityid=2 then'on/off break downs%>
			<tr><td><font size="1">On Peak</font></td><td><font size="1"><%if isnumeric(PAbldgOnPeak) then response.write formatnumber(PAbldgOnPeak,2)%></font></td></tr>
			<%if calcintpeak OR extusageflag then %>
			<tr><td><font size="1">Int Peak</font></td><td><font size="1"><%if isnumeric(PAbldgIntPeak) then response.write formatnumber(PAbldgIntPeak,2)%></font></td></tr>
			<%end if%>
			<tr><td><font size="1">Off Peak</font></td><td><font size="1"><%if isnumeric(PAbldgOffPeak) then response.write formatnumber(PAbldgOffPeak,2)%></font></td></tr>
		<%end if%>
		<tr><td><font size="1">Total <%if utilityid=3 or utilityid=4 then%>C<%end if%><%=usage%></font></td><td><font size="1"><%if isnumeric(PAbldgTotalPeak) then response.write formatnumber(PAbldgTotalPeak,2)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<%if utilityid=2 then'demand total'or utilityid=6 %>
			<%if calcintpeak OR extusageflag then %>
				<tr><td><font size="1"><%=demand%> On</font></td><td><font size="1"><%if isnumeric(PAbldgTotalKWon) then response.write formatnumber(PAbldgTotalKWon,0)%></font></td></tr>
				<tr><td><font size="1"><%=demand%> Int</font></td><td><font size="1"><%if isnumeric(PAbldgTotalKWint) then response.write formatnumber(PAbldgTotalKWint,0)%></font></td></tr>
				<tr><td><font size="1"><%=demand%> Off</font></td><td><font size="1"><%if isnumeric(PAbldgTotalKWoff) then response.write formatnumber(PAbldgTotalKWoff,0)%></font></td></tr>
			<%end if
			'dim PAcombototal,MAcombototal,MAfinal
			'PAcombototal=PAbldgTotalKWon+PAbldgTotalKWoff+PAbldgTotalKWint
			'MAcombototal=MAbldgTotalKWon+MAbldgTotalKWoff+MAbldgTotalKWint
			
			'if PAcombototal > MAcombototal then 
			'MAfinal = PAcombototal-MAcombototal
			'else 
			'MAfinal = MAcombototal - PAcombototal
			'end if
			
			%>
			<tr><td><font size="1">Total <%=demand%></font></td><td><font size="1"><%if isnumeric(PAbldgTotalKWon+PAbldgTotalKWoff+PAbldgTotalKWint) then response.write formatnumber(PAbldgTotalKWon+PAbldgTotalKWoff+PAbldgTotalKWint,0)%></font></td></tr>
			<!--<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>-->
		<%end if%>
		<!--<tr><td><font size="1">Subtotal</font></td><td><font size="1"><'%if isnumeric(PAsubsubtotal) then response.write formatcurrency(PAsubsubtotal)%></font></td></tr>
		<tr><td><font size="1">Admin Fee</font></td><td><font size="1"><'%if isnumeric(PAbldgAdmin) then response.write formatcurrency(PAbldgAdmin)%></font></td></tr>
		<tr><td><font size="1">Service Fee</font></td><td><font size="1"><'%if isnumeric(PAbldgService) then response.write formatcurrency(PAbldgService)%></font></td></tr>
		<tr><td><font size="1">Credit</font></td><td><font size="1"><'%if isnumeric(PAbldgCredit) then response.write formatcurrency(PAbldgCredit)%></font></td></tr>
		<tr><td><font size="1">Subtotal</font></td><td><font size="1"><'%if isnumeric(PAbldgSubtotal) then response.write formatcurrency(PAbldgSubtotal)%></font></td></tr>
		<tr><td><font size="1">Tax</font></td><td><font size="1"><'%if isnumeric(PAbldgTax) then response.write formatcurrency(PAbldgTax,2)%></font></td></tr>
		<tr><td><font size="1">Total</font></td><td><font size="1"><'%if isnumeric(PAbldgTotalAmt) then response.write formatcurrency(PAbldgTotalAmt,2)%></font></td></tr>
		--><%if utilityid<>10 and utilityid<>6 and not(rst1.eof) then%>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<% if pid = "108" then %>
			<tr><td><font size="1">Total Bill Count</font></td><td><font size="1"><% response.Write FormatNumber(PAlngTenantCount,0)%></font></td></tr>		
			<% else %>
			<tr><td><font size="1">% Recoup</font></td><td><font size="1"><%if isnumeric(PAbldgTotalAmt) and trim(rst1("TotalBillAmt"))<>"0" then response.write formatpercent(PAbldgTotalAmt/cdbl(rst1("TotalBillAmt")),2)%></font></td></tr>
			<% end if %>
		<%end if%>
		</table>
	</td></tr></table>
	 	
 </td>
	<%end if
	if MAflag Then
	%>
	
<td align="center" valign="top">
		<table border="0" cellspacing="2" cellpadding="5" bgcolor="black" width="250">
		<tr bgcolor="white">
		<td>
		<div align="center">PA Master Meters</div>
		<table width="100%">
		<%if utilityid=2 then'on/off break downs%>
			<tr><td><font size="1">On Peak</font></td><td><font size="1"><%if isnumeric(MAbldgOnPeak) then response.write formatnumber(MAbldgOnPeak,2)%></font></td></tr>
			<%if calcintpeak OR extusageflag then %>
			<tr><td><font size="1">Int Peak</font></td><td><font size="1"><%if isnumeric(MAbldgIntPeak) then response.write formatnumber(MAbldgIntPeak,2)%></font></td></tr>
			<%end if%>
			<tr><td><font size="1">Off Peak</font></td><td><font size="1"><%if isnumeric(MAbldgOffPeak) then response.write formatnumber(MAbldgOffPeak,2)%></font></td></tr>
		<%end if%>
		<tr><td><font size="1">Total <%if utilityid=3 or utilityid=4 then%>C<%end if%><%=usage%></font></td><td><font size="1"><%if isnumeric(MAbldgTotalPeak) then response.write formatnumber(MAbldgTotalPeak,2)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<%if utilityid=2 then'demand total'or utilityid=6 %>
			<%if calcintpeak OR extusageflag then %>
				<tr><td><font size="1"><%=demand%> On</font></td><td><font size="1"><%if isnumeric(MAbldgTotalKWon) then response.write formatnumber(MAbldgTotalKWon,0)%></font></td></tr>
				<tr><td><font size="1"><%=demand%> Int</font></td><td><font size="1"><%if isnumeric(MAbldgTotalKWint) then response.write formatnumber(MAbldgTotalKWint,0)%></font></td></tr>
				<tr><td><font size="1"><%=demand%> Off</font></td><td><font size="1"><%if isnumeric(MAbldgTotalKWoff) then response.write formatnumber(MAbldgTotalKWoff,0)%></font></td></tr>
			<%end if
			dim MAcombototal2
			MAcombototal2=MAbldgTotalKWon+MAbldgTotalKWoff+MAbldgTotalKWint
			%>
			<tr><td><font size="1">Total <%=demand%></font></td><td><font size="1"><%if isnumeric(MAbldgTotalKWon+MAbldgTotalKWoff+MAbldgTotalKWint) then response.write formatnumber(MAcombototal2,0)%></font></td></tr>
			<!--<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>-->
		<%end if%>
		<!--<tr><td><font size="1">Subtotal</font></td><td><font size="1"><'%if isnumeric(MAsubsubtotal) then response.write formatcurrency(MAsubsubtotal)%></font></td></tr>
		<tr><td><font size="1">Admin Fee</font></td><td><font size="1"><'%if isnumeric(MAbldgAdmin) then response.write formatcurrency(MAbldgAdmin)%></font></td></tr>
		<tr><td><font size="1">Service Fee</font></td><td><font size="1"><'%if isnumeric(MAbldgService) then response.write formatcurrency(MAbldgService)%></font></td></tr>
		<tr><td><font size="1">Credit</font></td><td><font size="1"><'%if isnumeric(MAbldgCredit) then response.write formatcurrency(MAbldgCredit)%></font></td></tr>
		<tr><td><font size="1">Subtotal</font></td><td><font size="1"><'%if isnumeric(MAbldgSubtotal) then response.write formatcurrency(MAbldgSubtotal)%></font></td></tr>
		<tr><td><font size="1">Tax</font></td><td><font size="1"><'%if isnumeric(MAbldgTax) then response.write formatcurrency(MAbldgTax,2)%></font></td></tr>
		<tr><td><font size="1">Total</font></td><td><font size="1"><'%if isnumeric(MAbldgTotalAmt) then response.write formatcurrency(MAbldgTotalAmt,2)%></font></td></tr>
		--><%if utilityid<>10 and utilityid<>6 and not(rst1.eof) then%>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<% if pid = "108" then %>
			<tr><td><font size="1">Total Bill Count</font></td><td><font size="1"><% response.Write FormatNumber(MAlngTenantCount,0)%></font></td></tr>		
			<% else %>
			<tr><td><font size="1">% Recoup</font></td><td><font size="1"><%if isnumeric(MAbldgTotalAmt) and trim(rst1("TotalBillAmt"))<>"0" then response.write formatpercent(MAbldgTotalAmt/cdbl(rst1("TotalBillAmt")),2)%></font></td></tr>
			<% end if %> 
		<%end if%>
		</table>
		</td>
		</tr></table>
		<% end if 
			if AMFlag then
		%>
		<td align="center" valign="top">
		<table border="0" cellspacing="2" cellpadding="5" bgcolor="black" width="250" ID="Table1">
		<tr bgcolor="white">
		<td>
		<div align="center">PA Audit Meters</div>
		<table width="100%" ID="Table2">
		<%if utilityid=2 then'on/off break downs%>
			<tr><td><font size="1">On Peak</font></td><td><font size="1"><%if isnumeric(AMbldgOnPeak) then response.write formatnumber(AMbldgOnPeak,2)%></font></td></tr>
			<%if calcintpeak OR extusageflag then %>
			<tr><td><font size="1">Int Peak</font></td><td><font size="1"><%if isnumeric(AMbldgIntPeak) then response.write formatnumber(AMbldgIntPeak,2)%></font></td></tr>
			<%end if%>
			<tr><td><font size="1">Off Peak</font></td><td><font size="1"><%if isnumeric(AMbldgOffPeak) then response.write formatnumber(AMbldgOffPeak,2)%></font></td></tr>
		<%end if%>
		<tr><td><font size="1">Total <%if utilityid=3 or utilityid=4 then%>C<%end if%><%=usage%></font></td><td><font size="1"><%if isnumeric(AMbldgTotalPeak) then response.write formatnumber(AMbldgTotalPeak,2)%></font></td></tr>
		<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
		<%if utilityid=2 then'demand total'or utilityid=6 %>
			<%if calcintpeak OR extusageflag then %>
				<tr><td><font size="1"><%=demand%> On</font></td><td><font size="1"><%if isnumeric(AMbldgTotalKWon) then response.write formatnumber(AMbldgTotalKWon,0)%></font></td></tr>
				<tr><td><font size="1"><%=demand%> Int</font></td><td><font size="1"><%if isnumeric(AMbldgTotalKWint) then response.write formatnumber(AMbldgTotalKWint,0)%></font></td></tr>
				<tr><td><font size="1"><%=demand%> Off</font></td><td><font size="1"><%if isnumeric(AMbldgTotalKWoff) then response.write formatnumber(AMbldgTotalKWoff,0)%></font></td></tr>
			<%end if
			dim AMcombototal2
			AMcombototal2=AMbldgTotalKWon+AMbldgTotalKWoff+AMbldgTotalKWint
			%>
			<tr><td><font size="1">Total <%=demand%></font></td><td><font size="1"><%if isnumeric(AMbldgTotalKWon+AMbldgTotalKWoff+AMbldgTotalKWint) then response.write formatnumber(AMcombototal2,0)%></font></td></tr>
			<!--<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>-->
		<%end if%>
		<!--<tr><td><font size="1">Subtotal</font></td><td><font size="1"><'%if isnumeric(MAsubsubtotal) then response.write formatcurrency(MAsubsubtotal)%></font></td></tr>
		<tr><td><font size="1">Admin Fee</font></td><td><font size="1"><'%if isnumeric(MAbldgAdmin) then response.write formatcurrency(MAbldgAdmin)%></font></td></tr>
		<tr><td><font size="1">Service Fee</font></td><td><font size="1"><'%if isnumeric(MAbldgService) then response.write formatcurrency(MAbldgService)%></font></td></tr>
		<tr><td><font size="1">Credit</font></td><td><font size="1"><'%if isnumeric(MAbldgCredit) then response.write formatcurrency(MAbldgCredit)%></font></td></tr>
		<tr><td><font size="1">Subtotal</font></td><td><font size="1"><'%if isnumeric(MAbldgSubtotal) then response.write formatcurrency(MAbldgSubtotal)%></font></td></tr>
		<tr><td><font size="1">Tax</font></td><td><font size="1"><'%if isnumeric(MAbldgTax) then response.write formatcurrency(MAbldgTax,2)%></font></td></tr>
		<tr><td><font size="1">Total</font></td><td><font size="1"><'%if isnumeric(MAbldgTotalAmt) then response.write formatcurrency(MAbldgTotalAmt,2)%></font></td></tr>
		--><%if utilityid<>10 and utilityid<>6 and not(rst1.eof) then%>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<% if pid = "108" then %>
			<tr><td><font size="1">Total Bill Count</font></td><td><font size="1"><% response.Write FormatNumber(AMlngTenantCount,0)%></font></td></tr>		
			<% else %>
			<tr><td><font size="1">% Recoup</font></td><td><font size="1"><%if isnumeric(AMbldgTotalAmt) and trim(rst1("TotalBillAmt"))<>"0" then response.write formatpercent(AMbldgTotalAmt/cdbl(rst1("TotalBillAmt")),2)%></font></td></tr>
			<% end if %>
		<%end if%>
		</table>
		
		</td></tr>
		</table>

</td>
	
   <%end if
   End If
   %>
	<%if ExmpData then%>
		</td><td>&nbsp;</td><td align="center" valign="top">
		<table border="0" cellspacing="2" cellpadding="5" bgcolor="black" width="250"><tr bgcolor="white"><td><div align="center">Revenue Exempt Summary</div>
			<table width="100%">
			<%if utilityid=2 then'on/off break downs%>
				<tr><td><font size="1">On Peak</font></td><td><font size="1"><%if isnumeric(ExmpOnPeak) then response.write formatnumber(ExmpOnPeak,2)%></font></td></tr>
				<%if calcintpeak OR extusageflag then %>
				<tr><td><font size="1">Int Peak</font></td><td><font size="1"><%if isnumeric(ExmpIntPeak) then response.write formatnumber(ExmpIntPeak,2)%></font></td></tr>
				<%end if%>
				<tr><td><font size="1">Off Peak</font></td><td><font size="1"><%if isnumeric(ExmpOffPeak) then response.write formatnumber(ExmpOffPeak,2)%></font></td></tr>
			<%end if%>
			<tr><td><font size="1">Total <%if utilityid=3 then%>C<%end if%><%=usage%></font></td><td><font size="1"><%if isnumeric(ExmpTotalPeak) then response.write formatnumber(ExmpTotalPeak,2)%></font></td></tr>
			<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<%if utilityid=2 then'demand total'or utilityid=6 %>
				<%if calcintpeak OR extusageflag then %>
					<tr><td><font size="1"><%=demand%> On</font></td><td><font size="1"><%if isnumeric(ExmpTotalKWon) then response.write formatnumber(ExmpTotalKWon,0)%></font></td></tr>
					<tr><td><font size="1"><%=demand%> Int</font></td><td><font size="1"><%if isnumeric(ExmpTotalKWint) then response.write formatnumber(ExmpTotalKWint,0)%></font></td></tr>
					<tr><td><font size="1"><%=demand%> Off</font></td><td><font size="1"><%if isnumeric(ExmpTotalKWoff) then response.write formatnumber(ExmpTotalKWoff,0)%></font></td></tr>
				<%end if%>
				<tr><td><font size="1">Total <%=demand%></font></td><td><font size="1"><%if isnumeric(ExmpTotalKW) then response.write formatnumber(ExmpTotalKW,0)%></font></td></tr>
				<tr><td colspan="2"><hr size="1" color="#000000" noshade></td></tr>
			<%end if%>
			<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%if isnumeric(Exmpsubsubtotal) then response.write formatcurrency(Exmpsubsubtotal)%></font></td></tr>
			<tr><td><font size="1">Admin Fee</font></td><td><font size="1"><%if isnumeric(ExmpAdmin) then response.write formatcurrency(ExmpAdmin)%></font></td></tr>
			<tr><td><font size="1">Service Fee</font></td><td><font size="1"><%if isnumeric(ExmpService) then response.write formatcurrency(ExmpService)%></font></td></tr>
			<tr><td><font size="1">Credit</font></td><td><font size="1"><%if isnumeric(ExmpCredit) then response.write formatcurrency(ExmpCredit)%></font></td></tr>
			<tr><td><font size="1">Subtotal</font></td><td><font size="1"><%if isnumeric(ExmpSubtotal) then response.write formatcurrency(ExmpSubtotal)%></font></td></tr>
			<tr><td><font size="1">Tax</font></td><td><font size="1"><%if isnumeric(ExmpTax) then response.write formatcurrency(ExmpTax,2)%></font></td></tr>
			<tr><td><font size="1">Total</font></td><td><font size="1"><%if isnumeric(ExmpTotalAmt) then response.write formatcurrency(ExmpTotalAmt,2)%></font></td></tr>
		</table>
	<%end if%>
	</td></tr></table>
</td></tr></table>
<%if pdf then%><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><%end if%>
*Tenants marked with an asterisk are accounted for in the revenue exempt summary.<br>
*Consumption and Demand marked with an asterisk have been estimated.

