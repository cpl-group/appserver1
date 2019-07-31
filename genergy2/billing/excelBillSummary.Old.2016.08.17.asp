<%@ Language=VBScript %>
<%option explicit%>

<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--  
    METADATA  
    TYPE="typelib"  
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library"  
--> 
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"-->
<%end if

	function getNumber(number)
	'	response.write "|"&number&"|"
		if not(isNumeric(number)) then number = 0
		getNumber = number
	end function
	
    Dim objExcelReport 
    Dim objWorkBook 
    Dim objWorkSheet 
    Dim objCell 

    Set objExcelReport = CreateObject("Excel.Application")
    Set objWorkBook = objExcelReport.Workbooks.Add

	Dim  Billperiod, building, Billyear, PortFolioId, UtilityId, rpt, pdf, Genergy_Users, demo, sql, email
	
	' Set Parameters
	
	Billperiod = request("bperiod")
	building = request("building")
	Billyear = request("byear")
	PortFolioId = request("pid")
	UtilityId = trim(request("utilityid"))
	email = request("email")
	
	Dim rst1, rst2, rst3, cnn1
	Dim usage, demand, utilityname

	set rst1 = server.createobject("ADODB.Recordset")
	set rst2 = server.createobject("ADODB.Recordset")
	set rst3 = server.createobject("ADODB.Recordset")
	set cnn1 = server.createobject("ADODB.Connection")


	cnn1.open getLocalConnect(building)
	rst1.open "SELECT umeasure as usage, dmeasure as demand, utilitydisplay as utility " & _
			  "FROM tblutility WHERE UtilityId="&utilityid, getConnect(PortFolioId,building,"Billing")
	
	' Get Display names 
	If not rst1.eof then 
		usage = rst1("usage")
		demand = rst1("demand")
		utilityname = rst1("utility")
	End if
	rst1.close
	
	' Get Billyear, BillPeriod in case input parameters are blank.
	If trim(Billperiod)="" or trim(Billyear)="" then
		rst1.open "select top 1 BillYear, BillPeriod from tblmetersbyperiod "& _
					"WHERE bldgnum='"&building&"' and utility="&UtilityId&" ORDER BY billyear desc, billperiod desc", cnn1
		If rst1.eof then
			response.write "No information for this building"
			response.end
		Else
			Billyear = cint(rst1("billyear"))
			Billperiod = cint(rst1("billperiod"))
		End if
		rst1.close
	End if	

	dim DBlocalIP
	if trim(building)<>"" then DBlocalIP = ""	
	' Get Logo image path
	Dim summarylogo
	rst1.open "SELECT logo FROM portfolio p, billtemplates bt " & _
				" WHERE p.templateid=bt.id and p.id=(SELECT portfolioid FROM "&DBlocalIP&"buildings WHERE bldgnum='"&building&"')", getConnect(PortFolioId,building,"Billing")

	If not rst1.eof then 
		summarylogo = rst1("logo")
	End If
	rst1.close
	
	If allowGroups("Genergy Users") Then 
		rpt = "rpt_bill_summary" 
	Else 
		rpt = "rpt_Bill_summary_client"
	End If
	
	sql = "SELECT isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.ibsexempt, " & _
			 " r.unit_credit, b.strt, r.sqft, r.adminfee, isnull(r.RateModify,0) as RateModify, " & _
			 " isnull(r.fuelAdj,0) as fuelAdjnum, rt.[type], AvgKWH, tenantname, datestart-1 as datestart, " & _
			 " datediff(day, datestart-1, dateend) as days, dateend, ypid, r.leaseutilityid, r.billingname, " & _
			 " r.tenantnum, isnull(r.adjustment,0)-isnull(r.credit,0) as credit, isnull(subtotal,0) as subtotal, " & _
			 " isnull(energy,0) as energy, isnull(demand,0) as demand, isnull(serviceFee,0) as serviceFee, " & _
			 " isnull(tax,0) as tax, isnull(totalamt,0) as totalamt, b.btstrt, lup.calcintpeak, " & _
			 " isnull(adminfeedollar,0) as admindollar, isnull(r.extusg,0) as extusg, r.rate_servicefee, r.shadow " & _
	  " FROM "&rpt&" r, dbo.ratetypes rt, buildings b, tblleasesutilityprices lup, tblleases l " & _
	  " WHERE r.reject=0 and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid " & _
				" and b.bldgnum=r.bldgnum AND r.[type]=rt.id and billyear="&Billyear& _
				" and billperiod="&Billperiod&" and r.bldgnum='"&building & _
				"' and l.billsummaryexempt = 0 and r.utility="&utilityid & _
	 " ORDER BY TenantName"	
	 
	 rst1.open sql, cnn1
	
	if rst1.eof then
		rst1.close
		
		if allowGroups("Genergy Users") then 
			rpt = "rpt_bill_summary_nobill" 
		else 
			rpt = "rpt_Bill_summary_nobill_client"
		End If
		
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
					" and b.bldgnum=r.bldgnum AND r.[type]=rt.id and billyear="&Billyear& _
					" and billperiod="&Billperiod&" and r.bldgnum='"&building&"' and l.billsummaryexempt = 0 " & _
					" and r.utility="&UtilityId&_
			" ORDER BY TenantName"
		rst1.open sql, cnn1
	end if	 
	
	' Select the First Worksheet
	Set objWorkSheet = objExcelReport.Application.Workbooks(1).Sheets(1)
    'objWorkSheet.Cells(1,5) = "Bill Summary"
     
	Dim bldgOnPeak, bldgOffPeak,bldgintpeak, bldgTotalPeak, bldgTotalKWon, bldgTotalKWoff, bldgTotalKWint, bldgAdmin 
	Dim	bldgService, bldgCredit, bldgSubtotal, bldgTax, bldgTotalAmt, ExmpOnPeak, ExmpIntPeak, ExmpOffPeak, ExmpTotalPeak 
	Dim	ExmpTotalKW, ExmpTotalKWon, ExmpTotalKWoff, ExmpTotalKWint, ExmpTotalAmt, ExmpData, ExmpAdmin, ExmpService 
	Dim	ExmpCredit, ExmpSubtotal, ExmpTax, Exmpsubsubtotal
	      
    Dim totaldemand_PC, totaldemand_PCint, totaldemand_PCoff, totalOnpeak, totalOffPeak, totalIntPeak, totalKWH, meterdemandtemp 
	Dim	subsubtotal, meterdemandtempint, meterdemandtempoff,totalkwhoff, totalkwhint, usagedivisor


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

	objWorkSheet.Cells.Font.Name = "Book Antiqua"
	objWorkSheet.Cells.Font.Size = 9
	
If not rst1.eof and trim(request("noheader"))="" then

	objWorkSheet.Cells(1,5).Font.Bold = True
	objWorkSheet.Cells(1,5) = rst1("Strt")
	
	objWorkSheet.Cells(2,5) = "Submetering Summary Report" 
	objWorkSheet.Cells(2,5).Font.Bold = True

	objWorkSheet.Cells(3,8) = "Bill Year"
	objWorkSheet.Cells(3,8).Font.Bold = True
	objWorkSheet.Cells(3,9) = "Bill Period"
	objWorkSheet.Cells(3,9).Font.Bold = True
	objWorkSheet.Cells(3,10) = "Utility"
	objWorkSheet.Cells(3,10).Font.Bold = True

	objWorkSheet.Cells(4,8) = BillYear 
	objWorkSheet.Cells(4,9) = BillPeriod
	objWorkSheet.Cells(4,10) = utilityname
End if

Dim pagepart, calcintpeak, sql2, sInvoiceNo, extusageflag

If rst1("calcintpeak")="True" Then 
	calcintpeak = true 
Else 
	calcintpeak = false
End if

extusageflag = false
pagepart = 1

Dim iRow , iCol
	iCol = 1
	iRow = 6 
Do Until rst1.eof
	totaldemand_PC 		= 0
	totaldemand_PCint 	= 0
	totaldemand_PCoff 	= 0
	totalOnpeak 		= 0
	totalOffPeak 		= 0
  	totalIntPeak 		= 0
	totalKWH 			= 0
	totalKWHoff 		= 0
	totalKWHint 		= 0
	extusage 			= rst1("extusg")
	
	if  extusage  then
		extusageflag = true
	end if

	sql = "SELECT * FROM tblmetersbyperiod m, tblbillbyperiod b " & _
			" WHERE b.id=m.bill_id and b.reject=0 and m.leaseutilityid="&cdbl(rst1("leaseutilityid"))& _
			" and m.ypid="&cdbl(rst1("ypid"))&" ORDER BY meternum"
	
	'response.write sql
	'response.end

	rst2.open sql, cnn1
	objWorkSheet.Cells(iRow,1).Font.Bold = True
	If rst1("ibsexempt")="True" Then
		objWorkSheet.Cells(iRow,1) = "* " & rst1("billingname") & "(" & rst1("TenantNum") & ")"
	Else
		objWorkSheet.Cells(iRow,1) = rst1("billingname") & "(" & rst1("TenantNum") & ")"
	End If

	iRow= iRow + 1
	objWorkSheet.Cells(iRow,12) = "From"
	objWorkSheet.Cells(iRow,12).Font.Bold = True
	objWorkSheet.Cells(iRow,13) = "To"
	objWorkSheet.Cells(iRow,13).Font.Bold = True
	objWorkSheet.Cells(iRow,14) = "No. Of Days"
	objWorkSheet.Cells(iRow,14).Font.Bold = True
	iRow= iRow + 1
	
	objWorkSheet.Cells(iRow,12) = FormatDateTime(rst1("DateStart"),2)
	objWorkSheet.Cells(iRow,13) = FormatDateTime(rst1("DateEnd"),2)
	objWorkSheet.Cells(iRow,14) = rst1("Days")	
	
	Dim iSubTotalRow 
	iSubTotalRow = iRow + 2
	
	If extusage then
		
		objWorkSheet.Cells(iRow,1) = "Tenant"
		objWorkSheet.Cells(iRow,1).Font.Bold = True
		
		objWorkSheet.Cells(iRow,3) = "Readings"
		objWorkSheet.Cells(iRow,3).Font.Bold = True
		
		objWorkSheet.Cells(iRow,6) = "Consumption"
		objWorkSheet.Cells(iRow,6).Font.Bold = True
		
		if utilityid=2 or utilityid=1 or utilityid=6 then
			objWorkSheet.Cells(iRow,9) = "Demand"
			objWorkSheet.Cells(iRow,9).Font.Bold = True
		End If	
		
		iRow= iRow + 1
		
		objWorkSheet.Cells(iRow,1) = "Meter  No."
		objWorkSheet.Cells(iRow,1).Font.Bold = True
		
		objWorkSheet.Cells(iRow,2) = "Multi."
		objWorkSheet.Cells(iRow,2).Font.Bold = True
		
		objWorkSheet.Cells(iRow,3) = "Previous"
		objWorkSheet.Cells(iRow,3).Font.Bold = True
		
		objWorkSheet.Cells(iRow,4) = "Current"
		objWorkSheet.Cells(iRow,4).Font.Bold = True
		if UtilityId = 3 Then
			objWorkSheet.Cells(iRow,5) = "Total Usage " & "C" & usage 
		Else
			objWorkSheet.Cells(iRow,5) = "Total Usage " &  usage 
		End If
		objWorkSheet.Cells(iRow,5).Font.Bold = True
		
		If utilityid=2 or utilityid=1 or utilityid=6 Then
			objWorkSheet.Cells(iRow,9) = demand 
		End If
		
	Else	
		objWorkSheet.Cells(iRow,iCol) = "Tenant"
		objWorkSheet.Cells(iRow,iCol).Font.Bold = True
		
		objWorkSheet.Cells(iRow,iCol + 2) = "Readings"
		objWorkSheet.Cells(iRow,iCol + 2).Font.Bold = True
		
		
		If calcintpeak then
			objWorkSheet.Cells(iRow,iCol + 6) = "Consumption"
			objWorkSheet.Cells(iRow,iCol+6).Font.Bold = True
		Else
			objWorkSheet.Cells(iRow,iCol  + 5) = "Consumption"
			objWorkSheet.Cells(iRow,iCol+5).Font.Bold = True
		End If
		
		if utilityid=2 or utilityid=1 or utilityid=6 then
			If calcintpeak then
				objWorkSheet.Cells(iRow,10) = "Demand"
				objWorkSheet.Cells(iRow,10).Font.Bold = True
			Else
				objWorkSheet.Cells(iRow,9) = "Demand"
				objWorkSheet.Cells(iRow,9).Font.Bold = True
			End If
		End If
		
		iRow= iRow + 1
		
		objWorkSheet.Cells(iRow,iCol) = "Meter  No."
		objWorkSheet.Cells(iRow,iCol).Font.Bold = True
		objWorkSheet.Cells(iRow,iCol + 1) = "Multi."
		objWorkSheet.Cells(iRow,iCol+1).Font.Bold = True
		objWorkSheet.Cells(iRow,iCol + 2) = "Previous"
		objWorkSheet.Cells(iRow,iCol+2).Font.Bold = True
		objWorkSheet.Cells(iRow,iCOl + 3) = "Current"
		objWorkSheet.Cells(iRow,iCol+3).Font.Bold = True
		objWorkSheet.Cells(iRow,iCol + 4) = "On Peak"
		objWorkSheet.Cells(iRow,iCol+4).Font.Bold = True
		If calcintpeak then
			objWorkSheet.Cells(iRow,iCol + 5) = "Int Peak"
			objWorkSheet.Cells(iRow,iCol+5).Font.Bold = True
			objWorkSheet.Cells(iRow,iCol + 6) = "Off Peak"
			objWorkSheet.Cells(iRow,iCol+6).Font.Bold = True
			If UtilityId = 3 Then
				objWorkSheet.Cells(iRow, icol + 7) = "Total Usage " & "C" & usage 
			Else
				objWorkSheet.Cells(iRow,iCol + 7) = "Total Usage " &  usage 
			End If
			objWorkSheet.Cells(iRow,iCol+7).Font.Bold = True
			if utilityid=2 or utilityid=1 or utilityid=6 then
				If calcintpeak then
					objWorkSheet.Cells(iRow,iCol + 8) = "On " & demand 
					objWorkSheet.Cells(iRow,iCol + 9) = "Off " & demand
					
					objWorkSheet.Cells(iRow,iCol + 10) = "Int " & demand 
					
				Else
					objWorkSheet.Cells(iRow,iCol + 8) = "Demand"
				End If
				
			End If	
		Else
			objWorkSheet.Cells(iRow,iCol + 5) = "Off Peak"
			objWorkSheet.Cells(iRow,iCol+5).Font.Bold = True
			If UtilityId = 3 Then
				objWorkSheet.Cells(iRow,iCol + 6) = "Total Usage " & "C" & usage 
			Else
				objWorkSheet.Cells(iRow,iCol + 6) = "Total Usage " &  usage 
			End If			
			objWorkSheet.Cells(iRow,iCol+6).Font.Bold = True
			if utilityid=2 or utilityid=1 or utilityid=6 then
				If calcintpeak then
					objWorkSheet.Cells(iRow,iCol + 8) = "On " & demand 
					objWorkSheet.Cells(iRow,iCol + 9) = "Off " & demand 
					objWorkSheet.Cells(iRow,iCol + 10) = "Int " & demand 
				Else
					objWorkSheet.Cells(iRow,iCol + 8) = demand 
				End If
			End If		
		End If
		objWorkSheet.Cells(iRow,iCol+8).Font.Bold = True
		objWorkSheet.Cells(iRow,iCol+9).Font.Bold = True 
		objWorkSheet.Cells(iRow,iCol+10).Font.Bold = True
	End If	
		
	Dim metercount, intpeak, extusage,tester, flag

 	
	extusage = rst1("extusg")
	metercount = 0
	'PAlngTenantCount=0
	do until rst2.eof
		metercount = metercount+1
		meterdemandtemp = rst2("Demand_P")
		meterdemandtempint = rst2("Demand_int")
		meterdemandtempoff = rst2("Demand_off")
	    tester = InStr(rst2("tenantnum"),"MA")
		
       'response.write PACondition
	   'response.end 
	   
		intpeak = 0
		if isnumeric(rst2("IntPeak")) then intpeak = cdbl(rst2("IntPeak"))
		if rst2("coincident")="True" then
			meterdemandtemp = 0
			totaldemand_PC = cdbl(rst2("Demand_C")) + cdbl(rst2("Demand_int")) + cdbl(rst2("Demand_off"))
		else
			totaldemand_PC = totaldemand_PC + cdbl(meterdemandtemp)
			if calcintpeak then
				totaldemand_PCoff = totaldemand_PCoff + cdbl(meterdemandtempoff)
				totaldemand_PCint = totaldemand_PCint + cdbl(meterdemandtempint)
			end if
		end if
		
   		totalOffPeak = totalOffPeak + formatnumber(rst2("OffPeak"),2)
		
		if calcintpeak then 
			totalIntPeak 	= totalIntPeak + formatnumber(rst2("IntPeak"),2)
			totalOnPeak 	= totalOnpeak + formatnumber(rst2("OnPeak"),2)
		else
			totalOnpeak = totalOnpeak + formatnumber(cdbl(rst2("OnPeak"))+IntPeak,2)
		end if
		
		if extusage then 
			if rst2("mextusg") then 
				
				totalKWH 	= totalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2)
				totalkwhoff = totalKWHoff + formatnumber(cdbl(rst2("Usedoff"))/usagedivisor,2)
				totalkwhint = totalKWHint + formatnumber(cdbl(rst2("Usedint"))/usagedivisor,2)
			else
				totalKWH 	= totalKWH + formatnumber(cdbl(rst2("onpeak"))/usagedivisor,2)
				totalkwhoff = totalKWHoff + formatnumber(cdbl(rst2("offpeak"))/usagedivisor,2)
				totalkwhint = totalKWHint + formatnumber(cdbl(rst2("intpeak"))/usagedivisor,2)
			end if
		else
			totalKWH 	= totalKWH + formatnumber(cdbl(rst2("used"))/usagedivisor,2) 
		end if 
		
		if extusage and rst2("mextusg") then 
			metercount = metercount + 2
			
			iRow = iRow + 1
			objWorkSheet.Cells(iRow,iCol) = rst2("Meternum")
			objWorkSheet.Cells(iRow,iCol + 1) = rst2("ManualMultiplier")
			
			objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst2("rawPrevious"),2)
			objWorkSheet.Cells(iRow,iCol + 3) = formatnumber(rst2("rawCurrent"),2)
			objWorkSheet.Cells(iRow,iCol + 4) = formatnumber(cdbl(rst2("Used"))/usagedivisor,2)
			If utilityid=2 or utilityid=1 or utilityid=6 then
				objWorkSheet.Cells(iRow,iCol + 8) = formatnumber(meterdemandtemp,2)
			End If
			
			iRow = iRow + 1
			objWorkSheet.Cells(iRow,iCol + 1) = "Off Peak" 
			objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst2("rawPreviousoff"),2)
			objWorkSheet.Cells(iRow,iCol + 3) = formatnumber(rst2("rawCurrentoff"),2)
			objWorkSheet.Cells(iRow,iCol + 4) = formatnumber(cdbl(rst2("Usedoff"))/usagedivisor,2)
			objWorkSheet.Cells(iRow,iCol + 8) = Formatnumber(meterdemandtempoff,2)
			
			iRow = iRow + 1
			objWorkSheet.Cells(iRow,iCol + 1) = "Mid Peak" 
			objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst2("rawPreviousint"),2)
			objWorkSheet.Cells(iRow,iCol + 3) = formatnumber(rst2("rawCurrentint"),2)
			objWorkSheet.Cells(iRow,iCol + 4) = formatnumber(cdbl(rst2("Usedint"))/usagedivisor,2)
			objWorkSheet.Cells(iRow,iCol + 8) = Formatnumber(meterdemandtempint,2)
		Else
			iRow = iRow + 1
			objWorkSheet.Cells(iRow,iCol) = rst2("Meternum")
			objWorkSheet.Cells(iRow,iCol + 1) = rst2("ManualMultiplier")
			
			objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst2("rawPrevious"),2)
			objWorkSheet.Cells(iRow,iCol + 3) = formatnumber(rst2("rawCurrent"),2)
			
				
			if not extusage then
				if calcintpeak then
					objWorkSheet.Cells(iRow,iCol + 4) = formatnumber(cdbl(rst2("OnPeak"))/usagedivisor,2)
					objWorkSheet.Cells(iRow,iCol + 5) = formatnumber(cdbl(rst2("IntPeak"))/usagedivisor,2)
				else
					if utilityid<>4 then
						objWorkSheet.Cells(iRow,iCol + 4) = formatnumber((cdbl(rst2("OnPeak"))+IntPeak)/usagedivisor,2)
					Else
						objWorkSheet.Cells(iRow,iCol + 4) = ""
					End If
				End IF
					if utilityid<>4 then
						objWorkSheet.Cells(iRow,iCol + 5) = formatnumber(cdbl(rst2("OffPeak"))/usagedivisor,2)
					Else
						objWorkSheet.Cells(iRow,iCol + 5) = ""
					End If				
			End IF
			objWorkSheet.Cells(iRow,iCol + 6) = formatnumber(cdbl(rst2("Used"))/usagedivisor,2)
			
			if utilityid=2 or utilityid=1 or utilityid=6 then
				If rst2("coincident")="True" then
					objWorkSheet.Cells(iRow,iCol + 8)  =  0
				Else
					objWorkSheet.Cells(iRow,iCol + 8) = formatnumber(meterdemandtemp,2)
				End If
				if calcintpeak then
					If rst2("coincident")="True" then
						objWorkSheet.Cells(iRow,iCol + 8)  =  0
						objWorkSheet.Cells(iRow,iCol + 9)  =  0
					Else
						objWorkSheet.Cells(iRow,iCol + 8) = formatnumber(meterdemandtempint,2)
						objWorkSheet.Cells(iRow,iCol + 9) = formatnumber(meterdemandtempoff,2)	
					End If				
					
				End If
			End If
				
		End If

		rst2.movenext
	Loop
		
	If extusage then 
		iRow = iRow + 1
		
		objWorkSheet.Cells(iRow,iCol + 3) = "Meter Totals"
		objWorkSheet.Cells(iRow,iCol + 3).Font.Bold = True 
		objWorkSheet.Cells(iRow,iCol + 4) = "On"
		objWorkSheet.Cells(iRow,iCol + 6) = "Off"
		objWorkSheet.Cells(iRow,iCol + 7) = "Mid"
		objWorkSheet.Cells(iRow,iCol + 8) = "Total"
		
		iRow = iRow + 1
		
		objWorkSheet.Cells(iRow,iCol + 3) = usage
		objWorkSheet.Cells(iRow,iCol + 4) = formatnumber(totalKWH,2)
		objWorkSheet.Cells(iRow,iCol + 6) = formatnumber(totalKWHoff,2)
		objWorkSheet.Cells(iRow,iCol + 7) = formatnumber(totalKWHint,2)
		objWorkSheet.Cells(iRow,iCol + 8) = formatnumber(totalKWH+totalkwhoff+totalkwhint,2)
		
		iRow = iRow + 1
		objWorkSheet.Cells(iRow,iCol + 3) = demand
		objWorkSheet.Cells(iRow,iCol + 4) = formatnumber(totaldemand_PC)
		objWorkSheet.Cells(iRow,iCol + 6) = formatnumber(totaldemand_PCoff)
		objWorkSheet.Cells(iRow,iCol + 7) = formatnumber(totaldemand_PCint)
		objWorkSheet.Cells(iRow,iCol + 8) = formatnumber(totaldemand_PC+totaldemand_PCint+totaldemand_PCoff,2)
				
	Else 
		iRow = iRow + 1
		objWorkSheet.Cells(iRow,iCol + 3) = "Meter Totals"
		objWorkSheet.Cells(iRow,iCol + 3).Font.Bold = True 
		If utilityid <> 4 then
			objWorkSheet.Cells(iRow,iCol + 4) =formatnumber(totalOnPeak,2)
		Else
			objWorkSheet.Cells(iRow,iCol + 4) = ""
		End If
		if calcintpeak then
			icol = icol + 1
			objWorkSheet.Cells(iRow,iCol + 4) = formatnumber(totalIntPeak,0)
		End If	
		
		if utilityid <> 4 Then
			objWorkSheet.Cells(iRow,iCol + 5) = formatnumber(totalOffPeak,2)
		else
			objWorkSheet.Cells(iRow,iCol + 5) = ""
		end if
		objWorkSheet.Cells(iRow,iCol + 6) = formatnumber(totalKWH,2)
		if utilityid=2 or utilityid=1 or utilityid=6 then
			objWorkSheet.Cells(iRow,iCol + 8) = formatnumber(totaldemand_PC,2)
			if calcintpeak then
				objWorkSheet.Cells(iRow,iCol + 9) = formatnumber(totaldemand_PCint,2)
				objWorkSheet.Cells(iRow,iCol + 9) = formatnumber(totaldemand_PCoff,2)
			End If
		End If

	End If
	

	if not rst1("ibsexempt") then
		if extusage then 

				bldgOnPeak = bldgOnPeak + totalKWH
				bldgOffPeak = bldgOffPeak + totalkwhoff
				bldgIntPeak = bldgIntPeak + totalkwhint 
				bldgTotalPeak = bldgTotalPeak + totalKWH + totalkwhoff + totalkwhint
		else
				bldgOnPeak = bldgOnPeak + totalOnPeak
			if calcintpeak then 
					bldgIntPeak = bldgIntPeak + totalIntPeak 
			end if
			
				bldgOffPeak = bldgOffPeak + totalOffPeak		
				bldgTotalPeak = bldgTotalPeak + totalKWH 
		end if
		
	 		bldgTotalKWon = bldgTotalKWon + totaldemand_PC
			bldgTotalKWoff = bldgTotalKWoff + totaldemand_PCoff
			bldgTotalKWint = bldgTotalKWint + totaldemand_PCint

		bldgAdmin = bldgAdmin + (cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDbl(rst1("credit")))*cdbl(rst1("adminfee"))
		if ucase(trim(rst1("type"))) = "LPLS2" then bldgAdmin = bldgAdmin + cdbl(rst1("rate_servicefee"))

		bldgService = bldgService + cDbl(rst1("serviceFee"))
		bldgCredit = bldgCredit + cDbl(rst1("credit"))
		bldgSubtotal = bldgSubtotal + cDbl(rst1("subtotal"))
		bldgTax = bldgTax + cDbl(rst1("tax"))
		'PA condition goes here
		
		bldgTotalAmt = bldgTotalAmt + cDbl(rst1("TotalAmt"))
		
		subsubtotal = subsubtotal + cDbl(rst1("energy"))+cDBL(rst1("demand"))+cDBL(rst1("rateservicefee_dollar"))
		
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
	
	iRow = iRow + 2
	
	objWorkSheet.Cells(iRow,iCol + 1) = "Service Class"
	objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
	
	objWorkSheet.Cells(iRow,iCol + 2) = "Admin Fee"
	objWorkSheet.Cells(iRow,iCol + 2).Font.Bold = True
	
	
	if utilityid=2 then
		objWorkSheet.Cells(iRow,iCol + 3) = "El. Adj. Factor"
		objWorkSheet.Cells(iRow,iCol + 3).Font.Bold = True
	End If
	
	If utilityid=3 then
		objWorkSheet.Cells(iRow,iCol + 4) = "Sewer Charge "
	Else
		objWorkSheet.Cells(iRow,iCol + 4) = "Demand Charge "
	End If
	objWorkSheet.Cells(iRow,iCol + 4).Font.Bold = True
	
	If utilityid=3 then
		objWorkSheet.Cells(iRow,iCol + 5) = "Water Charge "
	Else
		objWorkSheet.Cells(iRow,iCol + 5) = "Consumption Charge "
	End If
	objWorkSheet.Cells(iRow,iCol + 5).Font.Bold = True
	
	If isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" then
		objWorkSheet.Cells(iRow,iCol + 6) = "LMEP Rate "
	Else
		objWorkSheet.Cells(iRow,iCol + 6) = "Modify Rate "
	End If
	objWorkSheet.Cells(iRow,iCol + 6).Font.Bold = True
		
	If ucase(trim(rst1("type"))) = "LPLS2" then 
		objWorkSheet.Cells(iRow,iCol + 7) = "Tenant Service Fee "
	Else
		objWorkSheet.Cells(iRow,iCol + 7) = "Service Fee "
	end If
	objWorkSheet.Cells(iRow,iCol + 7).Font.Bold = True
	
	If rst1("rate_servicefee")<>"0" Then 
		iCol = iCol + 1
		objWorkSheet.Cells(iRow,iCol + 7) = "Utility Service Fee "
	End If
	If rst1("adminfee")<>"0" then 
		iCol = iCol + 1
		objWorkSheet.Cells(iRow,iCol + 7) = "Admin Fee "
	End IF	
	objWorkSheet.Cells(iRow,iCol + 7).Font.Bold = True
	
	objWorkSheet.Cells(iRow,iCol + 8) = "Sqft "
	objWorkSheet.Cells(iRow,iCol + 8).Font.Bold = True
	if utilityid="2" then
		iCol = iCol + 1
		objWorkSheet.Cells(iRow,iCol + 8) = "Watts/Sqft "
	End If
	objWorkSheet.Cells(iRow,iCol + 8).Font.Bold = True
	
	iRow = iRow + 1
	iCol = 1
	
	if ucase(trim(rst1("type")))="AVG COST 1" then 
		objWorkSheet.Cells(iRow,iCol + 1) = rst1("type") & " "  & formatnumber(rst1("AvgKWH"),6)
	else
		objWorkSheet.Cells(iRow,iCol + 1) = rst1("type")
	end if
	objWorkSheet.Cells(iRow,iCol + 2) = formatpercent(rst1("AdminFee"),2)
	
	if utilityid=2 then
		objWorkSheet.Cells(iRow,iCol + 3) = formatnumber(rst1("fuelAdjnum"),5)
	End If
	
	objWorkSheet.Cells(iRow,iCol + 4) = formatcurrency(rst1("demand"),2)
	objWorkSheet.Cells(iRow,iCol + 5) = formatcurrency(rst1("energy"),2)
	
	If isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" Then
		objWorkSheet.Cells(iRow,iCol + 6) = formatcurrency(rst1("unit_credit"),6)
	Else
		objWorkSheet.Cells(iRow,iCol + 6) = formatcurrency(rst1("RateModify"),6)
	End If
	objWorkSheet.Cells(iRow,iCol + 7) = formatcurrency(cDbl(rst1("serviceFee")))
	
	if rst1("rate_servicefee")<>"0" then
		objWorkSheet.Cells(iRow,iCol + 7) = formatcurrency(cDbl(rst1("rate_servicefee")))
	End  If
	
	If rst1("adminfee")<>"0" then 
		iCol = iCol + 1
		objWorkSheet.Cells(iRow,iCol + 7) = formatcurrency((cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDBL(rst1("credit")))*cdbl(rst1("adminfee")),2)		
	End If
	
	objWorkSheet.Cells(iRow,iCol + 8) = getNumber(rst1("sqft"))
	if utilityid="2" then 
		if getNumber(rst1("sqft"))=0 then
			objWorkSheet.Cells(iRow,iCol + 9) = 0
		else
			objWorkSheet.Cells(iRow,iCol + 9) =  formatnumber((totaldemand_PC*1000)/cDbl(rst1("sqft"))) 
		end if
		iRow = iRow + 1
		objWorkSheet.Cells(iRow,iCol + 1) = "Annualized Cost Per Square Foot For This Bill"
		objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
		
		if getNumber(rst1("sqft"))=0 then
			objWorkSheet.Cells(iRow,iCol + 7) = 0
		else
			objWorkSheet.Cells(iRow,iCol + 7) =  formatcurrency((cdbl(rst1("TotalAmt"))*12)/cDbl(rst1("sqft")))
		end if
		
	end if
	
	iRow = iRow + 2
	iCol = 1

	if PortFolioId = "" then PortFolioId = 0
		objWorkSheet.Cells(iSubTotalRow,12) = "Sub Total: "
		objWorkSheet.Cells(iSubTotalRow,12).Font.Bold = True
		
		objWorkSheet.Cells(iSubTotalRow,13) = formatcurrency(cDbl(rst1("energy"))+cDBL(rst1("demand")),2)
		
		if cdbl(rst1("rateservicefee_dollar"))>0 then
			iSubTotalRow = iSubTotalRow + 1
			objWorkSheet.Cells(iSubTotalRow,12) = "Rate Service Fee:"
			objWorkSheet.Cells(iSubTotalRow,12).Font.Bold = True
			objWorkSheet.Cells(iSubTotalRow,13) = formatcurrency(rst1("rateservicefee_dollar"),2)
		End If
		if PortFolioId=49 then
			iSubTotalRow = iSubTotalRow + 1
			objWorkSheet.Cells(iSubTotalRow,12)= "Admin/Service Fee:"
			objWorkSheet.Cells(iSubTotalRow,12).Font.Bold = True
			objWorkSheet.Cells(iSubTotalRow,13) = formatcurrency(cdbl(rst1("servicefee"))+cdbl(rst1("admindollar")),2)	
		End IF
		
		if PortFolioId<>49 then
			iSubTotalRow = iSubTotalRow + 1
			objWorkSheet.Cells(iSubTotalRow,12) = "Admin/Service Fee:"
			objWorkSheet.Cells(iSubTotalRow,12).Font.Bold = True
			objWorkSheet.Cells(iSubTotalRow,13) = formatcurrency(cdbl(rst1("servicefee"))+cdbl(rst1("admindollar")),2)	
		End If
		
		iSubTotalRow = iSubTotalRow + 1
		objWorkSheet.Cells(iSubTotalRow,12) = "Sub Total:"
		objWorkSheet.Cells(iSubTotalRow,12).Font.Bold = True
		objWorkSheet.Cells(iSubTotalRow,13) = formatcurrency(rst1("subtotal"),2)	
		
		iSubTotalRow = iSubTotalRow + 1
		objWorkSheet.Cells(iSubTotalRow,12) = "Sales Tax:"
		objWorkSheet.Cells(iSubTotalRow,12).Font.Bold = True
		objWorkSheet.Cells(iSubTotalRow,13) = formatcurrency(cDbl(rst1("tax")),2)	
		
		
		iSubTotalRow = iSubTotalRow + 1
		if isnumeric(rst1("unit_credit")) and trim(rst1("unit_credit"))<>"0" then
			objWorkSheet.Cells(iSubTotalRow,12) = "LMEP Credit:"
		elseif PortFolioId=49 then
			objWorkSheet.Cells(iSubTotalRow,12) = "Restructuring Rate Reduction:"
		else
			objWorkSheet.Cells(iSubTotalRow,12) = "Credit/Adjustment:"
		end if
		objWorkSheet.Cells(iSubTotalRow,12).Font.Bold = True
		objWorkSheet.Cells(iSubTotalRow,13) = formatcurrency(rst1("credit"),2)			
		
		iSubTotalRow = iSubTotalRow + 1
		objWorkSheet.Cells(iSubTotalRow,12) = "Total Charges:"
		objWorkSheet.Cells(iSubTotalRow,13) = formatcurrency(cDbl(rst1("TotalAmt")))			
		objWorkSheet.Cells(iSubTotalRow,12).Font.Bold = True
		
		iRow = iRow + 1
	rst1.movenext
	'# Added by Tarun 07/10/2006
	'lngTenantCount = lngTenantCount + 1
	'# 
loop
rst1.close

select case utilityid
case 1 'steam
	sql = "SELECT SalesTax, (case when isnull(taxincluded,0)=1 then TotalBillAmt-isnull(SalesTax,0) else TotalBillAmt end) as subtotal, isnull(MLbUsage,0) as MLbUsage, (case when isnull(taxincluded,0)=0 then isnull(TotalBillAmt,0)+isnull(SalesTax,0) else isnull(TotalBillAmt,0) end) as TotalBillAmt, isnull(AvgCost,0) as AvgCost FROM Utilitybill_steam u, billyrperiod bp WHERE u.ypid=bp.ypid and bldgnum='"&building&"' and billyear="&Billyear&" and billperiod="&Billperiod
case 3 'water
	sql = "SELECT salestax, TotalBillAmt-salestax as subtotal, isnull(totalccf,0) as totalccf, isnull(watercharge,0) as watercharge, isnull(SewerCharge,0) as SewerCharge, isnull(TotalbillAmt,0) as TotalbillAmt, isnull(avgcost,0) as avgcost FROM Utilitybill_coldwater u, billyrperiod bp WHERE u.ypid=bp.ypid and bldgnum='"&building&"' and billyear="&Billyear&" and billperiod="&Billperiod 
case 4 'gas
	sql = "SELECT salestax, (case when isnull(taxincluded,0)=1 then TotalBillAmt-isnull(SalesTax,0) else TotalBillAmt end) as subtotal, isnull(ThermUsage,0) as ThermUsage, isnull(ccfUsage,0) as ccfUsage, (case when isnull(taxincluded,0)=0 then isnull(TotalBillAmt,0)+isnull(SalesTax,0) else isnull(TotalBillAmt,0) end) as TotalBillAmt, isnull(avgcosttherm,0) as avgcosttherm FROM Utilitybill_gas u, billyrperiod bp WHERE u.ypid=bp.ypid and bldgnum='"&building&"' and billyear="&Billyear&" and billperiod="&Billperiod 
case else 'electricity?
	sql = "select distinct ub.salestax, TotalBillAmt-ub.intot as subtotal, OnPeakKWH as OnPeakKWH,OffPeakKWH as OffPeakKWH, TotalKWH as TotalKWH, CostKWH as CostKWH,(case when TotalKWH=0 then 0 else CostKWH/TotalKWH end) as UnitCostKWH, TotalKW as TotalKW,CostKW as CostKW, (case when TotalKW=0 then 0 else CostKW/TotalKW end) as UnitCostKW, isnull(TotalBillAmt,0)+extot as TotalBillAmt,(case when [totalkw]*24*DateDiff(day,[ypiddatestart],[ypidDateEnd])=0 then 0 else [totalkwh]/([totalkw]*24*(DateDiff(day,[ypiddatestart],[ypidDateEnd])+1)) end) as loadfactor from "&rpt&" ub  WHERE ub.reject=0 and ub.bldgnum='"&building&"' and ub.billyear="&Billyear&" and ub.billperiod="&Billperiod&" and ub.utility="&utilityid
end select
'response.write sql
'response.end
rst1.open sql, cnn1

' Building Totals
iRow = iRow + 2
iCol = 1
objWorkSheet.Cells(iRow,iCol) = "Building Totals"
objWorkSheet.Cells(iRow,iCol).Font.Bold = True
Dim iBldgTotalsRow

if not rst1.eof then
    'changed for testing purpose
	if not PortFolioId = 108 then
	 'if PortFolioId = 108 then
		if not utilityid=6 then
			iRow = iRow + 1
			iBldgTotalsRow = iRow
			objWorkSheet.Cells(iRow,iCol + 1) = "Utility Expenses"	
			objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
			select case utilityid
					case 1
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = usage & " Used"		
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("MLbUsage"),0)	
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Subtotal"
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True	
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("subtotal"),2)	
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Sales Tax"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("salestax"),2)								
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Utility Bill Amount"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("TotalBillAmt"),2)								

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Average Cost per"	& usage
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("AvgCost"),2)						
					case 3	
					
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "CCF Used"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("totalccf"),0)
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Water Charge"
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True	
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("watercharge"),2)						

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Sewer Charge"
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True	
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("sewercharge"),2)	
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Subtotal"
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True	
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("subtotal"),2)	
											
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Sales Tax"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("salestax"),2)	

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Utility Bill Amount"
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True	
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("TotalBillAmt"),2)	
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Average Cost per CCF"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("AvgCost"),2)
						
					case 4
					
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "CCF Used"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("ccfUsage"),0)
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Therms Used"
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True	
						objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("ThermUsage"),0)						

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Subtotal"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("subtotal"),2)	
											
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Sales Tax"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("salestax"),2)	

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Utility Bill Amount"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("TotalBillAmt"),2)	
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Average Cost per Therm"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("avgcosttherm"),2)																									
					case 6
					case else
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "On Peak" & usage	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("OnPeakKWH")) then
							objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("OnPeakKWH"),0)
						end If

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Off Peak" & usage	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("OffPeakKWH")) then
							objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("OffPeakKWH"),0)
						end If				
								

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Total " & usage	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("TotalKWH")) then
							objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("TotalKWH"),0)
						end If				
																
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Cost " & usage	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("CostKWH")) then
							objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("CostKWH"),0)
						end If				

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Unit Cost " & usage	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("UnitCostKWH")) then
						
							objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("UnitCostKWH"),4)
						end If										

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Total" & demand 	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("TotalKW")) then
							objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(rst1("TotalKW"),0)
						end If	
						

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Cost" & demand 	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("CostKW")) then
							objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("CostKW"))
						end If	
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Unit Cost" & demand 	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("UnitCostKW")) then
							objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("UnitCostKW"),2)
						end If		
						
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Load Factor " 	 
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("loadfactor")) then
							objWorkSheet.Cells(iRow,iCol + 2) = formatpercent(rst1("loadfactor"),2)
						end If									
																							
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Subtotal"
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("subtotal")) then	
							objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("subtotal"),2)	
						end if
											
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Sales Tax"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("salestax")) then	
							objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("salestax"),2)	
						end if

						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Utility Bill Amount"	
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True
						if isnumeric(rst1("TotalBillAmt")) then	
							objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("TotalBillAmt"),2)	
						end if
						iRow = iRow + 1 
						objWorkSheet.Cells(iRow,iCol + 1) = "Average Cost"
						objWorkSheet.Cells(iRow,iCol + 1).Font.Bold = True	
						if isnumeric(rst1("TotalBillAmt")) and clng(rst1("TotalKWH"))<>0 then
							objWorkSheet.Cells(iRow,iCol + 2) = formatcurrency(rst1("TotalBillAmt")/rst1("TotalKWH"),6)					
						end if
				end select
			end if
		end if				
	end if	
	
	iRow = iRow + 1 
	iCol = 5
	
	'response.write("irow : " + cstr(iRow))
    'response.Write("iCol : " + cstr(iCol))
    'response.Write("totalsRow" + cstr(iBldgTotalsRow))
    'response.End 

    if utilityid <> 2 then
       iBldgTotalsRow = iRow - 1
    end if
       
    
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Sub-Meter Revenue"
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	if utilityid=2 then

		iBldgTotalsRow = iBldgTotalsRow + 1 
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "On Peak" 
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True	 
		if isnumeric(bldgOnPeak) then
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatnumber(bldgOnPeak,2)
		end If
		
		if calcintpeak OR extusageflag then
			iBldgTotalsRow = iBldgTotalsRow + 1
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Int Peak" 	
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
			if isnumeric(bldgIntPeak) then
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatnumber(bldgIntPeak,2)
			end if
		end if
		
			iBldgTotalsRow = iBldgTotalsRow + 1
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Off Peak" 
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True	
			if isnumeric(bldgOffPeak) then
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatnumber(bldgOffPeak,2)
			end if
	end if	

   
	
	iBldgTotalsRow = iBldgTotalsRow + 1 
	if utilityid=3 or utilityid=4 then 
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Total C" & usage 
	else
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Total " & usage 
	End if
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	if isnumeric(bldgTotalPeak) then 
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatnumber(bldgTotalPeak,2)
	end If
	
	if utilityid=2 then
		if calcintpeak OR extusageflag then
			iBldgTotalsRow = iBldgTotalsRow + 1 
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = demand & " On"
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
			if isnumeric(bldgTotalKWon) then 
				objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatnumber(bldgTotalKWon,2)
			end If
			
			iBldgTotalsRow = iBldgTotalsRow + 1 
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = demand & " Int"
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
			if isnumeric(bldgTotalKWint) then 
				objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatnumber(bldgTotalKWint,2)
			end If
			
			iBldgTotalsRow = iBldgTotalsRow + 1 
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = demand & " Off"
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
			if isnumeric(bldgTotalKWoff) then 
				objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatnumber(bldgTotalKWoff,2)
			end If			
		end if
			iBldgTotalsRow = iBldgTotalsRow + 1 
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Total " & demand
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
			if isnumeric(bldgTotalKWon+bldgTotalKWoff+bldgTotalKWint) Then
				objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatnumber(bldgTotalKWon+bldgTotalKWoff+bldgTotalKWint,0)
			end If			
	end if

	iBldgTotalsRow = iBldgTotalsRow + 1 
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Subtotal "
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	if isnumeric(subsubtotal) then
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatcurrency(subsubtotal)
	end If				
	
	iBldgTotalsRow = iBldgTotalsRow + 1 
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Admin Fee "
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	if isnumeric(bldgAdmin) then
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatcurrency(bldgAdmin)
	end If				
	
	iBldgTotalsRow = iBldgTotalsRow + 1 
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Service Fee "
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	if isnumeric(bldgService) then
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatcurrency(bldgService)
	end If		
	
	iBldgTotalsRow = iBldgTotalsRow + 1 
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Credit "
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	if isnumeric(bldgCredit) then
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatcurrency(bldgCredit)
	end If	
	
	iBldgTotalsRow = iBldgTotalsRow + 1 
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Subtotal "
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	if isnumeric(bldgSubtotal) then
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatcurrency(bldgSubtotal)
	end If		
	
	iBldgTotalsRow = iBldgTotalsRow + 1 
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Tax "
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	if isnumeric(bldgTax) then
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatcurrency(bldgTax,2)
	end If			
	
	iBldgTotalsRow = iBldgTotalsRow + 1 
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "Total "
	objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	
	if isnumeric(bldgTotalAmt) then
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatcurrency(bldgTotalAmt,2)
	end If	
	
	if utilityid<>10 and utilityid<>6 and not(rst1.eof) then
		iBldgTotalsRow = iBldgTotalsRow + 1 
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "% Recoup "
		objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
	
		if isnumeric(bldgSubtotal) and trim(rst1("TotalBillAmt"))<>"0" then
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatpercent(bldgSubtotal/cdbl(rst1("TotalBillAmt")),2)
		end If	
		if utilityid=2 then
			iBldgTotalsRow = iBldgTotalsRow + 1 
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1) = "% Recoup (KWH)"
			objWorkSheet.Cells(iBldgTotalsRow,iCol + 1).Font.Bold = True
			if isnumeric(bldgTotalPeak) and trim(rst1("TotalKWH"))<>"0" then
				objWorkSheet.Cells(iBldgTotalsRow,iCol + 2) = formatpercent(bldgTotalPeak/cdbl(rst1("TotalKWH")),2)
			end If				
		end if
	end if
	
	
	If ExmpData then
	
		iRow = iRow + 1 
		objWorkSheet.Cells(iRow,iCol + 1) = "Revenue Exempt Summary "
	
	
		iRow = iRow + 1
		if utilityid=2 then
			objWorkSheet.Cells(iRow,iCol + 1) = "On Peak"
			if isnumeric(ExmpOnPeak) then
				objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(ExmpOnPeak,2)
			end If
			
			if calcintpeak OR extusageflag then
				iRow = iRow + 1 
				objWorkSheet.Cells(iRow,iCol + 1) = "Int Peak"
				if isnumeric(ExmpIntPeak) then
					objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(ExmpIntPeak,2)
				end If			
			end if		
			objWorkSheet.Cells(iRow,iCol + 1) = "Off Peak"
			if isnumeric(ExmpOffPeak) then
				objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(ExmpOffPeak,2)
			end If			
		End if
		
		iRow = iRow + 1 
				
		if utilityid=3 then
			objWorkSheet.Cells(iRow,iCol + 1) = "Total C" & usage
		else
			objWorkSheet.Cells(iRow,iCol + 1) = "Total" & usage
		end if
		
		if isnumeric(ExmpTotalPeak) then
			objWorkSheet.Cells(iRow,iCol + 2) = formatnumber(ExmpTotalKW,0)
		End If
		
		if utilityid=2 then
			if calcintpeak OR extusageflag then 
				iRow = iRow + 1 
				objWorkSheet.Cells(iRow,iCol + 1) = demand & " On"
				if isnumeric(ExmpTotalKWon) then
					objWorkSheet.Cells(iRow,iCol + 2) =  formatnumber(ExmpTotalKWon,0)
				end if
				
				iRow = iRow + 1 
				objWorkSheet.Cells(iRow,iCol + 1) = demand & " Int"
				if isnumeric(ExmpTotalKWint) then
					objWorkSheet.Cells(iRow,iCol + 2) =  formatnumber(ExmpTotalKWint,0)
				end if
				
				iRow = iRow + 1 
				objWorkSheet.Cells(iRow,iCol + 1) = demand & " Off"
				if isnumeric(ExmpTotalKWoff) then
					objWorkSheet.Cells(iRow,iCol + 2) =  formatnumber(ExmpTotalKWoff,0)
				end if								
			end if
			
			iRow = iRow + 1 
			objWorkSheet.Cells(iRow,iCol + 1) = "Total " & demand 
			if isnumeric(ExmpTotalKW) then
				objWorkSheet.Cells(iRow,iCol + 2) =  formatnumber(ExmpTotalKW,0)
			end if		
					
		End if
		
		iRow = iRow + 1 
		objWorkSheet.Cells(iRow,iCol + 1) = "Subtotal" 
		if isnumeric(Exmpsubsubtotal) then
			objWorkSheet.Cells(iRow,iCol + 2) =  formatcurrency(Exmpsubsubtotal)
		end if	
		
		iRow = iRow + 1 
		objWorkSheet.Cells(iRow,iCol + 1) = "Admin Fee" 
		if isnumeric(ExmpAdmin) then 
			objWorkSheet.Cells(iRow,iCol + 2) =  formatcurrency(ExmpAdmin)
		end if		
		
		iRow = iRow + 1 
		objWorkSheet.Cells(iRow,iCol + 1) = "Service Fee" 
		if isnumeric(ExmpService) then 
			objWorkSheet.Cells(iRow,iCol + 2) =  formatcurrency(ExmpService)
		end if			
		
		iRow = iRow + 1 
		objWorkSheet.Cells(iRow,iCol + 1) = "Credit" 
		if isnumeric(ExmpCredit) then 
			objWorkSheet.Cells(iRow,iCol + 2) =  formatcurrency(ExmpCredit)
		end if	
		
		iRow = iRow + 1 
		objWorkSheet.Cells(iRow,iCol + 1) = "Subtotal" 
		if isnumeric(ExmpSubtotal) then 
			objWorkSheet.Cells(iRow,iCol + 2) =  formatcurrency(ExmpSubtotal)
		end if	

		iRow = iRow + 1 
		objWorkSheet.Cells(iRow,iCol + 1) = "Tax" 
		if isnumeric(ExmpTax) then 
			objWorkSheet.Cells(iRow,iCol + 2) =  formatcurrency(ExmpTax)
		end if		
		
		iRow = iRow + 1 
		objWorkSheet.Cells(iRow,iCol + 1) = "Total" 
		if isnumeric(ExmpTotalAmt) then 
			objWorkSheet.Cells(iRow,iCol + 2) =  formatcurrency(ExmpTotalAmt)
		end if		
							
							
	End if	
	objExcelReport.DisplayAlerts = False
	objWorkBook.SaveCopyAs("D:\WebSites\isabella\appserver1\BillSummary\" & building & Billperiod & Billyear & "BillSummary.xlsx")
	'objWorkBook.SaveAs("\\10.0.8.62\web_folders\finance\VO\BillSummary\" & building & Billperiod & Billyear & "BillSummary.xls")
	'objWorkBook.SaveAs("E:\websites\appserver1\genergy2\BillSummary\" & building & Billperiod & Billyear & "BillSummary.xls")
	objExcelReport.DisplayAlerts = True
	objExcelReport.Quit
	
	' Set up Email to be Sent
	Dim objEmail 
	Dim strSQL
	Dim strMailingList
	Dim rstMailingList
		
	'email=no
	if (email="no") then %>
	<table>
	    <tr>
	    <td>
	        <!-- for test
	        <a style="font-family:arial;font-size:12;text-decoration:none;" href="http://paappserver.genergy.com/genergy2/BillSummary/<%=building%><%=Billperiod%><%=Billyear%>BillSummary.xls" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'blue'"><b>click here to view the Excel file: <%=building%><%=Billperiod%><%=Billyear%>BillSummary.xls </b></a> 
	        <a style="font-family:arial;font-size:12;text-decoration:none;" href="https://appserver1.genergy.com/eri_TH/finance/VO/BillSummary/<%=building%><%=Billperiod%><%=Billyear%>BillSummary.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'blue'"><b>click here to view the Excel file: <%=building%><%=Billperiod%><%=Billyear%>BillSummary.xlsx </b></a> 
			-->
			<a style="font-family:arial;font-size:12;text-decoration:none;" href="https://appserver1.genergy.com/BillSummary/<%=building%><%=Billperiod%><%=Billyear%>BillSummary.xlsx" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'blue'"><b>click here to view the Excel file: <%=building%><%=Billperiod%><%=Billyear%>BillSummary.xlsx </b></a> 
	    </td>
	    </tr>
	</table>
	<%
	else 	
	    'email=yes	
		Dim Mail, cdoConfig, Fields
  
		Set cdoConfig = Server.CreateObject("CDO.Configuration")  
		Set Fields = cdoConfig.Fields
		With Fields  
			.Item(cdoSMTPServer) = "2021dc"  
			.Update  
		End With  
	    Set Mail = Server.CreateObject("CDO.Message")
	
	    Set rstMAilingList =  server.createobject("ADODB.Recordset")

	    strSQL = "SELECT email FROM contacts Where submeter_bills=1 and bldgnum ='" & building & "' and cid ='" & PortFolioId & "'"
	    strMailingList = ""
	    rstMAilingList.open strSQL , getConnect(PortFolioId,building,"Billing")
	    If not rstMailingList.EOF Then
		    Do While not rstMailingList.EOF 
			    if len(strMailingList) > 0 then 
				    strMailingList = strMailingList & ";" & rstMailingList("Email")
			    else
				    strMailingList = rstMailingList("Email")
			    end if
			    rstMailingList.MoveNext 
		    Loop 
	    End IF
	    ' If There is a mailing List then
	    If Len(strMailingList) >= 0 then
		    'objEmail.To = strMailingList
			With Mail
				Set .Configuration = cdoConfig
				.To = "rb@cplems.com,robertm@cplems.com,leroii@cplgroupusa.com,davideg@cplgroupusa.com"
				.From = "filestore@cplems.com"
				.Subject = "Bill Summary for Building " & building & " , Period " & Billperiod & " " & Billyear 
				.TExtBody = "Bill Summary for Building " & building & " , Period " & Billperiod & " " & Billyear & "BillSummary.xlsx"
				.AddAttachment "D:\WebSites\isabella\appserver1\BillSummary\" & building & Billperiod & Billyear & "BillSummary.xlsx"
				'objEmail.AttachFile "\\10.0.8.62\web_folders\finance\VO\BillSummary\" & building & Billperiod & Billyear & "BillSummary.xls" , building & Billperiod & Billyear & "BillSummary.xls"
				'objEmail.AttachFile "E:\websites\appserver1\genergy2\BillSummary\" & building & Billperiod & Billyear & "BillSummary.xls" , building & Billperiod & Billyear & "BillSummary.xls"		    
				.Send
			end With
			set Mail = nothing	
			Response.Write "<HTML><HEAD></HEAD><BODY><P> Bill Summary Excel Sheet Generated and sent to Billing<BR>"
			Response.Write strMailingList 
			Response.Write "</P></BODY></HTML>"
	    Else
		    Response.Write "<HTML><HEAD></HEAD><BODY><P> No Mailing List is Available for the Building <BR>"
		    Response.Write "</P></BODY></HTML>"
	    End IF
    end if 
    	
	
	set rstMailingList = Nothing
	set objExcelReport = Nothing
	set rst1 = Nothing
	set rst2 = Nothing
	set rst3 = Nothing
	set cnn1 = Nothing
	
	
%>	
	
<%
	Dim objSWbemServices, colProcess, objProcess, resultCode
	Set objSWbemServices = GetObject ("WinMgmts:Root\Cimv2")
	Set colProcess = objSWbemServices.ExecQuery ("Select * From Win32_Process WHERE Name LIKE '%EXCEL.EXE%'")
'	For Each objProcess In colProcess
'		response.write _
'		"<ul>"&_
'		"<li>Name="& objProcess.Name      &_
'		"<li>PID ="& objProcess.ProcessId &_
'		"</ul>"
'	Next
	For Each objProcess In colProcess
		resultCode = objProcess.Terminate()
	Next
'	response.end
%>