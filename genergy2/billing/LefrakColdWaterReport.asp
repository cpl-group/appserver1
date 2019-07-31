
<%@ Language=VBScript %>
<%option explicit%>

<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<!--  
    METADATA  
    TYPE="typelib"  
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library"  
--> 

<%
killExcel()
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

	Dim  Billperiod, building, Billyear, PortFolioId, UtilityId, rpt, pdf, Genergy_Users, demo, sql, email, billsummaryFile, ftproot, dailyperiod, dailyfile
	
	Dim uSql, objFSO, ctime, bperiodchar, csvPortfolio, csvUtility, csvPeriod, crlf, pbpath, filePath, csvFile, csvHeaders, tMeter, csvUtil, PcsvFile, resfile, putf, ppath
	dim cUsage, mDemand, mSub, mFees, mTax, mTotal , mRate, tUsage, tDemand, tSub, tAdmin, tService, tFees, tTax, tCredit, tClass, tFeeRate, bRate, billid, ftotal, pusage, hkwhvariance
	putf = false
	Dim rst1, rst2, rst3, rst9, cnn1, rst0
	Dim usage, demand, utilityname
	Dim PUTFStream
	Set PUTFStream = CreateObject("adodb.stream")
	PUTFStream.Type = adTypeText
	PUTFStream.Mode = adModeReadWrite
	PUTFStream.Charset = "UTF-8"
	PUTFStream.LineSeparator = adLF
	PUTFStream.Open
	
	set rst1 = server.createobject("ADODB.Recordset")
	set rst2 = server.createobject("ADODB.Recordset")
	set rst3 = server.createobject("ADODB.Recordset")
	set rst9 = server.createobject("ADODB.Recordset")
	set rst0 = server.createobject("ADODB.Recordset")
	set cnn1 = server.createobject("ADODB.Connection")
	cnn1.open getLocalConnect(building)
	
	Billperiod = request("bperiod")
	building = request("building")
	Billyear = request("byear")
	PortFolioId = request("pid")
	UtilityId = trim(request("utilityid"))
	email = request("email")
			Const NUMBER_PADDING = "000000000000" ' a few zeroes more just to make sure

			Function ZeroPadInteger(i, numberOfDigits)
			  ZeroPadInteger = Right(NUMBER_PADDING & i, numberOfDigits)
			End Function
			
			Class tenant
				public building, tNum, tName, mNum, dFrom, dTo, days, mUsage, mDemand, mSub, mFees, mTax, mTotal, mRate, cUsage, tDemand, tSub, tFees, tTax, tCredit, tClass, tFeeRate, mkwhvariance, unitcode, ucdiff, billid, unitaverage, ftotal, pusage, hkwhvariance, historicalaverage
			end Class	
			dim csvTenant
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
	'rst1.open "SELECT logo, portfolio FROM portfolio p, billtemplates bt " & _
	'			" WHERE p.templateid=bt.id and p.id=(SELECT portfolioid FROM "&DBlocalIP&"buildings WHERE bldgnum='"&building&"')", getConnect(PortFolioId,building,"Billing")
	rst1.open "select portfolio from [dbcore].[dbo].portfolio where id = "&portfolioid
	csvPortfolio = rst1("portfolio")

	rst1.close
	
	sql ="exec Lefrak_UnitAverageCalc "& billyear & ", " & billperiod & ", " & utilityid
	rst1.open sql, cnn1
	
	
	rst0.open "select distinct bldgnum from BillYrPeriod where Utility="&utilityid&" and bldgnum in (select bldgnum from buildings where portfolioid="&portfolioid&")", cnn1
	if not rst0.eof then
	do until rst0.eof 
		
		building = rst0("bldgnum")
		If allowGroups("Genergy Users") Then 
			rpt = "rpt_bill_summary" 
		Else 
			rpt = "rpt_Bill_summary_client"
		End If
		
	'	sql = "SELECT isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.ibsexempt, " & _
	'			 " r.unit_credit, b.strt, r.sqft, r.adminfee, isnull(r.RateModify,0) as RateModify, " & _
	'			 " isnull(r.fuelAdj,0) as fuelAdjnum, rt.[type], AvgKWH, tenantname, datestart-1 as datestart, " & _
	'			 " datediff(day, datestart-1, dateend) as days, dateend, ypid, r.leaseutilityid, r.billingname, " & _
	'			 " r.tenantnum, isnull(r.adjustment,0)-isnull(r.credit,0) as credit, isnull(subtotal,0) as subtotal, " & _
	'			 " isnull(energy,0) as energy, isnull(demand,0) as demand, isnull(serviceFee,0) as serviceFee, " & _
	'			 " isnull(tax,0) as tax, isnull(totalamt,0) as totalamt, b.btstrt, lup.calcintpeak, " & _
	'			 " isnull(adminfeedollar,0) as admindollar, isnull(r.extusg,0) as extusg, r.rate_servicefee, r.shadow, isnull(r.energydetail,0) as energydetail,isnull(r.demanddetail,'') as demanddetail , case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay" & _
	'	  " FROM "&rpt&" r, dbo.ratetypes rt, buildings b, tblleasesutilityprices lup, tblleases l " & _
	'	  " WHERE r.reject=0 and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid " & _
	'				" and b.bldgnum=r.bldgnum AND r.[type]=rt.id and billyear="&Billyear& _
	'				" and billperiod="&Billperiod&" and r.bldgnum='"&building & _
	'				"' and l.billsummaryexempt = 0 and r.utility="&utilityid & _
	'				" and r.tenantnum is not null and left(r.tenantnum,3)='t00' and len(r.tenantnum)= 8 and right(r.bldgnum,1) <> 'c' " &_
	'	 " ORDER BY TenantName asc"	
	'	 
	'	 rst1.open sql, cnn1
	'	
	'	if rst1.eof then
	'		rst1.close
	'		
	'		if allowGroups("Genergy Users") then 
	'			rpt = "rpt_bill_summary_nobill" 
	'		else 
	'			rpt = "rpt_Bill_summary_nobill_client"
	'		End If
	'		
	'		sql = "SELECT isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.ibsexempt, r.unit_credit, " & _
	'					" b.strt, r.sqft, isnull(r.adminfee,0) as adminfee, isnull(r.RateModify,0) as RateModify, " & _
	'					" isnull(r.fuelAdj,0) as fuelAdjnum, rt.[type], AvgKWH, tenantname, datestart-1 as datestart, " & _
	'					" datediff(day, datestart-1, dateend) as days, dateend, ypid, r.leaseutilityid, r.billingname, " & _
	'					" r.tenantnum, isnull(r.adjustment,0)-isnull(r.credit,0) as credit, isnull(subtotal,0) as subtotal, " & _
	'					" isnull(energy,0) as energy, isnull(demand,0) as demand, isnull(serviceFee,0) as serviceFee, " & _
	'					" isnull(tax,0) as tax, isnull(totalamt,0) as totalamt, b.btstrt, lup.calcintpeak, " & _
	'					" isnull(adminfeedollar,0) as admindollar, isnull(r.extusg,0) as extusg, lup.calcintpeak, " & _
	'					" r.rate_servicefee, r.shadow, isnull(r.energydetail,0) as energydetail, isnull(r.demanddetail,'') as demanddetail, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay" & _
	'			 " FROM "&rpt&" r, dbo.ratetypes rt, buildings b, tblleasesutilityprices lup, tblleases l " & _
	'			 " WHERE r.reject=0 and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid " & _
	'					" and b.bldgnum=r.bldgnum AND r.[type]=rt.id and billyear="&Billyear& _
	'					" and billperiod="&Billperiod&" and r.bldgnum='"&building&"' and l.billsummaryexempt = 0 " & _
	'					" and r.utility="&UtilityId&_
	'					" and r.tenantnum is not null and left(r.tenantnum,3)='t00' and len(r.tenantnum)= 8 and right(r.bldgnum,1) <> 'c' " &_
	'			" ORDER BY TenantName asc"
	'		rst1.open sql, cnn1
	'	end if	 
	'	'response.write sql & "</br>"
	'	'response.end
	'	if not rst1.eof then
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

				' Call in code:

				ctime = year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
				
				csvPeriod = Billyear & "." & ZeroPadInteger(Billperiod,2)
				dailyPeriod = Billyear & "." & ZeroPadInteger(Billperiod,2) & "." & ZeroPadInteger(day(now),2)
				
				uSql = "select utilitydisplay from tblutility where utilityid=" & UtilityId
				rst9.open uSql, cnn1, 3
				csvUtility = rst9("utilitydisplay")	
				rst9.close
				csvUtil = ZeroPadInteger(utilityid,2)
				
				crlf = chr(13) & chr(10)
				' Create new csv file 
				pbpath = "\" & csvPortfolio & "\" & building & "\"
				ppath = "\" & csvPortfolio & "\" 
				'response.write pbpath & "</br>"
				dim exportRoot, folders, fs, i, f
				
				exportRoot = "D:\Submetering"
				ftproot = "d:\FTP\ghnet\LeFrakData\Export Files\"
				filePath = exportRoot & pbpath
				csvFile = ucase(building) & "_" & csvPeriod & "_" & csvUtility & "_" & "Report" & ".txt"
				if portfolioid = 1171 then
					PcsvFile = "Residential" & "_" & csvPeriod & "_" & csvUtility & "_" & "Report" & ".csv"
				else
					PcsvFile = csvPeriod & "_" & csvUtility & "_" & "Report" & ".csv"
				end if
				dailyFile = building & "_" & dailyPeriod & "_" & csvUtility & "_" & "Report" & ".txt"
				billsummaryFile = building & "_" & csvPeriod & "_" & csvUtility & "_" & "BillSummary"  & ".xlsx"
				resfile = exportroot & "\" & csvportfolio & "\" & pcsvfile
				Set fs=Server.CreateObject("Scripting.FileSystemObject")

				folders = Split(pbpath, "\")

				'//Create folders if they don't exist
				'For i = 0 To UBound(folders)
			'		exportRoot = exportRoot & folders(i)
			'		response.write(exportRoot)&"</br>"
			'		If not(fs.FolderExists(exportRoot)) Then
			'			Set f=fs.CreateFolder(exportRoot)
			'			Set f=nothing		
			'		End If		
			'		exportRoot = exportRoot & "\"
			'	Next
				'response.end
  Dim arr, dir, path
  Dim oFs

  Set oFs = CreateObject("Scripting.FileSystemObject")
  arr = split(filepath, "\")
  path = ""
  For Each dir In arr
    If path <> "" Then path = path & "\"
    path = path & dir
	'response.write(path)&"</br>"
    If not(Fs.FolderExists(path)) Then 
		set f=Fs.CreateFolder(path) 
		set f=nothing
	end if
  Next
				
				' Set Parameters
				Dim UTFStream, csvMeasure, csvDemand
				Set UTFStream = CreateObject("adodb.stream")
				UTFStream.Type = adTypeText
				UTFStream.Mode = adModeReadWrite
				UTFStream.Charset = "UTF-8"
				UTFStream.LineSeparator = adLF
				UTFStream.Open

				' CSV headers
				select case UtilityId
				case 1
					csvMeasure = "(Mlbs)"
					csvDemand = "(Mlbhrs)"
				case 2
					csvMeasure = "(kWh)"
					csvDemand = "(kW)"
				case 3, 22, 4, 10
					csvMeasure = "(cCF)"
					csvDemand = ""
				case else
					csvMeasure = ""
					csvDemand = ""
				end select
				csvHeaders =   chr(34) & "Building" & chr(34) & "," _
							 & chr(34) & "Tenant No." & chr(34) & "," _
							 & chr(34) & "Tenant Name" & chr(34) & "," _
							 & chr(34) & "Meter No." & chr(34) & "," _
							 & chr(34) & "From" & chr(34) & ","_
							 & chr(34) & "To" & chr(34) & ","_
							 & chr(34) & "No. of Days" & chr(34) & ","_
							 & chr(34) & "Tenant Total ($)" & chr(34) & ","_
							 & chr(34) & "Tenant Total Final ($)" & chr(34) & ","_
							 & chr(34) & "Current Usage " & csvMeasure & chr(34) & ","_
							 & chr(34) & "Previous Usage " & csvMeasure & chr(34) & ","_
							 & chr(34) & "% change Usage vs Last Month (% of cCF)" & chr(34) & ","_
							 & chr(34) & "Historical Average " & csvMeasure & chr(34) & ","_
							 & chr(34) & "% change Usage vs Historical Average (% of cCF)" & chr(34) & ","_
							 & chr(34) & "Yardi Unit Type" & chr(34) & ","_
							 & chr(34) & "Unit Average " & csvMeasure & chr(34) & ","_
							 & chr(34) & "% diff Usage vs Unit Type (% of cCF)" & chr(34) & ","_
							 & chr(34) & "Service Class" & chr(34) & ","_
							 & chr(34) & "Billid" & chr(34)
				UTFStream.WriteText csvHeaders 
				UTFStream.WriteText crlf
				
				if not putf then 
					PUTFStream.WriteText csvHeaders 
					PUTFStream.WriteText crlf
					putf = true
				end if
				'response.write filePath & csvFile & "</br>"
				
		'Zero out building peak numbers
		bldgonpeak 	= 0
		bldgoffpeak = 0
		bldgintpeak = 0
		if building = "96-02-57" then building = "RLC1LONA" end if


		Dim pagepart, calcintpeak, sql2, sInvoiceNo, extusageflag

		'If rst1("calcintpeak") and rst1("calcintpeak")="True" Then 
		'	calcintpeak = true 
		'Else 
		'	calcintpeak = false
		'End if

		extusageflag = false
		pagepart = 1

		Dim iRow , iCol
			iCol = 1
			iRow = 6 
		'Do Until rst1.eof
			totaldemand_PC 		= 0
			totaldemand_PCint 	= 0
			totaldemand_PCoff 	= 0
			totalOnpeak 		= 0
			totalOffPeak 		= 0
			totalIntPeak 		= 0
			totalKWH 			= 0
			totalKWHoff 		= 0
			totalKWHint 		= 0
			'extusage 			= rst1("extusg")
			'response.write rst1("billingname") & "</br>"
			if  extusage  then
				extusageflag = true
			end if

			'sql = "SELECT * FROM tblmetersbyperiod m, tblbillbyperiod b " & _
			'		" WHERE b.id=m.bill_id and b.reject=0 and m.leaseutilityid="&cdbl(rst1("leaseutilityid"))& _
			'		" and m.ypid="&cdbl(rst1("ypid"))&" ORDER BY meternum"
			sql = "Exec LefrakColdWaterReport '" & building & "', " & billyear & ", " & billperiod & ", " & utilityid '& ", " & rst1("ypid")
			'response.write sql & "</br>"
			'response.end

			rst2.open sql, cnn1

			
			Dim metercount, intpeak, extusage,tester, flag

			
			'extusage = rst1("extusg")
			metercount = 0
			'PAlngTenantCount=0
			tMeter = 0
			if rst2.eof then
			
			else
			
			do until rst2.eof
			
				metercount = metercount+1
				
				'tester = InStr(rst2("tenantnum"),"MA")
				
			   'response.write PACondition
			   
			   'response.write tMeter
			   'response.end
				'csv export
				'reDim preserve csvTenant
				set csvTenant = new tenant
				csvTenant.building = ucase(building)   										'building
				csvTenant.tNum = ucase(rst2("TenantNum"))   						't num
				csvTenant.tName = rst2("tenantname") 						't name
				csvTenant.mNum = ucase(rst2("Meternum")) 							'meter
				csvTenant.dFrom = rst2("DateStart") 		'from
				csvTenant.dTo = rst2("DateEnd") 			'to
				csvTenant.days = rst2("Days")								'Days	
				'csvTenant.blank1 = " " 										'blank	   
				'response.write "add:" & tMeter & " | " & csvTenant.tNum & " // " & csvTenant.mNum & "</br>"
				
				'intpeak = 0
				'if isnumeric(rst2("IntPeak")) then intpeak = cdbl(rst2("IntPeak"))
				'if rst2("coincident")="True" then
				'	meterdemandtemp = 0
				'	mDemand = cdbl(rst2("Demand_C")) + cdbl(rst2("Demand_int")) + cdbl(rst2("Demand_off"))
				'	totaldemand_PC = mDemand
				'else
				'	mDemand = cdbl(meterdemandtemp)
				'	totaldemand_PC = totaldemand_PC + mDemand
				'	if calcintpeak then
				'		totaldemand_PCoff = totaldemand_PCoff + cdbl(meterdemandtempoff)
				'		totaldemand_PCint = totaldemand_PCint + cdbl(meterdemandtempint)
				'	end if
				'end if
				'
				'totalOffPeak = totalOffPeak + formatnumber(rst2("OffPeak"),2)
				'
				'if calcintpeak then 
				'	totalIntPeak 	= totalIntPeak + formatnumber(rst2("IntPeak"),2)
				'	totalOnPeak 	= totalOnpeak + formatnumber(rst2("OnPeak"),2)
				'else
				'	totalOnpeak = totalOnpeak + formatnumber(cdbl(rst2("OnPeak"))+IntPeak,2)
				'end if
				'
				'if extusage then 
				'	if rst2("mextusg") then 
				'		mUsage = formatnumber(cdbl(rst2("used"))/usagedivisor,2)
				'		totalKWH 	= totalKWH + mUsage
				'		totalkwhoff = totalKWHoff + formatnumber(cdbl(rst2("Usedoff"))/usagedivisor,2)
				'		totalkwhint = totalKWHint + formatnumber(cdbl(rst2("Usedint"))/usagedivisor,2)
				'	else
				'		mUsage = formatnumber(cdbl(rst2("onpeak"))/usagedivisor,2)
				'		totalKWH 	= totalKWH + mUsage
				'		totalkwhoff = totalKWHoff + formatnumber(cdbl(rst2("offpeak"))/usagedivisor,2)
				'		totalkwhint = totalKWHint + formatnumber(cdbl(rst2("intpeak"))/usagedivisor,2)
				'	end if
				'else
					cUsage = formatnumber(cdbl(rst2("currentused")))'/usagedivisor,2) 
					pUsage = formatnumber(cdbl(rst2("previousused")))'/usagedivisor,2) 
					totalKWH 	= totalKWH + cUsage
				'end if 
				'
				'if extusage and rst2("mextusg") then 
				'	metercount = metercount + 2
				'	
				'Else
					
						
				'End If
				
					
				'csv export
				csvTenant.cUsage = cUsage 			'm usage
				'response.write "mu:"& musage &"...........t:"&tmeter&"</br>"
				csvTenant.mDemand = mDemand 			'm demand
				tMeter = tMeter + 1
				
			'	rst2.movenext
			'Loop

			'set csvTenant = new tenant
			
			
			
			if extusage then
				tUsage = formatnumber(totalKWH+totalkwhoff+totalkwhint,2)
				tDemand = formatnumber(totaldemand_PC+totaldemand_PCint+totaldemand_PCoff,2)
			else
				tUsage = formatnumber(totalKWH,2)
				if utilityid=2 or utilityid=1 or utilityid=6 then
					tDemand = formatnumber(totaldemand_PC,2)
				end if
			end if

			'if not rst1("ibsexempt") then
			'	'response.write csvTenant.tName & " is not ibsexmpt </br>"
			'	if extusage then 
			'		bldgOnPeak = bldgOnPeak + totalKWH
			'		bldgOffPeak = bldgOffPeak + totalkwhoff
			'		bldgIntPeak = bldgIntPeak + totalkwhint 
			'		bldgTotalPeak = bldgTotalPeak + totalKWH + totalkwhoff + totalkwhint
			'	else
			'		bldgOnPeak = bldgOnPeak + totalOnPeak
			'		if calcintpeak then 
			'			bldgIntPeak = bldgIntPeak + totalIntPeak 
			'		end if
			'		
			'		bldgOffPeak = bldgOffPeak + totalOffPeak		
			'		bldgTotalPeak = bldgTotalPeak + totalKWH 
			'	end if
			'	
			'		bldgTotalKWon = bldgTotalKWon + totaldemand_PC
			'		bldgTotalKWoff = bldgTotalKWoff + totaldemand_PCoff
			'		bldgTotalKWint = bldgTotalKWint + totaldemand_PCint
			'		tAdmin = (cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDbl(rst1("credit")))*cdbl(rst1("adminfee"))
			'		bldgAdmin = bldgAdmin + tAdmin
			'	if ucase(trim(rst1("type"))) = "LPLS2" then 
			'		tLpls = cdbl(rst1("rate_servicefee"))
			'		bldgAdmin = bldgAdmin + tLpls
			'		tAdmin = tAdmin + tLpls
			'	end if
			'			
			'	tService = cDbl(rst1("serviceFee"))
			'	bldgService = bldgService + tService
			'	tCredit = cDbl(rst1("credit"))
			'	bldgCredit = bldgCredit + tCredit
			'	bldgSubtotal = bldgSubtotal + cDbl(rst1("subtotal"))
			'	tTax = cDbl(rst1("tax"))
			'	bldgTax = bldgTax + tTax
			'	'PA condition goes here
			'	
				mTotal = rst2("totalamt")
			'	bldgTotalAmt = bldgTotalAmt + tTotal
			'	
			'	tSub = cDbl(rst1("energy"))+cDBL(rst1("demand"))
			'	subsubtotal = subsubtotal + tSub +cDBL(rst1("rateservicefee_dollar"))
			'elseif rst1("shadow")="False" then
			'	'response.write csvTenant.tName & " is not shadow</br>"
			'	if extusage then 
			'		'response.write "extusage:"
			'		ExmpOnPeak = ExmpOnPeak + totalKWH
			'		ExmpOffPeak = ExmpOffPeak + totalkwhoff
			'		ExmpIntPeak = ExmpIntPeak + totalkwhint 
			'		ExmpTotalPeak = formatnumber(ExmpTotalPeak + totalKWH + totalkwhoff + totalkwhint, 2)
			'	else
			'		'response.write "not extusage:"
			'		ExmpOnPeak = ExmpOnPeak + totalOnPeak
			'		if calcintpeak then 
			'			ExmpIntPeak = ExmpIntPeak + totalIntPeak 
			'		end if
			'		ExmpOffPeak = ExmpOffPeak + totalOffPeak		
			'		ExmpTotalPeak = ExmpTotalPeak + totalKWH 
			'	end if
			'	'response.write ExmpTotalPeak & "</br>"
			'	ExmpTotalKWon = ExmpTotalKWon + totaldemand_PC
			'	ExmpTotalKWoff = ExmpTotalKWoff + totaldemand_PCoff
			'	ExmpTotalKWint = ExmpTotalKWint + totaldemand_PCint
			'	ExmpTotalKW = ExmpTotalKW + totaldemand_PC + totaldemand_PCint + totaldemand_PCoff
			'	tAdmin = (cDbl(rst1("energy"))+cDBL(rst1("demand"))-cDbl(rst1("credit")))*cdbl(rst1("adminfee"))
			'	ExmpAdmin = ExmpAdmin + tAdmin
			'	if ucase(trim(rst1("type"))) = "LPLS2" then 
			'		tLpls = cdbl(rst1("rate_servicefee"))
			'		ExmpAdmin = ExmpAdmin + tLpls
			'	end if
			'	tService = cDbl(rst1("serviceFee"))
			'	ExmpService = ExmpService + tService
			'	tCredit = cDbl(rst1("credit"))
			'	ExmpCredit = ExmpCredit + tCredit
			'	ExmpSubtotal = ExmpSubtotal + cDbl(rst1("subtotal"))
			'	tTax = cDbl(rst1("tax"))
			'	ExmpTax = ExmpTax + tTax
			'	tSub = cDbl(rst1("energy"))+cDBL(rst1("demand"))+cDBL(rst1("rateservicefee_dollar"))
			'	Exmpsubsubtotal = Exmpsubsubtotal + tSub
			'	tTotal = cDbl(rst1("TotalAmt"))
			'	ExmpTotalAmt = ExmpTotalAmt + tTotal
			'	ExmpData = true
			'end if
			
			
			'csv export
			'csvTenant. = " " 										'blank	csvTenant(metercount, 15) = mRate 			'm rate
			tSub = formatcurrency(tSub,6)
			'response.write "tSub:"&tSub&"  tUsage:"&tUsage&"</br>"
			'if isnumeric(rst1("energydetail")) then
			'	bRate = cdbl(rst1("energydetail"))
			'else
			'	bRate = cdbl(rst1("akwhdisplay"))
			'end if
			'if isnumeric(rst1("demanddetail")) then
			'	brate = brate + cdbl(rst1("demanddetail"))
			'end if
			
			if metercount = 1 then 
				mRate = bRate
			else
				if tSub = 0 then 
					'response.end
					mRate = 0
				else
					mRate = tSub/tUsage
				end if
			end if
			tFees = tAdmin + tService
			csvTenant.mRate = formatcurrency(mRate, 4) 			'm Rate
			csvTenant.cUsage = cUsage 			't usage
			csvTenant.pUsage = pUsage 			't usage
			csvTenant.tDemand = tDemand	or 0		't demand
			csvTenant.tSub = formatcurrency(tSub)			't sub
			csvTenant.tFees = formatcurrency(tFees)			't fees
			csvTenant.tTax = formatcurrency(tTax)			't tax
			csvTenant.tCredit = formatcurrency(tCredit)			't credits
			csvTenant.mTotal = mTotal			't total
			csvTenant.mkwhvariance = rst2("mkwhvariance")
			csvTenant.hkwhvariance = rst2("hkwhvariance")
			csvTenant.unitcode = rst2("ucode")
			csvTenant.ucdiff = rst2("ucdiff")
			
			'	mSub = (csvTenant(i).mUsage * mRate)
			'	if utilityid = 3 then
			'	'	mSub = cdbl(rst1("energy")) + cdbl(rst1("demand"))
			'	end if
			'	csvTenant(i).mSub = formatcurrency(mSub) 			'm sub
			'	if tFees > 0 then
			'		mFees = tFees / tUsage * csvTenant(i).mUsage
			'	else
			'		mFees = 0
			'	end if
			'		csvTenant(i).mFees = formatcurrency(mFees)			'm fees
			'	if tTax > 0 then
			'		mTax = tTax / tUsage * csvTenant(i).mUsage
			'	else
			'		mTax = 0
			'	end if
			'	csvTenant(i).mTax = formatcurrency(mTax) 			'm tax
			'	mTotal = mSub + mFees + mTax
			'	csvTenant(i).mTotal = formatcurrency(mTotal) 			'm total
			'next
			
			'csvTenant.building = ucase(building)   										'building
			'csvTenant.tNum = ucase(rst1("TenantNum"))   						't num
			'csvTenant.tName = rst1("billingname") 						't name
			'csvTenant.dFrom = FormatDateTime(rst1("DateStart"),2) 		'from
			'csvTenant.dTo = FormatDateTime(rst1("DateEnd"),2) 			'to
			'csvTenant.days = rst1("Days")								'Days	
			
			'rst2.close
			
			'if ucase(trim(rst1("type")))="AVG COST 1" then 
			'	tClass = rst1("type") & " "  & formatnumber(rst1("AvgKWH"),6)
			'	
			'else
				tClass = rst2("rate")
				billid = rst2("billid")
			'	
			'end if
			'tFeeRate = formatpercent(rst1("AdminFee"),2)
			
			' csv export
			csvTenant.tClass = tClass 			't Class
			csvTenant.tFeeRate = tFeeRate			't fees rate
			csvTenant.billid = billid
			csvTenant.unitaverage = rst2("unitaverage")
			csvTenant.historicalaverage = rst2("historicalaverage")
			'


				'response.end
				
				'response.write "Meter Count:" & metercount & "</br>"
				dim meter, field, meterRow, tenantLines
				tenantLines = ""
				
				'for meter = 0 to metercount
					'response.write "read:" & meter & " | " & csvTenant.mNum & "</br>"
					'response.end
					meterRow = ""
					
					'if meter < metercount or metercount > 1 then
						meterRow = meterRow & chr(34) & csvTenant.building & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.tNum & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.tName & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.mNum & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.dFrom & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.dTo & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.days & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.mTotal & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.FTotal & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.cUsage & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.pUsage & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.mkwhvariance & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.historicalaverage & chr(34) & ","
					'end if
					
					'if meter = metercount then
						meterRow = meterRow & chr(34) & csvTenant.hkwhvariance & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.unitcode & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.unitaverage & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.ucdiff & chr(34) & ","
						meterRow = meterRow & chr(34) & csvTenant.tClass & chr(34)& ","
						meterRow = meterRow & chr(34) & csvTenant.billid & chr(34)
						meterRow = meterRow & crlf
					'end if
					
					'if metercount > 1 and meter < metercount then
						'meterRow = meterRow & crlf
					'end if
					
					tenantLines = tenantLines + meterRow
					'response.write meterRow & "</br>"'crlf
				'next
				'response.write"</br>"'(crlf)
				'UTFStream.WriteText crlf
				UTFStream.WriteText tenantLines
				PUTFStream.WriteText tenantLines
				
				rst2.movenext
			Loop
			end if
			rst2.close
	'			rst1.movenext
				'# Added by Tarun 07/10/2006
				'lngTenantCount = lngTenantCount + 1
				'# 
	'		loop
			'response.end
			'rst1.close
			PUTFStream.WriteText crlf
			UTFStream.Position = 3 'skip BOM

			Dim BinaryStream
			Set BinaryStream = CreateObject("adodb.stream")
			BinaryStream.Type = adTypeBinary
			BinaryStream.Mode = adModeReadWrite
			BinaryStream.Open

			'Strips BOM (first 3 bytes)
			UTFStream.CopyTo BinaryStream
			
			'UTFStream.SaveToFile "d:\temp\adodb-stream1.csv", adSaveCreateOverWrite
			UTFStream.Flush
			UTFStream.Close
				'response.write filePath & csvFile & "</br>"
			BinaryStream.SaveToFile filePath & csvFile, adSaveCreateOverWrite
			'BinaryStream.SaveToFile ftproot & dailyFile, adSaveCreateOverWrite
			BinaryStream.Flush
			BinaryStream.Close
			
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			If objFSO.FileExists(filePath & csvFile) Then 
				%>
				<p> Building report has been generated :
				<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://pdfmaker.genergyonline.com/pdfmaker/<%=pbpath%><%=csvFile%>"&"?dt="&ctime target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=csvFile%></b></a> 
				</p>
				<%
			end if
		'else
		'	rst1.close
		'end if	
		rst0.movenext
	loop
	rst0.close

	PUTFStream.Position = 3 'skip BOM
	Dim PBinaryStream
	Set PBinaryStream = CreateObject("adodb.stream")
	PBinaryStream.Type = adTypeBinary
	PBinaryStream.Mode = adModeReadWrite
	PBinaryStream.Open		
	PUTFStream.CopyTo PBinaryStream
	
	'UTFStream.SaveToFile "d:\temp\adodb-stream1.csv", adSaveCreateOverWrite
	PUTFStream.Flush
	PUTFStream.Close
		'response.write "</br>"& resfile
	'PBinaryStream.SaveToFile filePath & csvFile, adSaveCreateOverWrite
	PBinaryStream.SaveToFile resfile, adSaveCreateOverWrite
	PBinaryStream.Flush
	PBinaryStream.Close
	
	
	If objFSO.FileExists(resfile) Then 
		%>
		<p></P>
		<p> All Buildings report has been generated :
		<a style="font-family:arial;font-size:12;text-decoration:none;color:black;" href="http://pdfmaker.genergyonline.com/pdfmaker/<%=ppath%><%=pcsvFile%>"&"?dt="&ctime target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'black'"><b><%=pcsvFile%></b></a> 
		</p>
		<%
	Else
		%>
		<p>There has been an error while generating the requested file. Please try and generate the file again. If the error persists, contact Genergy IT department for assistance.</p>
		<%
	end if
	set objFSO = Nothing


	
	'set rstMailingList = Nothing
	set objExcelReport = Nothing
	set rst1 = Nothing
	set rst2 = Nothing
	set rst3 = Nothing
	set rst9 = Nothing
	set cnn1 = Nothing
	killExcel()
	else
		%> <p> No reports have been generated for <%=utilityname%>. </p> <%
	end if
%>	
	
<%
	function killExcel()
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
	end function
%>