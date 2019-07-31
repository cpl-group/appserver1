<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Cache-Control", "must-revalidate"

	'On Error Resume Next
	
	Response.AddHeader "content-disposition", "attachment;filename=reveditrpt.csv"
	Dim rs,outputRS,dsRS, SQL, cnn, utilityid,bperiod,byear,building,link,strsql,rmode,extusg,previousMeterid,tenantname,reccount,isposted,needAcceptButton
	Dim kwvartemp,kwhvartemp,kwhvarbold
	link = request.querystring ("link")
	utilityid =  request.querystring ("u")
	bperiod = request.querystring ("bp")
	byear = request.querystring ("by")
	building = request.querystring ("bldg")
	rmode = request.QueryString("r")

	if link = "s" then
		'Supervisor
		strsql = _
			"SELECT (case when (avgKWH=0 or avgKW=0 or kwhvarience > variance*100+8 or kwvarience > variance*100 or ((kwhOFFvarience > variance*100+8 or kwhINTvarience > variance*100+8 or kwOFFvarience > variance*100 or kwINTvarience > variance*100) and extusg=1) or AMTvarience > variance*100) then '0' else '1' end) as belowVarience, * FROM ("&vbcrlf&_
			"SELECT distinct m.meterid, m.meternum, m.extusg, m.variance, v.revdate, c.validate, c.svalidate, bbp.posted, m.bldgnum,c.[current], isNull(c.used,0) as kwhused, isNull(c.usedoff,0) as kwhoff, isNull(c.usedint,0) as kwhint, isNull(pd.demand,0) as demand, isNull(pd.demand_off,0) as demand_off, isNull(pd.demand_int,0) as demand_int, l.tenantnum, l.billingname, isnull(l.billingid,'') as billingid, isNull(bbp.totalamt,0) as totalamt, bbp.adminfee, bbp.sqft, v.biller, v.org_kwh, v.org_kw, case when bbp.sqft=0 then 0 else(bbp.demand/bbp.sqft)end as wsqft, lup.coincident,lup.coincident_peak, lup.leaseutilityid, isnull(cd.demand,0) as coindemand, isnull(avgKWH,0) as avgKWH, isnull(avgKWHoff,0) as avgKWHoff, isnull(avgKWHint,0) as avgKWHint, isNuLL(avgKW,0) as avgKW, isNuLL(avgKWoff,0) as avgKWoff, isNuLL(avgKWint,0) as avgKWint, isNuLL(avgAmt,0) as avgAmt, "&vbcrlf&_
			"isNull(case when isNull(avgKWH,0)=0 then '0' else abs((c.used - (isNull(avgKWH,0)))/isNull(avgKWH,0)*100) end, 0) as kwhvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWHoff,0)=0 then '0' else abs((c.usedoff - (isNull(avgKWHoff,0)))/isNull(avgKWHoff,0)*100) end, 0) as kwhOFFvarience, "&_
			"isNull(case when isNull(avgKWHint,0)=0 then '0' else abs((c.usedint - (isNull(avgKWHint,0)))/isNull(avgKWHint,0)*100) end, 0) as kwhINTvarience, "&_
			"isNull(case when isNull(avgKW,0)=0 then '0' else abs((pd.demand - (isNull(avgKW,0)))/isNull(avgKW,0)*100) end, 0) as kwvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWoff,0)=0 then '0' else abs((pd.demand_off - (isNull(avgKWoff,0)))/isNull(avgKWoff,0)*100) end, 0) as kwOFFvarience, "&_
			"isNull(case when isNull(avgKWint,0)=0 then '0' else abs((pd.demand_int - (isNull(avgKWint,0)))/isNull(avgKWint,0)*100) end, 0) as kwINTvarience, "&_
			"isNull(case when isNull(avgAmt,0)=0 then '0' else abs((bbp.totalamt - (isNull(avgAmt,0)))/isNull(avgAmt,0)*100) end, 0) as Amtvarience "&vbcrlf&_
			"FROM consumption c "&_
			"INNER JOIN meters m ON m.Meterid=c.Meterid "&_
			"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
			"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
			"INNER JOIN tblleases l on lup.billingid=l.billingid "&_
			"LEFT JOIN ("&_
			"SELECT c2.meterid, isNull(avg(used),0) as avgKWH, avg(usedoff) as avgKWHoff, avg(usedint) as avgKWHint, isNull(avg(demand),0) as avgKW, avg(demand_off) as avgKWoff, avg(demand_int) as avgKWint FROM consumption c2 INNER JOIN peakdemand d2 ON d2.meterid=c2.meterid and c2.billyear=d2.billyear and c2.billperiod=d2.billperiod WHERE ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)) GROUP BY c2.meterid "&_
			") CAvg ON CAvg.meterid=m.meterid "&_
			"LEFT JOIN ("&_
			"SELECT leaseutilityid, isNull(avg(totalamt),0) as avgAmt FROM tblbillbyperiod WHERE ((billyear="&byear&"-1 and billperiod>="&bperiod&"+9)or(billyear="&byear&" and billperiod<"&bperiod&" and billperiod>="&bperiod&"-3)) and reject=0 GROUP BY leaseutilityid "&_
			") BBAvg ON BBAvg.leaseutilityid=lup.leaseutilityid "&_
			"LEFT JOIN coincidentdemand cd on cd.leaseutilityid = lup.leaseutilityid and cd.billyear = c.billyear and cd.billperiod = c.billperiod "&_
			"LEFT JOIN tblbillbyperiod bbp on bbp.reject=0 and m.leaseutilityid=bbp.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod "&_
			"LEFT JOIN validation v on m.Meterid=v.Meterid and c.billyear=v.billyear and c.billperiod=v.billperiod "&_
			"WHERE c.billyear="&byear&" and c.billperiod="&bperiod&" and 	m.bldgnum='"&building&"' and lup.utility="&utilityid&" and Online='1' and leaseexpired=0 "&_
			") final ORDER BY belowVarience, billingname, meternum, revdate desc"
	else 'Biller
		strsql = _
			"SELECT (case when (avgKWH=0 or avgKW=0 or kwhvarience > variance*100+8 or kwvarience > variance*100 or ((kwhOFFvarience > variance*100+8 or kwhINTvarience > variance*100+8 or kwOFFvarience > variance*100 or kwINTvarience > variance*100) and extusg=1) or AMTvarience > variance*100) then '0' else '1' end) as belowVarience, * FROM ("&vbcrlf&_
			"SELECT Distinct m.meterid, m.meternum, m.extusg, m.variance, c.validate, isNull(bbp.totalamt,0) as totalamt, bbp.posted, c.svalidate, m.bldgnum,c.[current], isNull(c.used,0) as kwhused, isNull(c.usedoff,0) as kwhoff, isNull(c.usedint,0) as kwhint, isNull(pd.demand,0) as demand, isNull(pd.demand_off,0) as demand_off, isNull(pd.demand_int,0) as demand_int, l.tenantnum, isnull(l.billingname,'') as billingname, isnull(l.billingid,'') as billingid,lup.coincident,lup.coincident_peak, isnull(cd.demand,0) as coindemand,lup.leaseutilityid, isnull(avgKWH,0) as avgKWH, isnull(avgKWHoff,0) as avgKWHoff, isnull(avgKWHint,0) as avgKWHint, isNuLL(avgKW,0) as avgKW, isNuLL(avgKWoff,0) as avgKWoff, isNuLL(avgKWint,0) as avgKWint, isNuLL(avgAmt,0) as avgAmt, "&vbcrlf&_
			"isNull(case when isNull(avgKWH,0)=0 then '0' else abs((c.used - (isNull(avgKWH,0)))/isNull(avgKWH,0)*100) end, 0) as kwhvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWHoff,0)=0 then '0' else abs((c.usedoff - (isNull(avgKWHoff,0)))/isNull(avgKWHoff,0)*100) end, 0) as kwhOFFvarience, "&_
			"isNull(case when isNull(avgKWHint,0)=0 then '0' else abs((c.usedint - (isNull(avgKWHint,0)))/isNull(avgKWHint,0)*100) end, 0) as kwhINTvarience, "&_
			"isNull(case when isNull(avgKW,0)=0 then '0' else abs((pd.demand - (isNull(avgKW,0)))/isNull(avgKW,0)*100) end, 0) as kwvarience, "&vbcrlf&_
			"isNull(case when isNull(avgKWoff,0)=0 then '0' else abs((pd.demand_off - (isNull(avgKWoff,0)))/isNull(avgKWoff,0)*100) end, 0) as kwOFFvarience, "&_
			"isNull(case when isNull(avgKWint,0)=0 then '0' else abs((pd.demand_int - (isNull(avgKWint,0)))/isNull(avgKWint,0)*100) end, 0) as kwINTvarience, "&_
			"isNull(case when isNull(avgAmt,0)=0 then '0' else abs((bbp.totalamt - (isNull(avgAmt,0)))/isNull(avgAmt,0)*100) end, 0) as Amtvarience "&vbcrlf&_
			"FROM consumption c "&_
			"INNER JOIN meters m ON m.Meterid=c.Meterid "&_
			"INNER JOIN peakDemand pd on m.Meterid=pd.Meterid and c.billyear=pd.billyear and c.billperiod=pd.billperiod "&_
			"INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid "&_
			"INNER JOIN tblleases l on lup.billingid=l.billingid "&_
			"LEFT JOIN ("&_
			"SELECT c2.meterid, isNull(avg(used),0) as avgKWH, avg(usedoff) as avgKWHoff, avg(usedint) as avgKWHint, isNull(avg(demand),0) as avgKW, avg(demand_off) as avgKWoff, avg(demand_int) as avgKWint FROM consumption c2 INNER JOIN peakdemand d2 ON d2.meterid=c2.meterid and c2.billyear=d2.billyear and c2.billperiod=d2.billperiod WHERE ((d2.billyear="&byear&"-1 and d2.billperiod>="&bperiod&"+9)or(d2.billyear="&byear&" and d2.billperiod<"&bperiod&" and d2.billperiod>="&bperiod&"-3)) GROUP BY c2.meterid "&_
			") CAvg ON CAvg.meterid=m.meterid "&_
			"LEFT JOIN ("&_
			"SELECT leaseutilityid, isNull(avg(totalamt),0) as avgAmt FROM tblbillbyperiod WHERE ((billyear="&byear&"-1 and billperiod>="&bperiod&"+9)or(billyear="&byear&" and billperiod<"&bperiod&" and billperiod>="&bperiod&"-3)) and reject=0 GROUP BY leaseutilityid "&_
			") BBAvg ON BBAvg.leaseutilityid=lup.leaseutilityid "&_
			"LEFT JOIN coincidentdemand cd on cd.leaseutilityid = lup.leaseutilityid and cd.billyear = c.billyear and cd.billperiod = c.billperiod "&_
			"LEFT JOIN tblbillbyperiod bbp on bbp.reject=0 and m.leaseutilityid=bbp.leaseutilityid and c.billyear=bbp.billyear and c.billperiod=bbp.billperiod "&_
			"WHERE c.billyear="&byear&" and c.billperiod="&bperiod&" and 	m.bldgnum='"&building&"' and lup.utility="&utilityid&" and Online='1' and leaseexpired=0 "&_
			") final ORDER BY belowVarience, billingname, meternum"
	end if
   set cnn = server.createobject("ADODB.Connection")
	set rs = server.createobject("ADODB.Recordset")
	set outputRS = server.createobject("ADODB.Recordset")
	set dsRS = server.createobject("ADODB.Recordset")
	cnn.Open getLocalConnect(building)
	'create temporary record set
	outputRS.Fields.Append "Meter", adVarChar, 100
	outputRS.Fields.Append "Tenant_Name", adVarChar, 100
	outputRS.Fields.Append "Tenant_Number", adVarChar, 100
	outputRS.Fields.Append "Average_Usage", adVarChar, 100
	outputRS.Fields.Append "Current_Usage", adVarChar, 100
	outputRS.Fields.Append "Variance_KWH", adVarChar, 100
	outputRS.Fields.Append "Average_KW", adVarChar, 100
	outputRS.Fields.Append "Current_KW", adVarChar, 100
	outputRS.Fields.Append "Variance_KW", adVarChar, 100
	outputRS.Fields.Append "Load_Factor", adVarChar, 100
	outputRS.Fields.Append "Bill_Amount", adVarChar, 100
	outputRS.Fields.Append "Average_Amount", adVarChar, 100
	outputRS.Fields.Append "Variance_Amount", adVarChar, 100
	
		
	rs.open strsql, cnn
	if not(rs.eof) then
	'then there are meters
	
		previousMeterid = ""
		outputRS.Open
		do until rs.eof 
		
			
			if lcase(rs("extusg"))="true" then extusg = true else extusg = false
			
					reccount= reccount+ 1
					isposted = false
					if rs("posted")="True" then isposted = true else needAcceptButton = true
					
					if previousMeterid<>trim(rs("meterid")) AND (cint(rmode)=cint(rs("belowVarience"))) then
						
						kwvartemp = rs("kwvarience")
						if not isnumeric(trim(kwvartemp)) then kwvartemp = 0
						kwhvartemp = rs("kwhvarience")
						if not isnumeric(trim(kwhvartemp)) then kwhvartemp = 0
					
					tenantname = trim(rs("billingname"))
					
					outputRS.AddNew
					
					outputRS("Meter") = rs("meternum")
					outputRS("Tenant_Name") = replace(tenantname,","," ")
					outputRS("Tenant_Number") = "'"&trim(rs("tenantnum"))
					
					'average
					Dim average
					average = ""
					if extusg then average =  "On:" end if
					average = average & " " & trim(rs("avgKWH"))
					if extusg then 
					average = average & " " & "Off: " & trim(("AvgKWHoff")) & " Int: " & trim(rs("AvgKWHint"))
                    end if
					outputRS("Average_Usage") = average
					'---------average
					'Current_Usage
					Dim currUsage
					currUsage = ""
					if extusg then currUsage =  "On:" end if
					 currUsage = currUsage & " " & trim(rs("kwhused"))
					if extusg then
					currUsage = currUsage & " " & "Off: " & trim(rs("KWHoff")) & " Int: " &  trim(rs("KWHint"))
					end if
					outputRS("Current_Usage") = currUsage
					'---------Current_Usage
					'Variance
					Dim Variance
					Variance = ""
					if extusg then Variance =  "On:" end if
					 Variance = Variance & " " & trim(kwhvartemp)
					if extusg then
					Variance = Variance & " " & "Off: " & trim(rs("KWHoffvarience")) & " Int: " &  trim(rs("KWHintvarience"))
					end if

					outputRS("Variance_KWH") = Variance&"%"
					'---------Variance
					'Average_KW
					Dim Average_KW
					Average_KW=""
					if extusg then Average_KW = "On:" end if
					Average_KW = Average_KW & " " & trim(rs("AvgKW"))
					if lcase(trim(rs("coincident"))) = "true" or lcase(trim(rs("coincident_peak"))) = "true" then
						dim tempRST
						set tempRST = server.createobject("adodb.recordset")
						dim tempBPeriod1, tempBPeriod2, tempBYear1, tempBYear2
						tempBPeriod1 = bperiod - 1
						tempBYear1 = bYear
						if tempBPeriod1 <= 0 then
							tempBPeriod1 = tempBPeriod2 + 12
							tempBYear1 = tempBYear1 - 1
						end if
							tempBPeriod2 = tempBPeriod1 -1
							tempBYear2 = tempBYear1
						if tempBPeriod2 <= 0 then
							tempBPeriod2 = tempBPeriod2 + 12
							tempBYear2 = tempBYear2 - 1
						end if
						dim tempSQL
						tempSQL = "select isnull(avg(demand),0) as avgcoindemand from coincidentdemand where ((billyear= " & byear & " and billperiod = " & _
						bperiod & ") or  (billyear = " & tempbyear1 & " and billperiod = " & tempbperiod1 & ") or  (billyear = " & tempbyear2 & _
						" and billperiod = " & tempbperiod2 & ")) AND leaseutilityid = " & rs("leaseutilityid")
						tempRST.open tempSQL, cnn

						Average_KW = Average_KW & "Coincident " & trim(temprst("avgcoindemand"))
					end if
					if extusg then
					Average_KW = Average_KW & "Off: "  & trim(rs("AvgKWoff")) & " Int: " & trim(rs("AvgKWint"))
					end if
					outputRS("Average_KW") = Average_KW
					'---------Average_KW
					'Current_KW
					Dim Current_KW
					Current_KW = ""
					if extusg then Current_KW =  "On:"  end if
					Current_KW = Current_KW & " " & trim(rs("demand"))
					if lcase(trim(rs("coincident"))) = "true" or lcase(trim(rs("coincident_peak"))) = "true" then
						Current_KW = Current_KW & "Coincident Demand" & trim(rs("coindemand"))
					end if
					
					if extusg then
					Current_KW = Current_KW & " " & "Off: " & trim(rs("demand_off")) & " Int: " &  trim(rs("demand_int"))
					end if
					
					outputRS("Current_KW") = Current_KW
					'----------Current_KW
					'Variance_KW
					Dim Variance_KW
					Variance_KW = kwhvarbold
					if extusg then Variance_KW = Variance_KW &  " On: " 
					Variance_KW = Variance_KW & trim(kwvartemp) 
					if extusg then
					Variance_KW = Variance_KW & " Off: "  & trim(rs("KWoffvarience")) & " Int: "  & trim(rs("KWintvarience"))
					end if
	 				outputRS("Variance_KW") = Variance_KW & "%"
					'----------Variance_KW
					'Load_Factor
					Dim load_fact, dem, usage,dSql,dateStart,dateEnd,monthDays,monthHours
					dSql = "select datestart,dateend from billyrperiod where bldgnum = '"&building&"' and billperiod = "&bperiod&" and billyear = " & byear
					dim dates
					set dates = server.createobject("adodb.recordset")
					dates.Open dSql,cnn
					if not dates.eof then
					
					dateStart = Cdate(dates("datestart"))
					dateEnd = Cdate(dates("dateend"))
					
					end if
					'calculations
					monthDays = DateDiff("d",dateStart,dateEnd) + 1 ' to include the last day
					monthHours = monthDays * 24 'days in month times 24 hours
					'response.Write(monthDays)
					usage = Cdbl(rs("kwhused"))
					dem = Cdbl(rs("demand"))
					'formula: (total usage / total demand) / total hours in month
					if (dem = 0) then
					load_fact = 0
					else
					load_fact = (usage / dem) / monthHours
					end if
					
				'outputRS("Load_Factor") = "(" & usage & " / " & dem & ") / " & DateDiff("d",dateStart,dateEnd) + 1
					
					outputRS("Load_Factor") = trim(load_fact) & "%"
					'-----------Load_Factor
			
					
					
					outputRS("Bill_Amount") = "$"&trim(rs("totalamt"))
					outputRS("Average_Amount") = "$"&trim(rs("avgAmt"))
					outputRS("Variance_Amount") = trim(rs("Amtvarience")) & "%"
					
					
					
			
					previousMeterid=trim(rs("meterid"))
					end if

	rs.movenext
	loop

	rs.close
		outputRS.Updatebatch
		outputRS.Movefirst
		Dim F, Head
		For Each F In outputRS.Fields
		  Head = Head & ", " & replace(F.Name,"_"," ")
		Next
		Head = Mid(Head,3) & vbCrLf
	Response.ContentType = "text/plain"
	Response.Write Head
	Response.Write outputRS.GetString(,,", ",vbCrLf,"")
	outputRS.Close
	end if
	
	
	%>