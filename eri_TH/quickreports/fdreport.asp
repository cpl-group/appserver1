<%@Language="VBScript"%>
<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, bldgname, cnnLocal, rst, pid,testing
bldg 	= request("bldg")
'bldg 	= 'jfk'

'testing =  "driver={SQL Server};server=10.0.7.7;uid=sa;pwd=!general!;database=MainModule;"
set rst = server.createobject("adodb.recordset")
rst.open "select strt, portfolioid from buildings where bldgnum = '" & bldg & "'", getLocalConnect(bldg)
'response.write getLocalConnect(bldg)
'response.end
if rst.eof then
	response.write ("invalid building")
	response.End()
end if

bldgname = rst("strt")
pid = cint(rst("portfolioid"))
rst.close

dim monthp,yearp,utilityp
monthp = request("month")
if monthp = "" then
	monthp = month(date()) - 1
end if
yearp = request("year")
if yearp = "" then
	yearp = 1
end if
utilityp = request("utility")
if utilityp = "" then
	utilityp = 1
end if

%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		


</head>
<form name="fdhamparams" action="fdreport.asp" method="post">
<input type="hidden" value="<%=bldg%>" name="bldg">
<body bgcolor="#eeeeee" text="#000000">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
	<tr> 
		
    <td width="100%" bgcolor="#6699cc"><span class="standardheader">Demand Analysis 
      Report: <%=bldgname%></span></td>
	</tr>
	<tr bgcolor="#eeeeee"> 
		<td valign="top" style="border-bottom:1px solid #cccccc;">
			Select your year, month, and utility, then click View Report
		</td>
	</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
	<tr> 
		<td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" width="15%">
			Year&nbsp;&nbsp;&nbsp;
			<select name="year">
				<%
				rst.open "select distinct billyear from billyrperiod where bldgnum = '" & bldg & "' order by billyear desc", getLocalConnect(bldg)
				
				if not rst.eof then
					do while not rst.eof
						%><option value=<%=rst("billyear")%> <%if cint(rst("billyear")) = cint(yearp) then %> selected<%end if%>><%=rst("billyear")%></option><%
						rst.movenext
					loop
				end if
				rst.close
				%>
			</select>		
		</td>
		<td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" width="15%">
			Month&nbsp;&nbsp;&nbsp;
			<select name="month">
				<option value=1 <%if cint(monthp) = 1 then %> selected<%end if%>>January</option>
				<option value=2 <%if cint(monthp) = 2 then %> selected<%end if%>>February</option>
				<option value=3 <%if cint(monthp) = 3 then %> selected<%end if%>>March</option>
				<option value=4 <%if cint(monthp) = 4 then %> selected<%end if%>>April</option>
				<option value=5 <%if cint(monthp) = 5 then %> selected<%end if%>>May</option>
				<option value=6 <%if cint(monthp) = 6 then %> selected<%end if%>>June</option>
				<option value=7 <%if cint(monthp) = 7 then %> selected<%end if%>>July</option>
				<option value=8 <%if cint(monthp) = 8 then %> selected<%end if%>>August</option>
				<option value=9 <%if cint(monthp) = 9 then %> selected<%end if%>>September</option>
				<option value=10 <%if cint(monthp) = 10 then %> selected<%end if%>>October</option>
				<option value=11 <%if cint(monthp) = 11 then %> selected<%end if%>>November</option>
				<option value=12 <%if cint(monthp) = 12 then %> selected<%end if%>>December</option>
			</select>		
		</td>
		<td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" width="15%">
<!--			Utility&nbsp;&nbsp;&nbsp;
			<select name="utility">
				<%
'				dim sql
'				sql = "SELECT distinct tu.utilityid, tu.utilitydisplay FROM [" & application("superip") & "].mainmodule.dbo.tblutility tu INNER JOIN tblLeasesUtilityPrices tlup ON tlup.Utility = tu.utilityid INNER JOIN tblLeases tl ON tlup.BillingId = tl.BillingId WHERE tl.BldgNum = '" & bldg & "'"
	'			rst.open sql, getLocalConnect(bldg)
	'			if not rst.eof then
	'				do while not rst.eof
						%><option value=<%'=rst("utilityid")%> <%'if cint(rst("utilityid")) = cint(utilityp) then %> selected<%'end if%>><%'=rst("utilitydisplay")%></option><%
					'	rst.movenext
			'		loop
		'		end if
				%>
			</select>
			-->
			<input type="hidden" name="utility" value=2>		
		</td>
		<td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
			<input type="submit" name="action" value="View Report" onclick="document.getElementById('hideable').innerHTML = 'Loading...';">
		</td>
	</tr>
</table>
</form>
<div id="hideable">
<%
if request("action") = "" then 
	response.end
end if

dim startdate, enddate, utility
utility = cint(request("utility"))
startdate = Cdate(  request("month") & " - 1 - " & request("year") )

dim tempMonth, tempYear
tempMonth = (cint(request("month")) + 1)
tempYear =  cint(request("year"))
if tempMonth = 13 then 
	tempMonth = 1
	tempYear  = tempYear + 1
end if 
enddate = Cdate( tempMonth & " - 1 - " & tempYear) ' add one to the month, cast it to a date, subtract one day, and recast
%>
<table height="85%" width="100%" cellpadding="10" cellspacing="0">		<%'this table makes 2 big blocks with info centered in them.  the top one is the graph, bottom is the chart.%>
	<tr height="100%">
		<td width="1%"></td>
		<td align="center"  valign="top"  width="32%" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;">
			<table width="100%" cellpadding="0" cellspacing="0">
				<tr bgcolor="6699cc">
					<td width="25%" align="right" style="border-bottom:2px solid #eeeeee;">
						<span class="standardheader">Date</span></td>
					<td style="border-right:2px solid #ffffff;border-bottom:2px solid #eeeeee;">&nbsp;&nbsp;&nbsp;</td>
					<td width="30%" align="right" style="border-bottom:2px solid #eeeeee;">
						<span class="standardheader">Usage</span></td>
					<td style="border-right:2px solid #ffffff;border-bottom:2px solid #eeeeee;">&nbsp;&nbsp;&nbsp;</td>
					<td width="25%" align="right" style="border-bottom:2px solid #eeeeee;">
						<span class="standardheader">Peak Demand</span></td>
					<td style="border-right:2px solid #ffffff;border-bottom:2px solid #eeeeee;">&nbsp;&nbsp;&nbsp;</td>
					<td width="20%" align="right" style="border-bottom:2px solid #eeeeee;">
						<span class="standardheader">Time</span></td>
					<td style="border-bottom:2px solid #eeeeee;">&nbsp;&nbsp;&nbsp;</td>
				</tr>
			</table>
			<div style="height:90%; overflow:auto;" id="narf!">
				<table width="100%" cellpadding="0" cellspacing="0">
					<%
					
					Dim rid, billingid, meterid,  utilityid, groupname, interval

					rid = 0
					billingid = 0
					meterid = 0
					utilityid = cint(utility)
					groupname = 0
					interval = 2
					
					dim ip
					if trim(bldg)<>"" and trim(bldg)<>"0" and instr(bldg,"|")=0 then
					  ip = getBuildingIP(bldg)
					   
					
					else
					  ip = getPidIP(pid)
					  
					end if
					
					dim lmp, myHeight, myWidth
					
				    Set lmp = CreateObject("lmpchartFordam.lmpcontrol")
				  '   response.write bldg
					 'response.end                
					lmp.setLocalIP 0, ip
					lmp.utility = utilityid
					lmp.interval = interval
					lmp.loadaggs = false
					lmp.projectionSeries = -1
					lmp.SmallSize = false
					lmp.setCost 0, 0
					lmp.setCost 1, 0
																'building
					'response.write "0, ""bottom"", "&pid&", "&rid&", "&bldg&", "&billingid&", "&meterid&", "&groupname&", """&startdate&""", """&enddate&""""&"<br>"
					'response.write ip &"," &"utilityid"&","&interval&","
					'response.Write " sp_LMPDATA_V2 " & """"&startdate&""", """&enddate&""","&meterid&","&billingid&","&bldg&","&groupname&","&interval&"," & utilityid & ",0,0,0,0"
					'response.end
			     'response.Write("setSeries 0, ""top"","& pid&","& rid&","& bldg&","& billingid&","& meterid&","& groupname&"," &startdate&","& enddate)
				 'response.end
				'lmp.setSeries 0, "top", pid, rid, bldg, billingid, meterid, groupname, startdate, enddate
				lmp.setSeries 0, "bottom", pid, rid, bldg, billingid, meterid, groupname, startdate, enddate

					
					
					dim someNum
					dim usages
				'	response.write "<!--" &  lmp.getSeriesData(0) & "--><br>"
				'	response.write "<!--" &  lmp.getSeriesLabels(0) & "--><br>"
				'	response.write "<!--" &  lmp.getSeriesLabels(3) & "--><br>"
					'response.end
					if not isnull(lmp.getSeriesData(0)) then
						usages = split(lmp.getSeriesData(0),",")
						someNum = ubound(usages)
					else
						someNum = 0
					end if
					
					dim pds
					if not isnull( lmp.getSeriesLabels(0) )then	
						dim tempString
						tempString = lmp.getSeriesLabels(0)
						tempString = replace(tempString,"KW","",1,-1,1)		
						pds = split(tempString,",")
					end if
					
					dim times, tempTimes
					times = usages
					if not isnull( lmp.getSeriesLabels(3)) then
						tempTimes = split(lmp.getSeriesLabels(3),",")
					else
						tempTimes = Array("")
					end if
					
					' since the usages and pds will be zeroes when the day hasnt happened yet, but the times will be null, the times array
					' should be filled with N/As
					dim counter
					
					if trim(join(tempTimes,"")) <> "" then		' if there is any data in temp time array, fill time array with it and then dashes
						for counter = 0 to ubound(tempTimes)	' for whats left over
							times(counter) = tempTimes(counter)
						next
						for counter = (ubound(tempTimes) + 1) to ubound(times)
							times(counter) = "-"
						next
					else													' if there is no data, fill with just dashes
						for counter = 0 to ubound(times)
							times(counter) = "-"
						next
					end if
					
					dim maxPeak, maxPeakAmt
					maxPeak = 0
					maxPeakAmt = 0
							
					for counter = 0 to someNum
						dim timePeak
						timePeak = trim(times(counter))
						if timePeak<>"-" then 
							timePeak =  formatDateTime(timePeak, vbShortTime)
							'timePeak = cdate(timePeak)
						end if
						dim tempPd
						tempPd = pds(counter)
						if not isNumeric(tempPd) then tempPd = 0
						if clng(tempPd) > clng(maxPeakAmt) then
							maxPeakAmt = tempPd
							maxPeak = counter
						end if
						%>
						
						<tr bgcolor="ffffff" id="row<%=counter%>">
							<td width="25%" align="right">
								<%=request("month")%>/<%=counter+1%>/<%=request("year")%>
							</td>
							<td style="border-right:2px solid #eeeeee;">&nbsp;&nbsp;&nbsp;</td>
							<td width="30%" align="right"><%=usages(counter)%> KWH</td>
							<td style="border-right:2px solid #eeeeee;">&nbsp;&nbsp;&nbsp;</td>
							<td width="25%" align="right"><%=pds(counter)%> KW</td>
							<td style="border-right:2px solid #eeeeee;">&nbsp;&nbsp;&nbsp;</td>
							<td width="20%" align="right"><%=timePeak%></td>
							<td style="border-right:2px solid #eeeeee;">&nbsp;&nbsp;&nbsp;</td>
						</tr>
						<%
					next
					%>
				</table>
			</div>
			<table width="75%">
				<tr>
					<td width="10%" bgcolor="lightgreen">&nbsp;</td>
					<td>:  Peak demand for the month</td>
				</tr>
			</table>
						
			<script>
				document.getElementById("row<%=maxPeak%>").style.backgroundColor = 'lightgreen';
			</script>
							
		</td>
		<td align="center" valign="top" width="66%" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;">
			<img src="fdchart.asp?pid=<%=pid%>&building=<%=bldg%>&startdate=<%=startdate%>&enddate=<%=enddate%>">
		</td>
		<td width="1%"></td>
	</tr>
</table>
</div>
</body>
</html>
