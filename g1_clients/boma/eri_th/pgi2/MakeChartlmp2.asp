<%@Language=VBScript%>
<%Option Explicit

Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #INCLUDE VIRTUAL="/includes/ChartConst.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim objChart			'Dundas Chart 2D object
Dim Bldgname

dim ctr					'loop counter
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

Dim cnn1, rst1, strsql, bldg, cmd, prm
Dim meterid, lmpstart, lmpend,lmpdate, accountname, tenantmeter, ishourly, interval, chartTimeInterval,l, pulsetable, utility, usage, units, groupname, lmptype, lmpcode, billingid, luid, total, part, datasource,pid

bldg=Request.Querystring("bldg")
meterid=Request.QueryString("meterid")
lmpdate=Request.QueryString("startdate")
billingid = Request.QueryString("billingid")
interval=Request.QueryString("interval")
tenantmeter = request.querystring("tenantmeter")
utility = request.querystring("utility")
groupname = request("groupname")
pid = Request.QueryString("pid")

dim Totalunits, unitsdem, peaklabel
Totalunits = ""
select case cint(utility)
case 1
	units = "Demand (Mlbs/Hr)"
	Totalunits = "Total Usage (Mlbs):"
	peaklabel = "Peak demand: "
case 2
	units = "Demand (KW)"
    unitsdem = "KHW"
    Totalunits = "Total Usage (KWH): "
	peaklabel = "Peak demand: "
case 3, 10
	units = "Usage (CF)"
	Totalunits = "Total Usage (CF):"
	peaklabel = "Peak demand: "
case 4
	units = "Usage (CF)"
	Totalunits = "Total Usage (CF):"
	peaklabel = "Peak demand: "
case 6
	units = "Capacity (Tons)"
	Totalunits = "Total Usage (Tons/hr): "
	peaklabel = "Peak demand: "
case 9
	units = "KVA"
	Totalunits = "Total Usage (KVARH): "
	peaklabel = "Peak demand: "
case 13
	peaklabel = "Max Temperature: "
case 14
	peaklabel = "Max 15' CFM: "
case 15
	peaklabel = "Max Humidity: "
case 17
	units = "Amps"
case else
	units = ""
	TotalUnits = "Total Usage/Hour: "
	peaklabel = "Peak demand: "
end select

if trim(interval)="" then interval=0
if int(interval)=0 then
  ishourly = false
  chartTimeInterval = 15
elseif interval=1 then 
	ishourly = true
	chartTimeInterval = 100
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.command")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.open getConnect(pid,bldg,"billing") 

cnn1.CursorLocation = adUseClient

rst1.open "SELECT * FROM tblutility WHERE utilityid="&trim(utility), getConnect(pid,bldg,"dbCore")
if not rst1.eof then usage = trim(rst1("measure")) else usage = "kwh"
rst1.close
if cint(utility) = 17 then
	usage="pulse"
end if

dim displaydate, chartname
displaydate = weekdayname(weekday(lmpdate)) &" "& day(lmpdate) &", "& Monthname(Month(lmpdate)) &" "& Year(lmpdate)
dim graphtypename

if trim(meterid)<>"" then
    graphtypename = "Meter "
    rst1.open "SELECT meternum FROM meters WHERE meterid="&meterid, cnn1
    if not(rst1.eof) then
        accountname = rst1("meternum")
    end if
    rst1.close
    lmptype="m"
    lmpcode = meterid
elseif trim(billingid)<>"" then
    graphtypename = "Tenant "
    rst1.open "SELECT BillingName, leaseutilityid FROM tblLeases l , tblleasesutilityprices lup WHERE l.billingid=lup.billingid AND lup.utility="&utility&" AND l.billingId="&billingid, cnn1
    if not(rst1.eof) then
        accountname = rst1("BillingName")
        luid = cint(rst1("leaseutilityid"))
    end if
    rst1.close
    lmpcode=luid
    lmptype="L"
elseif trim(bldg)<>"" then
    graphtypename = "Building "
    rst1.open "SELECT strt FROM buildings WHERE bldgnum='"&bldg&"'", cnn1
    if not(rst1.eof) then
        accountname = rst1("strt")
    end if
    rst1.close
    lmpcode=bldg
    lmptype="b"
elseif trim(pid) <> "" then 
	graphtypename = "Portfolio "

    rst1.open "SELECT name FROM portfolio WHERE id='"&pid&"'", cnn1
    if not(rst1.eof) then
        accountname = rst1("name")
    end if
    rst1.close
    lmpcode=pid
    lmptype="p"
end if

dim currentkwh
currentkwh = 0

Dim KWH
kwh = 0

dim tempkwh, lastkwh, workinghour, workingday, temphour, tempday, temptime, chartTime, peak, label, hasdata, PeakDemand, pointcolor
peak = 0
chartTime = 0


if trim(groupname)="" then
    cmd.ActiveConnection = cnn1


	if trim(pid) = "" then 

		cmd.CommandText = "sp_LMPDATA"
		cmd.CommandType = adCmdStoredProc
		Set prm = cmd.CreateParameter("from", adVarChar, adParamInput, 20)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("to", adVarChar, adParamInput, 20)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("code", adVarChar, adParamInput, 1)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("string", adVarChar, adParamInput, 30)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("utility", adInteger, adParamInput, 2)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("interval", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("title", adVarChar, adParamOutPut, 30)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("max", adDouble, adParamOutPut, 18,2)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("sum", adDouble, adParamOutPut, 18,2)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("peakdemand", adDouble, adParamOutPut)
		cmd.Parameters.Append prm
		
		cmd.Parameters("from")		= lmpdate
		cmd.Parameters("to")		= lmpdate + " 23:45:00"
		cmd.Parameters("code")		= lmptype
		cmd.Parameters("string")		= lmpcode
		cmd.Parameters("utility")		= utility
		cmd.Parameters("interval")		= interval
		'response.write "exec sp_LMPDATA '"&cmd.Parameters("from")&"','"&cmd.Parameters("to")&"','"&cmd.Parameters("code")&"','"&cmd.Parameters("string")&"',"&cmd.Parameters("utility")&","&cmd.Parameters("interval")&",0,0,0<br>"
		'response.write cmd.activeconnection
		'response.end
		set rst1 = cmd.execute
		peakdemand = cmd.Parameters("peakdemand")
		'usage = cmd.Parameters("title")
		if not rst1.eof then
			do while charttime < 2400
				pointcolor = RGB(40,120,255)
				if chartTime mod 100 = 0 then label = chartTime else label = ""
				if not rst1.eof then
					if hour(rst1("date"))*100+minute(rst1("date"))=charttime then
						tempkwh = cDbl(rst1(usage))
						if trim(rst1("est")) = "1" or trim(rst1("est")) = "True" then pointcolor = RGB(153, 52, 0)
						rst1.movenext
					else
						tempkwh = 0
					end if
				else
					tempkwh = 0
				end if
				objChart.AddData tempkwh, 0, label, pointcolor
				kwh = kwh + tempkwh
				chartTime = chartTime + chartTimeInterval
				if chartTime mod 100 > 59 then chartTime = chartTime + 100-(chartTime mod 100)
			loop
		  hasdata = true
		 if not(isnull(cmd.Parameters("sum"))) then currentkwh = cdbl(cmd.Parameters("sum"))
		else
		end if
		rst1.close
 
	else

		cmd.CommandText = "sp_PLMPDATA"
		cmd.CommandType = adCmdStoredProc
		Set prm = cmd.CreateParameter("from", adVarChar, adParamInput, 12)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("to", adVarChar, adParamInput, 12)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("utility", adInteger, adParamInput, 2)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("interval", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("pid", adInteger, adParamInput)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("title", adVarChar, adParamOutPut, 30)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("max", adDouble, adParamOutPut, 18,2)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("sum", adDouble, adParamOutPut, 18,2)
		cmd.Parameters.Append prm
		Set prm = cmd.CreateParameter("pdemand", adDouble, adParamOutPut)
		cmd.Parameters.Append prm
		
		cmd.Parameters("from")			= lmpdate
		cmd.Parameters("to")			= dateadd("d",1,lmpdate)
		cmd.Parameters("utility")		= utility
		cmd.Parameters("interval")		= interval
		cmd.Parameters("pid")			= pid
		set rst1 = cmd.execute
		peakdemand = cmd.Parameters("pdemand")
		if not rst1.eof then
			pointcolor = RGB(40,120,255)
			do while charttime < 2400
			
			if chartTime mod 100 = 0 then label = chartTime else label = ""
			if not rst1.eof then
				tempkwh = cDbl(rst1(usage))
				rst1.movenext
			end if
			objChart.AddData tempkwh, 0, label, pointcolor
			kwh = kwh + tempkwh
				chartTime = chartTime + chartTimeInterval
				if chartTime mod 100 > 59 then chartTime = chartTime + 100-(chartTime mod 100)
			tempkwh=0	
			loop
		  hasdata = true
		end if
		if not(isnull(cmd.Parameters("sum"))) then 
			currentkwh = cdbl(cmd.Parameters("sum"))
		end if
		rst1.close
	end if
	if not(isnumeric(peakdemand)) then peakdemand = 0
else
      hasdata = false
end if



if lmptype = "L" then 
   if len(accountname)>25 then chartname = left(accountname,24) & "..., Aggregated" else chartname = accountname & ", Aggregated"
elseif lmptype = "m" Then
   chartname = accountname
elseif lmptype = "b" then
   if len(accountname)>25 then chartname = left(accountname,24) & "..." else chartname = accountname
elseif lmptype ="p" then 
   if len(accountname)>25 then chartname = left(accountname,24) & "..." else chartname = accountname
end if

'################ next graph
strsql = ""
if trim(pid) = "" then 
	if luid = "" and meterid = "" then
	  strsql = "select t.billyear,t.billperiod,t.consumption,t.demand from (Select top 13 billyear, billperiod, sum(totalkwh) as Consumption, sum(totalkw) as Demand from utilitybill, billyrperiod where billyrperiod.ypid=utilitybill.ypid and billyrperiod.bldgnum='"&bldg&"' group by billyear, billperiod order by billyear desc, billperiod desc) as t order by t.billyear,t.billperiod"
	elseif meterid = "" then
    strsql = "select m.billyear,m.billperiod,sum(demand_p) as Demand,sum(used) as Consumption from tblmetersbyperiod m, tblbillbyperiod b where b.id=m.bill_id and b.reject=0 and m.leaseutilityid=" & luid & " and m.billyear='"& year(lmpdate) & "' group by m.billyear, m.billperiod"
  else
    strsql = "select b.billyear, b.billperiod,demand_p as Demand,used as Consumption from tblmetersbyperiod m, tblbillbyperiod b where b.id=m.bill_id and b.reject=0 and meterid='"&meterid&"' and m.billyear='"&year(lmpdate)&"'"
	end if
	dim hasdemanddata
	hasdemanddata = false
	rst1.Open strsql, cnn1, adOpenStatic
	If not rst1.EOF then 
	hasdemanddata = true
	Dim numRecords, ArrYearlyKWH(), ArrYearlyKW(), index
	numRecords = rst1.RecordCount
	ReDim ArrYearlyKWH(numRecords)
	ReDim ArrYearlyKW(numRecords)
	
	for index=1 to numRecords
		ArrYearlyKWH(index) = clng(rst1("Consumption"))
		ArrYearlyKW(index)= clng(rst1("Demand"))
		rst1.movenext					
	next
	rst1.movefirst
	
	for ctr=1 to numRecords
		objChart.AddData ArrYearlyKWH(ctr), 2, rst1("billperiod") & "/" & rst1("billyear"),RGB(40,120,255)	
		objChart.AddData ArrYearlyKW(ctr), 3, rst1("billperiod") & "/" & rst1("billyear"),RGB(40,120,255)	
		rst1.movenext
	next
	
  if trim(utility)="2" then objChart.Legend.Enabled = true
  objChart.AddStaticText unitsdem,587  ,225,RGB(100,100,100),"Arial",8,0,2,90
	end if
	rst1.close
end if


'######################################### has data ######################################### 
if hasdata then
if utility<>2 then hasdemanddata=false
if hasdemanddata then
  objChart.ChartArea(1).AddChart column_CHART, 3, 3, ,0  'second chart,fourth series\
  objChart.ChartArea(1).AddChart line_CHART, 2, 2	, , 2'second chart,fourth series
  objChart.AddStaticText "Monthly Billed Consumption / Demand",55,175,RGB(100,100,100),"Arial",8,1
end if
'setup the chart types
objChart.ChartArea(0).AddChart LINE_CHART, 0, 0	'first chart,first series
objChart.Rectangle3DEffect()

'chart colors
objChart.SetSeriesColor 0, RGB(40,120,255)	'first series, light blue
'objChart.SetSeriesColor 2, RGB(150,0,150)	'third series, purple
objChart.SetSeriesColor 2, RGB(0,0,0)		'fourth series, black
objChart.SetSeriesColor 3, RGB(40,120,255)	
objChart.SetSeriesColor 1, RGB(0,0,0)		'second series, black

'setup background colors
objChart.BackgroundColor= RGB(255,255,255)
objChart.ChartArea(0).BackgroundColor = RGB(232,232,232)

'chart grid line colors
objChart.ChartArea(0).GridColor = rgb(200,200,200)
objChart.ChartArea(1).GridColor = rgb(200,200,200)
objChart.ChartArea(0).LineWidth = 2
objChart.ChartArea(1).LineWidth = 2
objChart.ChartArea(1).BackgroundColor = RGB(232,232,232)


'Set up the chart axis for the second chart
objChart.ChartArea(0).Axis(1).Angle = 45
'objChart.ChartArea(0).Axis(1).Interval= 15
objChart.ChartArea(0).Axis(2).Enabled=true
objChart.ChartArea(1).Axis(0).SetNumberFormat 1,0
if peak<5 then
  objChart.ChartArea(0).Axis(0).SetNumberFormat 1,2
  objChart.ChartArea(0).Axis(2).SetNumberFormat 1,2
else
  objChart.ChartArea(0).Axis(0).SetNumberFormat 1,0
  objChart.ChartArea(0).Axis(2).SetNumberFormat 1,0
end if
objChart.ChartArea(1).Axis(2).Enabled=true
objChart.ChartArea(1).Axis(2).SetNumberFormat 1,0


'Adjust the position and size of the ChartArea so that it matches the background picture.
objChart.ChartArea(0).SetPosition 55, 20, 525, 130
objChart.ChartArea(1).SetPosition 55, 190, 525,260

'add text to the chart
objChart.AddStaticText displaydate,55,3,RGB(100,100,100),"Arial",8,1
'objChart.AddStaticText graphtypename &"Profile",526,3,RGB(100,100,100),"Arial",8,1,1
objChart.AddStaticText chartname,526,3,RGB(100,100,100),"Arial",8,1,1
objChart.AddStaticText units,0,155,RGB(100,100,100),"Arial",8,0,2,90
if utility<>13 and utility<>14 and utility<>15 and utility<>17 then objChart.AddStaticText Totalunits&formatnumber(CurrentKWH,0),58,20,RGB(100,100,100),"Arial",8,1
if utility<>9 and utility<>4 and utility<>17 then objChart.AddStaticText peaklabel&formatnumber(Peakdemand,2)&" "&mid(units, instr(units,")")+1),58,32,RGB(100,100,100),"Arial",8,1
'objChart.AddStaticText "Energy (KWH)",588,155,RGB(100,100,100),"Arial",8,0,2,270
objChart.AddStaticText "Hours",300,165,RGB(100,100,100),"Arial",8,0,2
'objChart.AddStaticText "Billing Period",300,295,RGB(100,100,100),"Arial",8,0,2
if trim(billingid)<>"" and trim(meterid)="" then
    total = 0
    part = 0
    strsql = "SELECT count(meterid), datasource FROM meters m, tblleasesutilityprices lup WHERE m.leaseutilityid=lup.leaseutilityid and lup.billingid="&billingid&" and lup.utility="&utility&"and m.lmp = 0 GROUP BY datasource "
    rst1.open strsql, cnn1
    if not rst1.eof then
      total = rst1(0)
      datasource = rst1(1)
    end if
    rst1.close
    rst1.open "SELECT * FROM sysobjects WHERE name='"&datasource&"' and xtype='U'", cnn1
    if rst1.eof then datasource = ""
    rst1.close
    if trim(datasource)<>"" then
      strsql = "SELECT count(distinct meterid) FROM ["&datasource&"] p WHERE meterid in (SELECT meterid FROM meters m, tblleasesutilityprices lup WHERE m.leaseutilityid=lup.leaseutilityid and lup.billingid="&billingid&" and lup.utility="&utility&" and m.lmp=0)"
      rst1.open strsql, cnn1
      part = cint(rst1(0))+part
      rst1.close
    end if
    objChart.AddStaticText "*Aggregated view is made up of "&part&" out of "&total&" meters",0,296,RGB(100,100,100),"Arial",7,0,0
end if

objChart.SetColorFromPoint 0
objChart.Legend.FontSize = 8
objChart.Legend.Add "KW", RGB(40,120,255)
objChart.Legend.Add "KWH", RGB(0,0,0)
objChart.Legend.BorderColor = RGB(110,0,0)
objChart.Legend.BackgroundColor = RGB(230,230,230)
objChart.Legend.FontColor = RGB(0,0,110)
objChart.Legend.FontSize = 8
objChart.Legend.Transparent = true 'set to false, so that the background color can be seen
objChart.Legend.SetPosition 550,287,600,310 
'######################################### no  data ######################################### 
else
  objChart.AddStaticText "No "& graphtypename &"Profile Data Available for "& displaydate,300,3,RGB(0,0,100),"Arial",13,1,2
end if


'Return the chart as a JPEG
objChart.SendJpeg 600,310,50

'cleanup
set objChart = nothing
%>