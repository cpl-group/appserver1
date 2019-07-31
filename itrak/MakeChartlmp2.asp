<%@Language=VBScript%>
<%Option Explicit

Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!-- #INCLUDE file ="ChartConst.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->

<%
dim objChart			'Dundas Chart 2D object
Dim Bldgname

dim ctr					'loop counter
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

Dim cnn1, rst1, strsql, bldg
Dim meterid, lmpstart, lmpend,lmpdate, luid, lmp, tenantname, tenantmeter, ishourly, modulus, chartTimeInterval

bldg=Request.Querystring("b")
meterid=Request.QueryString("m")
lmpstart=Request.QueryString("s")
lmpend=Request.QueryString("e")
lmpdate=Request.QueryString("d")
luid = Request.QueryString("luid")
lmp = Request.QueryString("lmp")
modulus=Request.QueryString("i")
tenantmeter = request.querystring("tenantmeter")

if trim(modulus)="" then modulus=0
ishourly = false
chartTimeInterval = 15
if modulus=100 then 
	ishourly = true
	chartTimeInterval = 100
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy1")

dim currentkwh
currentkwh = 0


dim displaydate, chartname
displaydate = weekdayname(weekday(lmpdate)) &" "& day(lmpdate) &", "& Monthname(Month(lmpdate)) &" "& Year(lmpdate)
dim graphtypename
if trim(luid)<>"" then
    graphtypename = "Tenant "
    rst1.open "SELECT BillingName FROM tblLeases WHERE billingId in (SELECT billingid FROM tblLeasesUtilityPrices WHERE leaseutilityid="& luid &")", cnn1
    if not(rst1.eof) then
        tenantname = rst1("BillingName")
        rst1.close
    end if
elseif trim(lmp)<>"" then
    graphtypename = "Building "
elseif trim(meterid)<>"" then
    graphtypename = "Meter "
else
    graphtypename = " "
end if

if luid <> "" and tenantmeter<>"1" then 
	strsql = "select (datepart(hour,p.date)*100)+datepart(minute,p.date) as time, sum(kwh) as kwh, datepart(hour,p.date) as hour, datepart(day,p.date) as day from pulse_"& bldg &" p where meterid in (select meterid from meters where LeaseUtilityId='" & luid & "' and pp<>1) and left(date,11) = convert(datetime,'" & lmpdate & "',101) AND ((datepart(hour,p.date)*100)+datepart(minute,p.date) BETWEEN '" & lmpstart & "' AND '" & lmpend & "') group by datepart(hour,date), [date] order by date"
else
    strsql = "SELECT Meters.MeterNum, Strt, Meters.MeterId, p.date as date, (datepart(hour,p.date)*100)+datepart(minute,p.date) as time, (meters.multiplier) as multiplier, p.kwh, datepart(hour,p.date) as hour, datepart(day,p.date) as day  FROM (pulse_"& bldg &" p INNER JOIN Meters ON p.meterid = Meters.MeterId) INNER JOIN buildings on buildings.BldgNum=Meters.bldgnum  WHERE (left(p.date,11)) = convert(datetime,'" & lmpdate & "',101) AND (p.meterid = '" & meterid & "') AND ((datepart(hour,p.date)*100)+datepart(minute,p.date) BETWEEN '" & lmpstart & "' AND '" & lmpend & "')  ORDER BY date"
end if
'response.write strsql
'response.end

rst1.Open strsql, cnn1, adOpenStatic

if not rst1.eof then 
	'to find table title
	if luid <> "" and tenantmeter<>"1" then 
	    chartname = tenantname & ", Aggregated"
	elseif tenantname<>"" Then
	    chartname = tenantname &", "& rst1("MeterNum")
	else
	    chartname = rst1("Strt") &", "& rst1("MeterNum")
	end if

Dim KWH
kwh = 0

dim tempkwh, lastkwh, workinghour, workingday, temphour, tempday, temptime, chartTime
chartTime = 0
do until rst1.eof
	if chartTime <> cInt(rst1("time")) then
		do until chartTime = cInt(rst1("time")) or charttime > 2400 
			if chartTime mod 100 = 0 then 
				objChart.AddData 0, 0, chartTime
			elseif not(ishourly) then
				objChart.AddData 0, 0
			end if
			chartTime = chartTime + 15
			if chartTime mod 100 > 59 then chartTime = chartTime + 100-(chartTime mod 100)
'			response.write cInt(rst1("time"))&chartTime&"<br>"
		loop
	end if
	tempkwh=0
	if ishourly then 'calculates hour peak demand
		lastkwh=0
		tempkwh=0
		workinghour = cint(rst1("hour"))
		workingday = cint(rst1("day"))
		temphour = workinghour
		tempday = workingday
		do until rst1.eof or workingday<>tempday or workinghour<>temphour
			if trim(rst1("kwh"))<>"" then 
				kwh = kwh + clng(rst1("kwh"))
				if lastkwh+clng(rst1("kwh"))> tempkwh then tempkwh = lastkwh+clng(rst1("kwh"))
				lastkwh = cDbl(rst1("kwh"))
			end if
			rst1.movenext
			if not rst1.eof then
				tempday = cint(rst1("day"))
				temphour = cint(rst1("hour"))
			else
				tempday = null
				temphour = null
			end if
		loop
		tempkwh=tempkwh*2
		objChart.AddData tempkwh, 0, workinghour*100
	else
		temptime = cint(rst1("time"))
		if trim(rst1("kwh"))<>"" then 
			tempkwh = cDbl(rst1("kwh"))
			kwh = kwh + clng(rst1("kwh"))
		end if
		if temptime mod 100 = 0 then 
			objChart.AddData tempkwh, 0, temptime
		else
			objChart.AddData tempkwh, 0
		end if
		rst1.movenext
	end if
	chartTime = chartTime + 15
	if chartTime mod 100 > 59 then chartTime = chartTime + 100-(chartTime mod 100)
loop
rst1.close
currentkwh = kwh






'################ next graph
if lmp ="" then 
	if luid = "" then
		strsql = "select * from (select  top 12 billyear, billperiod, demand_p as Demand, KWHUsed as Consumption  from tblmetersbyperiod where meterid = " & meterid & " order by id desc) as cat order by cat.billyear, cat.billperiod "
	else
		strsql = "select billyear, billperiod, sum(demand) as Demand, sum(energy) as Consumption from tblbillbyperiod where leaseutilityid ='" & luid & "' and billyear ='"& year(lmpdate) & "' group by billyear, billperiod"
	end if
else
	strsql = "select * from (Select top 12 billyear, billperiod, totalkwh as Consumption, totalkw as Demand from utilitybill, billyrperiod where billyrperiod.ypid=utilitybill.ypid and billyrperiod.bldgnum='" & bldg & "'  order by  id desc) as cat order by cat.billyear, cat.billperiod"

end if

'response.write strsql
'response.end


rst1.Open strsql, cnn1, adOpenStatic
If not rst1.EOF then 

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

'response.write displaydate
'response.end

'objChart.ChartArea(1).AddChart column_CHART, 3, 3, ,0  'second chart,fourth series\
'objChart.ChartArea(1).AddChart line_CHART, 2, 2	, , 2'second chart,fourth series
'objChart.AddStaticText "Monthly Billed Consumption / Demand",55,175,RGB(100,100,100),"Arial",8,1


end if
rst1.close


'setup the chart types
objChart.ChartArea(0).AddChart column_CHART, 0, 0	'first chart,first series
'objChart.ChartArea(0).AddChart LINE_CHART, 1, 1, , 2 'first chart,second series
objChart.Rectangle3DEffect()

'objChart.ChartArea(0).AddChart LINE_CHART, 2, 2 'first chart,third series


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
objChart.ChartArea(0).Axis(0).SetNumberFormat 1,0
objChart.ChartArea(0).Axis(2).SetNumberFormat 1,0
objChart.ChartArea(1).Axis(2).Enabled=true
objChart.ChartArea(1).Axis(2).SetNumberFormat 1,0


'Adjust the position and size of the ChartArea so that it matches the background picture.
objChart.ChartArea(0).SetPosition 55, 20, 500, 130
objChart.ChartArea(1).SetPosition 55, 190, 500,260

'add text to the chart
objChart.AddStaticText "Load Profile for " & displaydate,55,3,RGB(100,100,100),"Arial",8,1
'objChart.AddStaticText graphtypename &"Profile",526,3,RGB(100,100,100),"Arial",8,1,1
objChart.AddStaticText chartname,526,3,RGB(100,100,100),"Arial",8,1,1
objChart.AddStaticText "Demand (KW)",0,100,RGB(100,100,100),"Arial",8,0,2,90
objChart.AddStaticText "Total KWH: "&formatnumber(CurrentKWH,0),58,20,RGB(100,100,100),"Arial",8,1
objChart.AddStaticText "Energy (KWH)",545,100,RGB(100,100,100),"Arial",8,0,2,270
objChart.AddStaticText "Hours",300,165,RGB(100,100,100),"Arial",8,0,2
'objChart.AddStaticText "Billing Period",300,195,RGB(100,100,100),"Arial",8,0,2
if luid<>"" and tenantmeter<>"1" then
    strsql = "SELECT (select count(DISTINCT meterid) from pulse_"& bldg &"h where date ='"& lmpdate &"' and meterid IN (SELECT meterid FROM meters WHERE LeaseUtilityId="& luid &" and pp<>1)) as part, (SELECT count(meterid) from meters where (LeaseUtilityId="& luid &" and online=1 and lmnum is not NULL and pp<>1) or (leaseUtilityId="& luid &" and online=1 and EXISTS (select * from tblLeasesUtilityPrices where LeaseUtilityId="& luid &" and LoadProfile=1 and pp<>1))) as total"
'response.write strsql
'response.end
    rst1.open strsql, cnn1
    objChart.AddStaticText "*Aggregated view is made up of "&rst1("part")&" out of "&rst1("total")&" meters",0,296,RGB(100,100,100),"Arial",7,0,0
    rst1.close
end if

'--------------------------------------------
' setup the legend
'--------------------------------------------
objChart.Legend.Enabled = true  'enable the legend (it is disabled by default)
objChart.Legend.FontSize = 8
'setup the labels for each series, ese names will then appear in the legend
'Note: if using a legend with a pie chart the labels are retrieved from the charts
'data points (see the Chart Object's AddData member for more details) instead of
' from data series.
'Note: you can also specify custom legend entries by calling the Add method of
' the legend object.
objChart.Legend.Add "KWH", RGB(0,0,0)	
objChart.Legend.Add "KW", RGB(40,120,255)
'set the position of the legend
objChart.Legend.SetPosition 490,170,550,210 

'optional legend settings
objChart.Legend.BorderColor = RGB(110,0,0)
objChart.Legend.BackgroundColor = RGB(230,230,230)
objChart.Legend.FontColor = RGB(0,0,110)
objChart.Legend.FontSize = 8
objChart.Legend.Transparent = true 'set to false, so that the background color
                                    ' can be seen
'--------------------------------------------
' finished setting up the legend
'--------------------------------------------
else


objChart.AddStaticText "No "& graphtypename &"Profile Data Available for "& displaydate,300,3,RGB(0,0,100),"Arial",13,1,2

end if
'Return the chart as a JPEG
objChart.SendJpeg 550,210,50

'cleanup
set objChart = nothing
%>