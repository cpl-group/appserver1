<%@Language=VBScript%>
<%Option Explicit

Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!-- #INCLUDE file ="ChartConst.inc" -->
<!--#include file="adovbs.inc"-->

<%
dim objChart			'Dundas Chart 2D object
Dim Bldgname
dim ctr					'loop counter
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
'Set objChart = Server.CreateObject("Dundas.ChartServer")
Dim Title(12) 'Pair title
Title(1) = "Jan"
Title(2) = "Feb"
Title(3) = "Mar"
Title(4) = "Apr"
Title(5) = "May"
Title(6) = "Jun"
Title(7) = "Jul"
Title(8) = "Aug"
Title(9) = "Sep"
Title(10) = "Oct"
Title(11) = "Nov"
Title(12) = "Dec"

Dim cnn1
Dim rst1
Dim strsql
Dim bldg
Dim meterid, lmpstart, lmpend,lmpdate, luid, lmp

bldg=Request.Querystring("b")
meterid=Request.QueryString("m")
lmpstart=Request.QueryString("s")
lmpend=Request.QueryString("e")
lmpdate=Request.QueryString("d")
luid = Request.QueryString("luid")
lmp = Request.QueryString("lmp")
Dim modulus
modulus=Request.QueryString("i")
Dim Graphtype 

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
if luid <> "" then 

	strsql = "select time, sum(kwh) as kwh from pulse_" & bldg & " where meterid in (select meterid from meters where LeaseUtilityId='" & luid & "') and left(date,11) = convert(datetime,'" & lmpdate & "',101) AND (pulse_" & bldg & ".time BETWEEN '" & lmpstart & "' AND '" & lmpend & "') group by time order by time "
	graphtype = area_chart

else

	strsql = "SELECT pulse_" & bldg & ".date as date, pulse_" & bldg & ".time, (pulse_" & bldg & ".delta) AS Pulse, (meters.multiplier) as multiplier, pulse_" & bldg & ".kwh FROM pulse_" & bldg & " INNER JOIN Meters ON pulse_" & bldg & ".meterid = Meters.MeterId WHERE (left(pulse_" & bldg & ".date,11)) = convert(datetime,'" & lmpdate & "',101) AND (pulse_" & bldg & ".meterid = '" & meterid & "') AND (pulse_" & bldg & ".time BETWEEN '" & lmpstart & "' AND '" & lmpend & "') AND (pulse_" & bldg & ".time % 1 = 0) ORDER BY pulse_" & bldg & ".time"
	graphtype = column_chart
end if
rst1.Open strsql, cnn1, adOpenStatic

if not rst1.eof then 
Dim numRecords
numRecords = rst1.RecordCount
dim ArrLMP()		'Array of first set of data - LMP 
dim ArrYearlyKWH()
dim ArrYearlyKW()
dim ArrKWH()
Dim TimeArr()


if (numRecords mod 4 <> 0) and (numrecords < 48) then 
	modulus = 1 
	ReDim ArrLMP(numRecords)
else
	if modulus = 100 then 
		ReDim ArrLMP(numRecords / 4)
	else
		ReDim ArrLMP(numRecords)
	end if 
end if 

ReDim ArrKWH(numRecords)
ReDim TimeArr(numRecords)

Dim KW
Dim KWH
kwh = 0
kw=0
Dim index
Dim aindex
aindex = 1
	
	for index=1 to numRecords
		
		if modulus = 100 then 
			if rst1("time") mod 100 = 0 then 
					if  clng(rst1("kwh")) > kw then 
			   			kw = clng(rst1("kwh"))
					end if 
					ArrLMP(aindex) = clng(kw) * 4
					TimeArr(aindex) = rst1("time")
					kw=0
					aindex= aindex + 1		
				else
					if  clng(rst1("kwh")) > kw then 
			   			kw = clng(rst1("kwh"))
					end if 
			end if
		else
			ArrLMP(aindex) = clng(rst1("kwh")) * 4
			TimeArr(aindex) = rst1("time")
			aindex= aindex + 1		
		end if
		kwh = kwh + clng(rst1("kwh"))
		ArrKWH(index) = kwh
		rst1.movenext
		
	next
rst1.movefirst



for ctr=1 to aindex-1
	if modulus = 100 then 
		objChart.AddData ArrLMP(ctr), 0, timearr(ctr) 	'Add data to Data Series 0 
	else
		if timearr(ctr) mod 100 = 0 then 
		objChart.AddData ArrLMP(ctr), 0,timearr(ctr) 	'Add data to Data Series 0 
		else
		objChart.AddData ArrLMP(ctr), 0 	'Add data to Data Series 0 
		end if
	end if
	objChart.AddData ArrKWH(ctr), 1		'and assign a label to this data
next

rst1.close

if lmp ="" then 
	if luid = "" then
		strsql = "select * from (select  top 12 billyear, billperiod, demand_p as Demand, KWHUsed as Consumption  from tblmetersbyperiod where meterid = " & meterid & " order by id desc) as cat order by cat.billyear, cat.billperiod "
	else
		strsql = "select billyear, billperiod, sum(demand_p) as Demand, sum(KWHUsed) as Consumption from tblmetersbyperiod where leaseutilityid ='" & luid & "' and billyear ='"& year(lmpdate) & "' group by billyear, billperiod"
	end if
else
	strsql = "select * from (Select top 12 billyear, billperiod, totalkwh as Consumption, totalkw as Demand from utilitybill, billyrperiod where billyrperiod.ypid=utilitybill.ypid and billyrperiod.bldgnum='" & bldg & "'  order by  id desc) as cat order by cat.billyear, cat.billperiod"

end if

'response.write strsql
'response.end

rst1.Open strsql, cnn1, adOpenStatic
If not rst1.EOF then 

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
objChart.ChartArea(1).AddChart column_CHART, 3, 3, ,0  'second chart,fourth series\
objChart.ChartArea(1).AddChart line_CHART, 2, 2	, , 2'second chart,fourth series
objChart.AddStaticText "Monthly Billed Consumption / Demand",55,175,RGB(100,100,100),"Arial",8,1


end if
rst1.close


'setup the chart types
objChart.ChartArea(0).AddChart graphtype, 0, 0	'first chart,first series
objChart.ChartArea(0).AddChart LINE_CHART, 1, 1, , 2 'first chart,second series
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
objChart.ChartArea(0).SetPosition 55, 20, 525, 130
objChart.ChartArea(1).SetPosition 55, 190, 525,260

'add text to the chart
objChart.AddStaticText "Load Profile for " & lmpdate,55,3,RGB(100,100,100),"Arial",8,1
objChart.AddStaticText "Demand (KW)",0,105,RGB(100,100,100),"Arial",8,0,0,90
objChart.AddStaticText "Energy (KWH)",578,35,RGB(100,100,100),"Arial",8,0,0,270
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
objChart.Legend.SetPosition 550,270,600,310 

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
objChart.AddStaticText "No Profile Data Available",165,3,RGB(0,0,100),"Arial",14

end if
'Return the chart as a JPEG
objChart.SendJpeg 600,310

'cleanup
set objChart = nothing
%>