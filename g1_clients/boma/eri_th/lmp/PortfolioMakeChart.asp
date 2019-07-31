<%Option Explicit
dim date, s, e, pid, summation, cnn, rst, cnn2, rst2, objChart,bldglist
date = Request.QueryString("d")
pid = Request.QueryString("pid")
s = 100
e = 2400

Dim Bldgname
Dim AggKW(23)
Dim AggKWH(23)
Dim title(23)
Set cnn = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")
Set cnn2 = Server.CreateObject("ADODB.Connection")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set bldglist = Server.CreateObject("ADODB.recordset")
Set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

cnn.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
cnn2.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
dim sql

sql = "SELECT bldgnum FROM buildings WHERE portfolioid = '" & pid & "' AND bldgnum IN (SELECT bldgnum FROM master.dbo.rm)"
bldglist.open sql, cnn

dim displaydate, chartname
displaydate = weekdayname(weekday(date)) &" "& day(date) &", "& Monthname(Month(date)) &" "& Year(date)
chartname=""
rst2.open "SELECT DISTINCT bldg_name from clientsites WHERE bldgid='"& pid &"'", cnn2
if not(rst2.EOF) then chartname = rst2("bldg_name")
rst2.close

if not bldglist.EOF then 
	Dim index
	Dim KWH
	kwh=0
	while not bldglist.EOF 
		sql = "SELECT date, (datepart(hour,date)*100)+datepart(minute,date) as time, SUM(kwh) as kwh " &_
		"FROM pulse_" & bldglist("bldgnum") & " "  &_
		"WHERE meterid IN " &_
		"(SELECT meterid FROM [dbo].[Meters] WHERE pp = 1 AND meterid IN" &_
		"(SELECT meterid FROM meters WHERE bldgnum IN " &_
		"(SELECT bldgnum FROM buildings WHERE portfolioid = '"&pid&"'))) "&_
		"AND left(convert(char(20),date,101),11) = left(convert(char(20),convert(datetime,'"&date&"',101),101),11) "&_
		"AND ((datepart(hour,date)*100)+datepart(minute,date) BETWEEN "& s &" AND "& e &") GROUP BY date, (datepart(hour,date)*100)+datepart(minute,date) ORDER BY (datepart(hour,date)*100)+datepart(minute,date), date DESC"
		rst.open sql, cnn
	
		index = 1
		
		while(not(rst.EOF))
			title(index) = rst("time")
			AggKW(index) = formatnumber(AggKW(index) + rst("kwh"), 0)
			index=index + 1
			rst.moveNext
		wend
		rst.close
		bldglist.movenext
		
	wend
	AggKWH(1) = AggKW(1)
	for index = 2 to 23 
	
		AggKWH(index) = clng(AggKWH(index-1)) + clng(AggKW(index))
	
	next
	
		for index = 1 to 23

				objChart.AddData AggKW(index), 0, title(index)
				objChart.AddData AggKWH(index), 1
				
		next 
	
		objChart.ChartArea(0).SetPosition 55, 20, 525, 130
		objChart.ChartArea(0).AddChart 6, 0, 0
		objChart.ChartArea(0).AddChart 1, 1, 1, , 2
		objChart.SetSeriesColor 0, RGB(40,120,255)	'first series, light blue
		objChart.SetSeriesColor 1, RGB(0,0,0)		'second series, black
		objChart.ChartArea(0).LineWidth = 2
		objChart.ChartArea(0).BackgroundColor = RGB(232,232,232)
		objChart.ChartArea(0).GridColor = rgb(200,200,200)
		objChart.ChartArea(0).LineWidth = 2
		objChart.ChartArea(0).Axis(2).Enabled=true
		objChart.ChartArea(0).Axis(0).SetNumberFormat 1,0
		objChart.ChartArea(0).Axis(2).SetNumberFormat 1,0
		objChart.ChartArea(0).Axis(1).Angle = 45
		objChart.Rectangle3DEffect()
		objChart.BackgroundColor= RGB(255,255,255)
		
		objChart.AddStaticText "Load Profile for "&displaydate,55,3,RGB(100,100,100),"Arial",8,1
		objChart.AddStaticText chartname,567,3,RGB(100,100,100),"Arial",8,1,2
		objChart.AddStaticText "Demand (KW)",0,105,RGB(100,100,100),"Arial",8,0,0,90
		objChart.AddStaticText "Hour",275,170,RGB(100,100,100),"Arial",8,1
		objChart.AddStaticText "Energy (KWH)",588,75,RGB(100,100,100),"Arial",8,0,2,270
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
		objChart.Legend.SetPosition 270,190,300,230 
		
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
	objChart.AddStaticText "Portfolio Data Not Available",165,3,RGB(0,0,100),"Arial",14
end if
	objChart.SendJPEG 600, 310
	set objChart = nothing 
%>