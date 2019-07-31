<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--METADATA NAME="TeeChart Pro v5 ActiveX Control" TYPE="TypeLib" UUID="{B6C10482-FB89-11D4-93C9-006008A7EED4}"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
sub incFifteen(byref chour, byref cminute)
	if cminute=45 then
		cminute = 0
		chour = chour + 1
		if chour=24 then chour = 0
	else
		cminute = cminute + 15
	end if
end sub

dim building, startdate, enddate, currentdate, monthday, dayhour, utilityid, columntitle, ispeakday, utusage, uttemp, uthumid, rid
building = request("building")
startdate = request("startdate")
utilityid = request("utilityid")
enddate = request("enddate")
ispeakday = request("ispeakday")
currentdate = startdate
if trim(session("utusage"))<>"" then utusage = session("utusage") else utusage = request("utusage")
if trim(session("uttemp"))<>"" then uttemp = session("uttemp") else uttemp = request("uttemp")
if trim(session("uthumid"))<>"" then uthumid = session("uthumid") else uthumid = request("uthumid")

dim rst1, rst2, cnn1, cmd, prm
set cmd = server.createobject("ADODB.command")
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(building)
cnn1.CursorLocation = adUseClient

rst2.open "SELECT region FROM buildings WHERE bldgnum='"&building&"'", cnn1
rid = trim(rst2("region"))
rst2.close

rst2.open "SELECT measure FROM tblutility WHERE utilityid="&utilityid, getConnect(0,building,"billing")
columntitle = trim(rst2("measure"))
rst2.close

dim Chart
Set Chart = CreateObject("TeeChart.TChart")
Chart.AddSeries(scBar)
Chart.AddSeries(scLine)
Chart.AddSeries(scLine)

dim chour, cminute, ckwh, ctemp, label, templow, hum, temphigh, daylabelinterval, hourlabelinterval
chour = 0
cminute = 0
templow = 200
temphigh = 0
hum = 0
rst2.open "SELECT date, ou_t as temperature, isnull(ou_h,0) as humidity FROM deg_day WHERE region="&rid&" and date>='"&startdate&"' and date<dateadd(day,1,'"&enddate&"') GROUP BY date, ou_t, ou_h ORDER BY date", cnn1
cmd.ActiveConnection = cnn1
cmd.CommandText = "sp_LMPDATA"
cmd.CommandType = adCmdStoredProc
Set prm = cmd.CreateParameter("from", adVarChar, adParamInput, 12)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("to", adVarChar, adParamInput, 12)
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
Set prm = cmd.CreateParameter("peakdemand", adInteger, adParamOutPut)
cmd.Parameters.Append prm

cmd.Parameters("from")		= startdate
cmd.Parameters("to")		= enddate
cmd.Parameters("code")		= "b"
cmd.Parameters("string")		= building
cmd.Parameters("utility")		= utilityid
cmd.Parameters("interval")		= 0
set rst1 = cmd.execute
'response.write "sp_LMPDATA '"&startdate&"','"&enddate&"','b','"&building&"',"&utilityid&",2,0"
'response.end

daylabelinterval = round(datediff("d",startdate,enddate)/25)
if daylabelinterval<1 then daylabelinterval = 1
'response.write cmd.Parameters("title")
'response.end
hourlabelinterval = round(datediff("h",startdate,enddate)/25)
if hourlabelinterval < 1 then 
	hourlabelinterval = 1
elseif hourlabelinterval < 2 then 
	hourlabelinterval = 2
elseif hourlabelinterval < 4 then 
	hourlabelinterval = 4
elseif hourlabelinterval < 6 then 
	hourlabelinterval = 6
elseif hourlabelinterval < 8 then 
	hourlabelinterval = 8
elseif hourlabelinterval < 12 then 
	hourlabelinterval = 12
end if

monthday = 0
dayhour = 0
do until datediff("n",enddate,currentdate)>0
	if monthday mod daylabelinterval = 0 and minute(currentdate)=0 and hour(currentdate)=0 then
		label = month(currentdate)&"/"&day(currentdate)&"   "
	elseif dayhour mod hourlabelinterval = 0 and minute(currentdate)=0 and hourlabelinterval < 13 then
		label = hour(currentdate)
	else
		label = ""
	end if

	if not rst1.eof then
		if year(currentdate)=year(rst1("date")) and month(currentdate)=month(rst1("date")) and day(currentdate)=day(rst1("date")) and day(currentdate)=day(rst1("date")) and minute(currentdate)=minute(rst1("date")) then
			ckwh = cdbl(rst1(columntitle))
			rst1.movenext
		else
			ckwh = 0
		end if
	else
		ckwh = 0
	end if

	if not rst2.eof then
		if year(currentdate)=year(rst2("date")) and month(currentdate)=month(rst2("date")) and day(currentdate)=day(rst2("date")) and day(currentdate)=day(rst2("date")) and minute(currentdate)=minute(rst2("date")) and not(isnull(rst2("temperature"))) then
			hum = cdbl(rst2("humidity"))
			ctemp = cdbl(rst2("temperature"))
			if ctemp<templow then templow = ctemp
			if ctemp>temphigh then temphigh = ctemp
			rst2.movenext
		else
			'ctemp = -2000
			'hum = -2000
		end if
	else
		'ctemp = -2000
		'hum = -2000
	end if
  
  Chart.Series(0).Add ckwh, label, rgb(00,00,200)
  Chart.Series(1).Add ctemp, label, rgb(200,00,00)
  Chart.Series(2).Add cint(hum), label, rgb(00,200,00)
	currentdate = dateadd("n",15,currentdate)
	if minute(currentdate)=0 then dayhour = dayhour + 1
	if hour(currentdate)=0 and minute(currentdate)=0 then 
		monthday = monthday + 1
		dayhour = 0
	end if
loop

'objChart.ChartArea(0).GridVEnabled = false
if not(utusage) then
'  objChart.ChartArea(0).Axis(0).enabled = false
'  objChart.ChartArea(0).GridHEnabled = false
end if

if lcase(columntitle)="kwh" then columntitle = "kw"

if not utusage then Chart.Series(0).Active = False
if not uttemp then Chart.Series(1).Active = False
if not uthumid then Chart.Series(2).Active = False
if uttemp then 
  Chart.Axis.Right.Title.Caption = "Temperature"
'  Chart.Axis.Right.AutomaticMinimum = False
'  Chart.Axis.Right.AutomaticMaximum = False
'  Chart.Axis.Right.Minimum = -1
'  Chart.Axis.Right.Maximum = round(temphigh+2)
else 
  Chart.Axis.Right.Title.Caption = "Humidity"
  Chart.Axis.Right.AutomaticMinimum = False
  Chart.Axis.Right.AutomaticMaximum = False
  Chart.Axis.Right.Minimum = 0
  Chart.Axis.Right.Maximum = 100
end if

'SETTINGS
Chart.Series(0).Name = "Usage"
Chart.Series(0).Color = rgb(00,00,200)
Chart.Series(0).Marks.Visible = False
Chart.Series(0).asBar.BarPen.Visible = False

Chart.Series(1).Name = "Temp"
Chart.Series(1).Color = rgb(200,00,00)
Chart.Series(1).VerticalAxis = aRightAxis

Chart.Series(2).Name = "Humidity"
Chart.Series(2).Color = rgb(00,200,00)
Chart.Series(2).VerticalAxis = aRightAxis

Chart.Axis.Left.AutomaticMinimum = False
Chart.Axis.Left.Minimum = 0
Chart.Axis.Left.AxisPen.Width = 1
Chart.Axis.Left.GridPen.Style = psSolid
Chart.Axis.Left.GridPen.color = rgb(100,100,100)
Chart.Axis.Left.MinorTicks.Visible = False
Chart.Axis.Left.Title.Caption = "Demand ("&columntitle&")"
Chart.Axis.Left.Title.Font.Bold = True
Chart.Axis.Right.AxisPen.Width = 1
Chart.Axis.Right.GridPen.Visible = False
Chart.Axis.Right.GridPen.color = rgb(100,100,100)
Chart.Axis.Right.MinorTicks.Visible = False
Chart.Axis.Right.Title.Font.Bold = True
Chart.Axis.bottom.AxisPen.Width = 1
Chart.Axis.Bottom.GridPen.Visible = False
Chart.Axis.bottom.MinorTicks.Visible = False
Chart.Axis.Bottom.Title.Caption = "Hour"
Chart.Axis.Bottom.Title.Font.Bold = True
Chart.axis.Bottom.Increment = 10
Chart.Header.Text.Clear
Chart.Header.Alignment = taLeftJustify
Chart.Header.Font.Bold = True
Chart.Header.Font.color = rgb(0, 0, 0)
Chart.Header.Text.Add "Load Profile of daily "&columntitle&" usage for "&startdate&" "&ispeakday

Chart.Aspect.View3D = False
Chart.Legend.Transparent = True
Chart.Legend.LegendStyle = lsSeries
Chart.Panel.Color = vbWhite
Chart.Panel.BevelOuter = bvNone
Chart.Panel.BevelInner = bvNone
Chart.Panel.MarginBottom = 0
Chart.Panel.MarginLeft = 0
Chart.Panel.MarginRight = 0
Chart.Panel.MarginTop = 0
Chart.Width = 650
Chart.Height = 310
Set rst1=nothing
Set cnn1=nothing
Response.BinaryWrite(Chart.Export.asGif.SaveToStream)
Set Chart=nothing
%>
