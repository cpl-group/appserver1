<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
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

dim building, startdate, enddate, currentdate, monthday, dayhour, utilityid, columntitle, ispeakday, utusage, uttemp, uthumid
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
cnn1.open application("cnnstr_genergy2")
cnn1.CursorLocation = adUseClient

rst2.open "SELECT measure FROM tblutility WHERE utilityid="&utilityid, cnn1
columntitle = trim(rst2("measure"))
rst2.close

dim objChart
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

dim chour, cminute, ckwh, ctemp, label, templow, hum, temphigh, daylabelinterval, hourlabelinterval
chour = 0
cminute = 0
templow = 200
temphigh = 0
hum = 0
rst2.open "SELECT date, ou_t as temperature, isnull(ou_h,-2000) as humidity FROM deg_day WHERE date>='"&startdate&"' and date<dateadd(day,1,'"&enddate&"') GROUP BY date, ou_t, ou_h ORDER BY date", cnn1
'response.write "SELECT date, isnull(ou_t,-2000) as temperature, isnull(ou_h,-2000) as humidity FROM deg_day WHERE date>='"&enddate&"' and date<dateadd(day,1,'"&startdate&"') GROUP BY date, ou_t, ou_h ORDER BY date"
'response.end
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

	if not rst1.eof and utusage then
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
		if year(currentdate)=year(rst2("date")) and month(currentdate)=month(rst2("date")) and day(currentdate)=day(rst2("date")) and day(currentdate)=day(rst2("date")) and minute(currentdate)=minute(rst2("date")) then
			hum = cdbl(rst2("humidity"))
			if not isnull(rst2("temperature")) then
        ctemp = cdbl(rst2("temperature"))
  			if ctemp<templow then templow = ctemp
  			if ctemp>temphigh then temphigh = ctemp
      else
  			ctemp = -2000
      end if
			rst2.movenext
		else
			ctemp = -2000
			hum = -2000
		end if
	else
		ctemp = -2000
		hum = -2000
	end if
'	response.write rst2("date")
	objChart.AddData ckwh, 0, label
	objChart.AddData ctemp, 1
	objChart.AddData cint(hum), 2
	currentdate = dateadd("n",15,currentdate)
	if minute(currentdate)=0 then dayhour = dayhour + 1
	if hour(currentdate)=0 and minute(currentdate)=0 then 
		monthday = monthday + 1
		dayhour = 0
	end if
loop

objChart.ChartArea(0).AddChart 6,0,0
objChart.ChartArea(0).Transparent = true
objChart.ChartArea(0).Axis(0).FontSize = 8
objChart.ChartArea(0).Axis(1).FontSize = 8
objChart.ChartArea(0).GridVEnabled = false
if not(utusage) then
  objChart.ChartArea(0).Axis(0).enabled = false
  objChart.ChartArea(0).GridHEnabled = false
end if
objChart.ChartArea(0).SetPosition 50,30,575,270
objChart.ChartArea(0).Axis(0).SetNumberFormat 1, 0
if templow <> 200 then
	if uttemp then
    objChart.ChartArea(1).AddChart 1,1,1
  	objChart.ChartArea(1).Transparent = true
  	objChart.ChartArea(1).GridVEnabled = false
  	objChart.ChartArea(1).GridHEnabled = false
  	objChart.ChartArea(1).Axis(0).enabled = false
  	objChart.ChartArea(1).Axis(1).enabled = false
  	objChart.ChartArea(1).Axis(2).FontSize = 8
  	objChart.ChartArea(1).Axis(2).enabled = true
  	objChart.ChartArea(1).Axis(2).minimum = templow-2
  	objChart.ChartArea(1).Axis(0).minimum = templow-2
  	objChart.ChartArea(1).Axis(2).maximum = temphigh+2
  	objChart.ChartArea(1).Axis(0).maximum = temphigh+2
  	objChart.ChartArea(1).SetPosition 50,30,575,270
  	objChart.ChartArea(1).Axis(2).SetNumberFormat 1, 0
  end if
  if uthumid then
  	objChart.ChartArea(2).AddChart 1,2,2
  	objChart.ChartArea(2).Transparent = true
  	objChart.ChartArea(2).GridVEnabled = false
  	objChart.ChartArea(2).GridHEnabled = false
  	objChart.ChartArea(2).Axis(0).enabled = false
  	objChart.ChartArea(2).Axis(1).enabled = false
  	objChart.ChartArea(2).Axis(2).FontSize = 8
  	objChart.ChartArea(2).Axis(2).enabled = true
  	objChart.ChartArea(2).Axis(2).minimum = 0
  	objChart.ChartArea(2).Axis(0).minimum = 0
  	objChart.ChartArea(2).Axis(2).maximum = 100
  	objChart.ChartArea(2).Axis(0).maximum = 100
  	objChart.ChartArea(2).SetPosition 50,30,575,270
  end if
end if

objChart.SetSeriesColor 0, rgb(00,00,200)
objChart.SetSeriesColor 1, rgb(200,00,00)
objChart.SetSeriesColor 2, rgb(00,200,00)

if lcase(columntitle)="kwh" then columntitle = "kw"
objChart.AddStaticText "Load Profile of daily "&columntitle&" usage for "&startdate&" "&ispeakday, 1, 1, rgb(00,00,00), "Arial", 8, 1, 0
if utusage then objChart.AddStaticText "Demand ("&columntitle&")",0,140,RGB(100,100,100),"Arial",8,1,2,90
if uttemp then objChart.AddStaticText "Temperature",630,140,RGB(100,100,100),"Arial",8,1,2,90
if uthumid then objChart.AddStaticText "Humidity",630,140,RGB(100,100,100),"Arial",8,1,2,90
objChart.AddStaticText "Hour",325,290,RGB(100,100,100),"Arial",8,1,2

objChart.Legend.Enabled = true
if utusage then objChart.Legend.add "Usage", rgb(00,00,200)
if uttemp then objChart.Legend.add "Temperature", rgb(200,00,00)
if uthumid then objChart.Legend.add "Humidity", rgb(00,200,00)
objChart.Legend.FontSize = 6

dim legh 
legh = 275
if not(utusage) or (not(uttemp) and not(uthumid)) then legh = 287 
objChart.Legend.SetPosition 585,legh,650,300
objChart.SendJPEG 650, 310
%>
