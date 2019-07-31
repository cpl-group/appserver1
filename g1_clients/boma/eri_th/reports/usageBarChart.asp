<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim startdate, enddate, building, currentdate, utilityid, dutcdate, uddusage, udddegree, pdf
if trim(request("pdf"))="yes" then pdf = true else pdf = false
startdate = request("startdate")
enddate = request("enddate")
utilityid = request("utilityid")
building = request("building")
dutcdate = request("dutcdate")
currentdate = startdate
if trim(request("uddusage"))="True" then uddusage = true else uddusage = false
if trim(request("udddegree"))="True" then udddegree = true else udddegree = false

dim rst1, rst2, cnn1, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
set cmd = server.createobject("ADODB.command")
cnn1.open application("cnnstr_genergy2")
cnn1.CursorLocation = adUseClient

dim objChart
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

dim monthday, hasNegative, hasPositive, columntitle, labelinterval, label, color
monthday = 1

rst2.open "SELECT measure FROM tblutility WHERE utilityid="&utilityid, cnn1
columntitle = trim(rst2("measure"))
rst2.close

rst2.open "SELECT convert(datetime,left(date,12),101) as d, avg(deg) as deg FROM deg_day dd WHERE dd.date>='"&startdate&"' and dd.date<dateadd(day,1,'"&enddate&"') GROUP BY convert(datetime,left(date,12),101) ORDER BY d", cnn1
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
cmd.Parameters("interval")		= 2
'response.write "sp_LMPDATA '"&startdate&"','"&enddate&"','b','"&building&"',"&utilityid&",2,0,0,0"
'response.end

set rst1 = cmd.execute

labelinterval = round(datediff("d",startdate,enddate)/25)
if labelinterval<1 then labelinterval = 1
do until datediff("d",enddate,currentdate)>0
	if monthday mod labelinterval = 0 then label = month(currentdate)&"/"&day(currentdate) else label = ""
  if datediff("d",dutcdate,currentdate)=0 then color = rgb(100,100,255) else color = rgb(151,200,255)
  if not rst1.eof and uddusage then
		if year(currentdate)=year(rst1("date")) and month(currentdate)=month(rst1("date")) and day(currentdate)=day(rst1("date")) then
			objChart.AddData rst1(columntitle), 0, label, color
			rst1.movenext
		else
			objChart.AddData 0, 0, label, color
		end if
	else
		objChart.AddData 0, 0, label, color
	end if
	if not rst2.eof then
		if year(currentdate)=year(rst2("d")) and month(currentdate)=month(rst2("d")) and day(currentdate)=day(rst2("d")) then
			if not(clng(rst2("deg"))<0) then
				objChart.AddData monthday, 1
				objChart.AddData clng(rst2("deg")), 2
				hasNegative = true
			else
				objChart.AddData monthday, 3
				objChart.AddData abs(clng(rst2("deg"))), 4
				hasPositive = true
			end if
			rst2.movenext
		end if
	else
		objChart.AddData monthday, 1
		objChart.AddData 0, 2
		objChart.AddData monthday, 3
		objChart.AddData 0, 4
	end if 
	currentdate = dateadd("d",1,currentdate)
	monthday = monthday + 1
loop

if not(hasNegative) then
	objChart.AddData -10000,1
	objChart.AddData 0,2
end if
if not(hasPositive) then
	objChart.AddData -10000,3
	objChart.AddData 0,4
end if
objChart.ChartArea(0).AddChart 7,0,0
objChart.AddStaticText startdate&" - "&enddate, 575, 15, RGB(100,100,100),"Arial",8,1,1
objChart.ChartArea(0).Transparent = true
objChart.ChartArea(0).GridVEnabled = false
objChart.ChartArea(0).Axis(0).FontSize = 8
objChart.ChartArea(0).Axis(1).FontSize = 8
if not(uddusage) then objChart.ChartArea(0).Axis(0).enabled = false
objChart.ChartArea(0).SetPosition 50,30,575,270
objChart.ChartArea(0).Axis(0).SetNumberFormat 1, 0
if udddegree then
  objChart.ChartArea(1).SetPosition 50,30,575,270
  objChart.ChartArea(1).AddChart 3,1,2
  objChart.ChartArea(1).Transparent = true
  objChart.ChartArea(1).GridVEnabled = false
  objChart.ChartArea(1).Axis(0).enabled = false
  objChart.ChartArea(1).Axis(1).enabled = false
  objChart.ChartArea(1).Axis(1).maximum = monthday
  objChart.ChartArea(1).AddChart 3,3,4
  objChart.ChartArea(1).Transparent = true
  objChart.ChartArea(1).GridVEnabled = false
  objChart.ChartArea(1).Axis(0).enabled = false
  objChart.ChartArea(1).Axis(2).enabled = true
  objChart.ChartArea(1).Axis(1).enabled = false
  objChart.ChartArea(1).Axis(1).maximum = monthday
  objChart.ChartArea(1).Axis(1).Minimum = 0
  objChart.ChartArea(1).GridHEnabled = false
  objChart.ChartArea(1).GridHEnabled = false
  objChart.ChartArea(1).Axis(2).SetNumberFormat 1, 0
  objChart.ChartArea(1).Axis(2).FontSize = 8
  objChart.AddStaticText "Degree Days",630,140,RGB(100,100,100),"Arial",8,1,2,90
end if
objChart.SetColorFromPoint 0
objChart.SetSeriesColor 0, rgb(200,200,200)
objChart.SetSeriesColor 1, rgb(00,200,00)
objChart.SetSeriesColor 3, rgb(200,00,00)
objChart.AddStaticText "Usage Versus Degree Days", 1, 1, rgb(00,00,00), "Arial", 8, 1, 0
objChart.AddStaticText "Hour",325,305,RGB(100,100,100),"Arial",8,1,2
if lcase(columntitle)="tons" then columntitle = "Ton/hrs"
if uddusage then objChart.AddStaticText "Usage ("&columntitle&")",0,140,RGB(100,100,100),"Arial",8,1,2,90

objChart.Legend.Enabled = true
if uddusage then objChart.Legend.add "Usage", rgb(151,200,255)
if udddegree then 
  objChart.Legend.add "Cooling", rgb(00,200,00)
  objChart.Legend.add "Heating", rgb(200,00,00)
end if
objChart.Legend.FontSize = 6
'else
'	objChart.AddStaticText "The Period "&startdate&" - "&enddate&" contains no data.", 300, 250, RGB(100,100,100),"Arial",15,1,2
'end if
dim legh 
legh = 250
if uddusage and not(udddegree) then legh = 287 else if not(uddusage) and udddegree then legh = 275
objChart.Legend.SetPosition 585,legh,650,300
objChart.SendJPEG 650, 320
%>
