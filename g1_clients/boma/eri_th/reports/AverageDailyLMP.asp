<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim groupname, startdate, enddate, columntitle, wday, label, ckwh, holiday, where, title
startdate = request("startdate")
enddate = request("enddate")
wday = request("day")
holiday = request("holiday")
groupname = request("groupname")
dim rst1, cnn1, cmd, prm, objChart
set cmd = server.createobject("ADODB.command")
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
cnn1.open application("cnnstr_genergy2")
cnn1.CursorLocation = adUseClient

dim groupCol, hours, peakdemand, pdDate, loadFactor, avgTemp, peak, decimals
rst1.open "SELECT top 1 * FROM ["&groupname&"]", cnn1
groupCol = rst1.fields.Item(0).Name
rst1.close

hours = 0
peak = 0
decimals = 0
if trim(holiday)<>"" or trim(wday)<>"" then
  if trim(holiday)<>"" then
    where = "convert(datetime,left(lmp.date,11)) in (SELECT date from holidaysch WHERE holiday='"& join(split(holiday,","),"' or holiday='") &"')"
    title = "Selected Holiday"
  else
    where = "(datepart(weekday,lmp.date)="& join(split(wday,",")," or datepart(weekday,lmp.date)=") &") and left(lmp.date,11) not in (SELECT DISTINCT left(date,11) FROM holidaysch)"
    if instr(wday,"1")<>0 or instr(wday,"7")<>0 then title = "Selected Weekend Days" else title = "Selected Weekdays"
'response.write where&"<br>"
  end if
  'get peakdemand
  rst1.open "SELECT top 1 datepart(hour,date) as hours, datepart(minute,date) as minutes, avg(isnull(lmp.["&groupCol&"],0)) as lmp1, avg(isnull(lmp2.innerusage,0)) as lmp2, (avg(isnull(lmp.["&groupCol&"],0))+avg(isnull(lmp2.innerusage,0)))/2 as pd FROM ["&groupname&"] lmp LEFT JOIN (SELECT dateadd(n,-15,il.date) as innerdate, il.["&groupCol&"] as innerusage FROM ["&groupname&"] il) lmp2 ON lmp2.innerdate=lmp.date WHERE ("&where&") and lmp.date>'"&startdate&"' and lmp.date<'"&enddate&"' GROUP BY datepart(hour,date), datepart(minute,date) ORDER BY pd desc", cnn1
  if not rst1.eof then
    if cDbl(rst1("pd"))<>0 then peakdemand = cDbl(rst1("pd")) else peakdemand = 1
    pdDate = rst1("hours")&":"&rst1("minutes")
  else
    peakdemand = 1
    pdDate = 0
  end if
  rst1.close
  'get load factor and average temperature
  rst1.open "SELECT case when datediff(hour,'"&startdate&"','"&enddate&"')=0 then 0 else (isnull(sum(["&groupCol&"])/4,0)/"&peakdemand&")/datediff(hour,'"&startdate&"','"&enddate&"') end as loadfactor, isnull(avg(dd.ou_t),0) as avgTemp FROM ["&groupname&"] lmp FULL OUTER JOIN deg_day dd ON lmp.date=dd.date WHERE ("&where&") and lmp.date>'"&startdate&"' and lmp.date<'"&enddate&"'", cnn1
  if not rst1.eof then
    loadFactor = cDbl(rst1("loadfactor"))
    avgTemp = cDbl(rst1("avgTemp"))
  end if
  rst1.close
  'get pulse data averages
  rst1.open "SELECT avg(["&groupCol&"]) as usage, datepart(hour,date) as h FROM ["&groupname&"] lmp WHERE ("&where&") and date>='"&startdate&"' and date<='"&enddate&"' GROUP BY datepart(hour,date) ORDER BY datepart(hour,date)", cnn1
  do while hours < 24
  	if not rst1.eof then
  		if hours=cint(rst1("h")) then
  			ckwh = cdbl(rst1("usage"))
  			rst1.movenext
  		else
  			ckwh = 0
  		end if
  	else
  		ckwh = 0
  	end if
  	objChart.AddData ckwh, 0, hours
    if ckwh > peak then peak = ckwh
    hours = hours + 1
  loop
if peak = 0 then peak = 1
if peak >5 then decimals = 0 else decimals = 1

  objChart.ChartArea(0).AddChart 1,0,0
  objChart.ChartArea(0).Transparent = true
  objChart.ChartArea(0).GridVEnabled = false
  objChart.ChartArea(0).GridHEnabled = true
  objChart.ChartArea(0).Axis(0).enabled = true
  objChart.ChartArea(0).Axis(1).enabled = true
  objChart.ChartArea(0).Axis(2).FontSize = 8
  objChart.ChartArea(0).Axis(2).enabled = true
  objChart.ChartArea(0).Axis(2).minimum = 0
  objChart.ChartArea(0).Axis(0).minimum = 0
  objChart.ChartArea(0).Axis(0).maximum = cint(peak)
  objChart.ChartArea(0).SetPosition 100,13,524,165
  objChart.ChartArea(0).Axis(0).SetNumberFormat 1, decimals
  objChart.AddStaticText "Peak Demand: "&formatnumber(peakdemand) &" "&groupCol,103,12,RGB(100,100,100),"Arial",8,1
  objChart.AddStaticText "Load Factor: "&formatpercent(loadFactor,2),103,22,RGB(100,100,100),"Arial",8,1
  objChart.AddStaticText "Average Temp: "&formatnumber(avgTemp),103,32,RGB(100,100,100),"Arial",8,1
  objChart.AddStaticText "Average Usage For " & title & " " & startdate &"-"&enddate,524,0,RGB(100,100,100),"Arial",8,1,1
  objChart.AddStaticText "Usage (" & groupCol & ")",20,92,RGB(100,100,100),"Arial",8,1,2,90
  objChart.AddStaticText "hour",320,180,RGB(100,100,100),"Arial",8,1,2
end if
'objChart.SetSeriesColor 0, rgb(00,00,200)
'objChart.SetSeriesColor 1, rgb(200,00,00)
'objChart.SetSeriesColor 2, rgb(00,200,00)


objChart.SendJPEG 525, 195
%>