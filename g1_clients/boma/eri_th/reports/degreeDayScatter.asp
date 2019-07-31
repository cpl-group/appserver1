<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim building, startdate, enddate, currentdate, prm, utilityid, columntitle, wddscatter, wddtrend
startdate = request("startdate")
enddate = request("enddate")
building = request("building")
utilityid = request("utilityid")
currentdate = startdate
if trim(session("wddscatter"))<>"" then wddscatter = session("wddscatter") else wddscatter = request("wddscatter")
if trim(session("wddtrend"))<>"" then wddtrend = session("wddtrend") else wddtrend = request("wddtrend")

dim rst1, rst2, cnn1, cmd
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")
cnn1.CursorLocation = adUseClient
'SETTING UP STORED PROC
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

dim objChart
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

rst2.open "SELECT measure FROM tblutility WHERE utilityid="&utilityid, cnn1
columntitle = trim(rst2("measure"))
rst2.close

''''''''''''''''
dim ypeak, xpeak, hasNegative, hasPositive, xinterval, yinterval, addfour, min, max, trendlineC(8,2), trendlineH(8,2), interC, interH, tempindex
for tempindex = 0 to 8
  trendlineC(tempindex,0) = 0
  trendlineC(tempindex,1) = 0
  trendlineC(tempindex,2) = 1
  trendlineH(tempindex,0) = 0
  trendlineH(tempindex,1) = 0
  trendlineH(tempindex,2) = 1
next

ypeak = 10
xpeak = 10
getscatterpoints currentdate,enddate,false
'getscatterpoints dateadd("m",-1,startdate),dateadd("m",-1,enddate),true

objChart.AddStaticText "Usage ("&columntitle&")", 1, 150, rgb(100,100,100), "Arial", 8, 1, 2, 90
objChart.AddStaticText "Usage ("&columntitle&")", 388, 150, rgb(100,100,100), "Arial", 8, 1, 2, 90
objChart.AddStaticText "Heating & Cooling Days for the period "&startdate&" - "&enddate, 1, 1, rgb(00,00,00), "Arial", 8, 1, 0
objChart.AddStaticText "Degree Days", 200, 250, rgb(100,100,100), "Arial", 8, 1, 2

objChart.SetSeriesColor 0, rgb(200,150,150)
objChart.SetSeriesColor 2, rgb(150,150,200)
objChart.SetSeriesColor 4, rgb(200,00,00)
objChart.SetSeriesColor 6, rgb(00,00,200)
objChart.SetSeriesColor 8, rgb(00,00,200)
objChart.SetSeriesColor 10, rgb(00,00,200)

objChart.Legend.Enabled = true
objChart.Legend.add "Cooling", rgb(00,00,200)
objChart.Legend.add "Heating", rgb(200,00,00)
objChart.Legend.add "Cooling Pevious yr", rgb(150,150,200)
objChart.Legend.add "Heating Pevious yr", rgb(200,150,150)
objChart.Legend.SetPosition 300,265,400,300

'set scales
ypeak = round(ypeak)
xpeak = round(xpeak)
yinterval = clng((10^(len(ypeak)-1)))
xinterval = clng((10^(len(xpeak)-1)))
if ypeak/yinterval<4 then yinterval = yinterval/2
if xpeak/xinterval<4 then xinterval = xinterval/2
ypeak = (yinterval - ypeak mod yinterval) + ypeak
xpeak = (xinterval - xpeak mod xinterval) + xpeak

objChart.ChartArea(2).Axis(0).Maximum = yPeak
objChart.ChartArea(2).Axis(2).Maximum = yPeak
objChart.ChartArea(2).Axis(1).Minimum = -1 * xPeak
objChart.ChartArea(2).Axis(3).Minimum = -1 * xPeak
objChart.ChartArea(2).Axis(1).Interval = xinterval
objChart.ChartArea(2).Axis(0).Interval = yinterval
objChart.ChartArea(3).Axis(2).Interval = yinterval
objChart.ChartArea(3).Axis(0).Maximum = yPeak
objChart.ChartArea(3).Axis(2).Maximum = yPeak
objChart.ChartArea(3).Axis(1).Maximum = xPeak
objChart.ChartArea(3).Axis(3).Maximum = xPeak
objChart.ChartArea(3).Axis(1).Interval = xinterval
objChart.ChartArea(3).Axis(0).Interval = yinterval
objChart.ChartArea(3).Axis(2).Interval = yinterval

objChart.ChartArea(6).Axis(1).Maximum = xPeak
objChart.ChartArea(6).Axis(0).Maximum = yPeak
objChart.ChartArea(6).Axis(1).Minimum = -1 * xPeak
objChart.ChartArea(6).Axis(0).Minimum = 0
objChart.ChartArea(6).Axis(0).SetNumberFormat 1, 0
objChart.ChartArea(6).Axis(1).SetNumberFormat 1, 0
objChart.ChartArea(6).Axis(1).Interval = xinterval
objChart.ChartArea(6).Axis(0).Interval = yinterval
objChart.ChartArea(6).Axis(0).FontSize = 8
objChart.ChartArea(6).Axis(1).FontSize = 8
objChart.ChartArea(6).Axis(1).Angle = 90
objChart.ChartArea(7).Axis(1).Maximum = xPeak
objChart.ChartArea(7).Axis(0).Maximum = yPeak
objChart.ChartArea(7).Axis(1).Minimum = -1 * xPeak
objChart.ChartArea(7).Axis(0).Minimum = 0
objChart.ChartArea(7).Axis(0).SetNumberFormat 1, 0
objChart.ChartArea(7).Axis(1).SetNumberFormat 1, 0
objChart.ChartArea(7).Axis(1).Interval = xinterval
objChart.ChartArea(7).Axis(0).Interval = yinterval
objChart.ChartArea(7).Axis(0).FontSize = 8
objChart.ChartArea(7).Axis(1).FontSize = 8
objChart.ChartArea(7).Axis(1).Angle = 90


if addfour=0 then
  objChart.ChartArea(0).Axis(0).Maximum = yPeak
  objChart.ChartArea(0).Axis(2).Maximum = yPeak
  objChart.ChartArea(0).Axis(1).Minimum = -1 * xPeak
  objChart.ChartArea(0).Axis(3).Minimum = -1 * xPeak
  objChart.ChartArea(0).Axis(1).Interval = xinterval
  objChart.ChartArea(0).Axis(0).Interval = yinterval
  objChart.ChartArea(1).Axis(2).Interval = yinterval
  objChart.ChartArea(1).Axis(0).Maximum = yPeak
  objChart.ChartArea(1).Axis(2).Maximum = yPeak
  objChart.ChartArea(1).Axis(1).Maximum = xPeak
  objChart.ChartArea(1).Axis(3).Maximum = xPeak
  objChart.ChartArea(1).Axis(1).Interval = xinterval
  objChart.ChartArea(1).Axis(0).Interval = yinterval
  objChart.ChartArea(1).Axis(2).Interval = yinterval
end if

'response.end
objChart.SendJPEG 400, 310



sub getscatterpoints(currentdate,enddate,prevyear)
  if prevyear then addfour = 0 else addfour = 4
  hasPositive = false
  hasNegative = false
  for tempindex = 0 to 8
    trendlineC(tempindex,0) = 0
    trendlineC(tempindex,1) = 0
    trendlineC(tempindex,2) = 0
    trendlineH(tempindex,0) = 0
    trendlineH(tempindex,1) = 0
    trendlineH(tempindex,2) = 0
  next

  if addfour=4 then
    rst2.open "select max(degreeDay) as [max], min(degreeDay) as [min] from (select avg(deg) as degreeDay, convert(datetime,left(date,12),101) as d from deg_day dd where dd.date>='"&currentdate&"' and dd.date<dateadd(day,1,'"&enddate&"') group by convert(datetime,left(date,12),101)) t", cnn1
    if not rst2.eof then
      if isnull(rst2("min")) then min = 0 else if cDbl(rst2("min"))<0 then min = cDbl(rst2("min")) else min = 0
      if isnull(rst2("max")) then max = 0 else if cDbl(rst2("max"))>0 then max = cDbl(rst2("max")) else max = 0
      interC = (0-min)/8
      interH = max/8
    end if
    rst2.close
  end if
  rst2.open "select avg(deg) as degreeDay, convert(datetime,left(date,12),101) as d from deg_day dd where dd.date>='"&currentdate&"' and dd.date<dateadd(day,1,'"&enddate&"') group by convert(datetime,left(date,12),101) order by d", cnn1
  cmd.Parameters("from")		= currentdate
  cmd.Parameters("to")		= enddate
  cmd.Parameters("code")		= "b"
  cmd.Parameters("string")		= building
  cmd.Parameters("utility")		= utilityid
  cmd.Parameters("interval")		= 2
  
  'response.write "sp_LMPDATA '"&currentdate&"','"&enddate&"','b','"&building&"',"&utilityid&",2,0"
  'response.end
  set rst1 = cmd.execute
  
  dim ddtemp, utemp
  if rst1.eof then 
  	objChart.AddData 10,0+addfour
  	objChart.AddData 0,1+addfour
  	objChart.AddData -10,2+addfour
  	objChart.AddData 0,3+addfour
  end if
  do until datediff("d",enddate,currentdate)>0
  	ddtemp = 0
  	utemp = 0
  	if not rst1.eof then
  		if year(currentdate)=year(rst1("date")) and month(currentdate)=month(rst1("date")) and day(currentdate)=day(rst1("date")) then
  			utemp = cdbl(rst1(columntitle))
  			rst1.movenext
  		end if
  	end if
  	if not rst2.eof then
  		if year(currentdate)=year(rst2("d")) and month(currentdate)=month(rst2("d")) and day(currentdate)=day(rst2("d")) then
  			ddtemp = cDbl(rst2("degreeDay"))
  			rst2.movenext
  		end if
  	end if
    if wddscatter then
    	objChart.AddData ddtemp,0+addfour
    	objChart.AddData utemp,1+addfour
    	objChart.AddData ddtemp,2+addfour
    	objChart.AddData utemp,3+addfour
    end if
    if ddtemp<0 or ddtemp = 0 then
      if interC<>0 then tempindex = int((ddtemp-min)/interC) else tempindex = 0
'      if ddtemp>min and (ddtemp<0 or ddtemp=0) and utemp<>0 then
        trendlineC(tempindex,0) = trendlineC(tempindex,0) + ddtemp
        trendlineC(tempindex,1) = trendlineC(tempindex,1) + utemp
        trendlineC(tempindex,2) = trendlineC(tempindex,2) + 1
'      end if
    end if
    if ddtemp>0 or ddtemp = 0 then
      if interH<>0 then tempindex = int(ddtemp/interH) else tempindex = 0
'      if ddtemp<max and (ddtemp<0 or ddtemp=0) and utemp<>0 then
        trendlineH(tempindex,0) = trendlineH(tempindex,0) + ddtemp
        trendlineH(tempindex,1) = trendlineH(tempindex,1) + utemp
        trendlineH(tempindex,2) = trendlineH(tempindex,2) + 1
'      end if
    end if
    
  	if ddtemp < 0 then
  		hasNegative = true
  	else
  		hasPositive = true
  	end if 
  	if abs(ddtemp)>xpeak then xpeak = abs(ddtemp)
  	if abs(utemp)>ypeak then ypeak = abs(utemp)
  	currentdate = dateadd("d",1,currentdate)
  loop
  if not wddscatter then
  	objChart.AddData -10000,0+addfour
  	objChart.AddData 10000,1+addfour
  	objChart.AddData 10000,2+addfour
  	objChart.AddData 10000,3+addfour
  end if
  if not(hasNegative) then
  	objChart.AddData 10,0+addfour
  	objChart.AddData 0,1+addfour
  end if
  if not(hasPositive) then
  	objChart.AddData -10,2+addfour
  	objChart.AddData 0,3+addfour
  end if
  rst1.close
  rst2.close
  if addfour=4 then 
    for tempindex = 0 to 8
      if trendlineC(tempindex,2)<>0 then objChart.AddData trendlineC(tempindex,0)/trendlineC(tempindex,2), 8 else objChart.AddData min+(interC*tempindex), 8
      if trendlineC(tempindex,2)<>0 then objChart.AddData trendlineC(tempindex,1)/trendlineC(tempindex,2), 9 else objChart.AddData 0, 9
      if trendlineH(tempindex,2)<>0 then objChart.AddData trendlineH(tempindex,0)/trendlineH(tempindex,2), 10 else objChart.AddData interH*tempindex, 10
      if trendlineH(tempindex,2)<>0 then objChart.AddData trendlineH(tempindex,1)/trendlineH(tempindex,2), 11 else objChart.AddData 0, 11
      'if trendlineH(tempindex,2)<>0 then response.write trendlineH(tempindex,0)/trendlineH(tempindex,2)&"|"&trendlineH(tempindex,1)/trendlineH(tempindex,2)&"<br>"
    next
  end if
  addfour = addfour/2
  if addfour=2 and wddtrend then
    objChart.ChartArea(6).AddChart 2,8,9
    objChart.ChartArea(6).Transparent = true
    objChart.ChartArea(6).GridVEnabled = false
    objChart.ChartArea(6).SetPosition 60,30,340,220
    objChart.ChartArea(7).AddChart 2,10,11
    objChart.ChartArea(7).Transparent = true
    objChart.ChartArea(7).GridVEnabled = false
    objChart.ChartArea(7).SetPosition 60,30,340,220
  end if
'  if wddscatter then
    objChart.ChartArea(0+addfour).AddChart 3,0+(addfour*2),1+(addfour*2)
    objChart.ChartArea(0+addfour).Transparent = true
    objChart.ChartArea(0+addfour).GridVEnabled = false
    objChart.ChartArea(0+addfour).SetPosition 60,30,200,220
    objChart.ChartArea(0+addfour).Axis(1).Maximum = 0
    objChart.ChartArea(0+addfour).Axis(3).Maximum = 0
    objChart.ChartArea(0+addfour).Axis(0).FontSize = 8
    objChart.ChartArea(0+addfour).Axis(1).FontSize = 8
    objChart.ChartArea(0+addfour).Axis(0).SetNumberFormat 1, 0
    objChart.ChartArea(0+addfour).Axis(1).angle = 90
    objChart.ChartArea(0+addfour).Axis(0).TruncatedLabels = true
    
    objChart.ChartArea(1+addfour).AddChart 3,2+(addfour*2),3+(addfour*2)
    objChart.ChartArea(1+addfour).Transparent = true
    objChart.ChartArea(1+addfour).GridVEnabled = false
    objChart.ChartArea(1+addfour).SetPosition 200,30,340,220
    objChart.ChartArea(1+addfour).Axis(1).Minimum = 0
    objChart.ChartArea(1+addfour).Axis(3).Minimum = 0
    objChart.ChartArea(1+addfour).Axis(0).Enabled = false
    objChart.ChartArea(1+addfour).Axis(2).Enabled = true
    objChart.ChartArea(1+addfour).Axis(1).FontSize = 8
    objChart.ChartArea(1+addfour).Axis(2).FontSize = 8
    objChart.ChartArea(1+addfour).Axis(2).SetNumberFormat 1, 0
    objChart.ChartArea(1+addfour).Axis(1).angle = 90
'  end if
end sub
%>
