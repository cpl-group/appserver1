<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--METADATA NAME="TeeChart Pro v5 ActiveX Control" TYPE="TypeLib" UUID="{B6C10482-FB89-11D4-93C9-006008A7EED4}"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim building, startdate, enddate, currentdate, prm, utilityid, columntitle, wddscatter, wddtrend, rid
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
cnn1.open getLocalConnect(building)
cnn1.CursorLocation = adUseClient
'SETTING UP STORED PROC
cmd.ActiveConnection = cnn1
cmd.CommandText = "sp_LMPDATA"
cmd.CommandType = adCmdStoredProc
Set prm = cmd.CreateParameter("from", adVarChar, adParamInput, 12)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("to", adVarChar, adParamInput, 21)
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

dim Chart
Set Chart = CreateObject("TeeChart.TChart")
Chart.AddSeries(scPoint)
Chart.AddSeries(scPoint)
Chart.AddSeries(scLine)

rst2.open "SELECT region FROM buildings WHERE bldgnum='"&building&"'", cnn1
rid = trim(rst2("region"))
rst2.close

rst2.open "SELECT measure FROM tblutility WHERE utilityid="&utilityid, getConnect(0,building,"billing")
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

'set scales
xpeak = round(xpeak)
xinterval = clng((10^(len(xpeak)-1)))
if xpeak/xinterval<4 then xinterval = xinterval/2
xpeak = (xinterval - xpeak mod xinterval) + xpeak

'CURVE FUNCTION
Chart.Series(2).Color = rgb(00,00,200)
Chart.Series(2).SetFunction tfCurveFit
Chart.Series(2).FunctionType.PeriodStyle = psNumPoints
Chart.Series(2).FunctionType.PeriodAlign = paCenter
Chart.Series(2).FunctionType.Period = 50
on error resume next
Chart.Series(2).DataSource = "Series0"
on error goto 0

'SETTING
if not wddscatter then 
  Chart.Series(0).Active = False
  Chart.Series(1).Active = False
end if
if not wddtrend then Chart.Series(2).Active = False

Chart.Series(0).Name = "HeatingDays"
Chart.Series(0).Color = rgb(200,00,00)
Chart.Series(0).asPoint.Pointer.Style = psCircle

Chart.Series(1).Name = "CoolingDays"
Chart.Series(1).Color = rgb(00,00,200)
Chart.Series(1).VerticalAxisCustom = 0
Chart.Series(1).asPoint.Pointer.Style = psCircle

Chart.Series(2).Name = "Usage"
Chart.Series(2).Color = rgb(00,00,200)
Chart.Series(2).Marks.Visible = False
Chart.Series(2).ShowInLegend = False
Chart.Series(2).asLine.LinePen.Width = 2

Chart.Axis.Left.AutomaticMinimum = False
Chart.Axis.Left.Minimum = 0
Chart.Axis.Left.AxisPen.Width = 1
Chart.Axis.Left.GridPen.Style = psSolid
Chart.Axis.Left.GridPen.color = rgb(100,100,100)
Chart.Axis.Left.MinorTicks.Visible = False
Chart.Axis.Left.Title.Caption = "Usage ("&columntitle&")"
Chart.Axis.Left.Title.Font.Bold = True
Chart.Axis.AddCustom False
Chart.Axis.Custom(0).AutomaticMinimum = False
Chart.Axis.Custom(0).Minimum = 0
Chart.Axis.Custom(0).AxisPen.Width = 1
Chart.Axis.Custom(0).GridPen.Style = psSolid
Chart.Axis.Custom(0).GridPen.color = rgb(100,100,100)
Chart.Axis.Custom(0).MinorTicks.Visible = False
Chart.Axis.Custom(0).Title.Caption = "Usage ("&columntitle&")"
Chart.Axis.Custom(0).Title.Font.Bold = True
Chart.Axis.Custom(0).PositionPercent = 50
Chart.Axis.Bottom.AutomaticMinimum = False
Chart.Axis.Bottom.AutomaticMaximum = False
Chart.Axis.Bottom.Minimum = xPeak*-1
Chart.Axis.Bottom.Maximum = xPeak
Chart.Axis.Bottom.AxisPen.Width = 1
Chart.Axis.Bottom.GridPen.Visible = False
Chart.Axis.Bottom.MinorTicks.Visible = False
Chart.Axis.Bottom.Title.Caption = "Degree Days"
Chart.Axis.Bottom.Title.Font.Bold = True
Chart.axis.Bottom.Increment = 10

Chart.Tools.Add tcCursor
Chart.Tools.Items(0).asTeeCursor.Style = cssVertical

Chart.Header.Text.Clear
Chart.Header.Alignment = taLeftJustify
Chart.Header.Font.Bold = True
Chart.Header.Font.color = rgb(0, 0, 0)
Chart.Header.Text.Add "Heating & Cooling Days for the period "&startdate&" - "&enddate

Chart.Aspect.View3D = False
Chart.Legend.Transparent = True
Chart.Legend.LegendStyle = lsSeries
Chart.Legend.Alignment = laBottom
Chart.Panel.Color = vbWhite
Chart.Panel.BevelOuter = bvNone
Chart.Panel.BevelInner = bvNone
Chart.Panel.MarginBottom = 0
Chart.Panel.MarginLeft = 0
Chart.Panel.MarginRight = 2
Chart.Panel.MarginTop = 0
Chart.Width = 400
Chart.Height = 310
Set rst1=nothing
Set cnn1=nothing
Response.BinaryWrite(Chart.Export.asGif.SaveToStream)
Set Chart=nothing



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
    rst2.open "select max(isnull(degreeDay,0)) as [max], min(isnull(degreeDay,0)) as [min] from (select isnull(avg(deg),0) as degreeDay, convert(datetime,left(date,12),101) as d from dbo.deg_day dd where region="&rid&" and dd.date>='"&currentdate&"' and dd.date<dateadd(day,1,'"&enddate&"') group by convert(datetime,left(date,12),101)) t", cnn1
    if not rst2.eof then
      if isnull(rst2("min")) then min = 0 else if cDbl(rst2("min"))<0 then min = cDbl(rst2("min")) else min = 0
      if isnull(rst2("max")) then max = 0 else if cDbl(rst2("max"))>0 then max = cDbl(rst2("max")) else max = 0
      interC = (0-min)/8
      interH = max/8
    end if
    rst2.close
  end if
  rst2.open "select isnull(avg(deg),0) as degreeDay, convert(datetime,left(date,12),101) as d from dbo.deg_day dd where region="&rid&" and dd.date>='"&currentdate&"' and dd.date<dateadd(day,1,'"&enddate&"') group by convert(datetime,left(date,12),101) order by d", cnn1
  cmd.Parameters("from")		= currentdate
  cmd.Parameters("to")		= enddate & " 23:59:00"
  cmd.Parameters("code")		= "b"
  cmd.Parameters("string")		= building
  cmd.Parameters("utility")		= utilityid
  cmd.Parameters("interval")		= 2
  
  'response.write "sp_LMPDATA '"&cmd.Parameters("from")&"','"&cmd.Parameters("to")&"','b','"&building&"',"&utilityid&",2,0"
  'response.end
  set rst1 = cmd.execute
  
  dim ddtemp, utemp
  do until datediff("d",enddate,currentdate)>0
  	ddtemp = 0
  	utemp = 0
  	if not rst1.eof then
  		if not(isnull(rst1(columntitle))) and year(currentdate)=year(rst1("date")) and month(currentdate)=month(rst1("date")) and day(currentdate)=day(rst1("date")) then
  			utemp = cdbl(rst1(columntitle))
  			rst1.movenext
		else
  			utemp = 0
  			rst1.movenext
  		end if
  	end if
  	if not rst2.eof then
  		if not(isnull(rst2("degreeDay"))) and year(currentdate)=year(rst2("d")) and month(currentdate)=month(rst2("d")) and day(currentdate)=day(rst2("d")) then
  			ddtemp = cDbl(rst2("degreeDay"))
  			rst2.movenext
		else
  			ddtemp = 0
  			rst2.movenext
  		end if
  	end if
    if ddtemp >= 0 then Chart.Series(0).AddXY ddtemp, utemp, "", rgb(00,00,200) else Chart.Series(0).AddXY ddtemp, utemp, "", rgb(200,00,00)
  	if abs(ddtemp)>xpeak then xpeak = abs(ddtemp)
'	response.write currentdate&","&utemp&","&ddtemp&"<br>"
  	currentdate = dateadd("d",1,currentdate)
  loop
'  response.end
  rst1.close
  rst2.close
end sub
%>
