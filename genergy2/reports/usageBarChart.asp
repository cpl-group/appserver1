<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--METADATA NAME="TeeChart Pro v5 ActiveX Control" TYPE="TypeLib" UUID="{B6C10482-FB89-11D4-93C9-006008A7EED4}"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim startdate, enddate, building, currentdate, utilityid, dutcdate, uddusage, udddegree, pdf, sqlstr
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
cnn1.open getLocalConnect(building)
cnn1.CursorLocation = adUseClient

dim Chart
Set Chart = CreateObject("TeeChart.TChart")
Chart.AddSeries(scBar)
Chart.AddSeries(scLine)
Chart.AddSeries(scLine)

dim monthday, hasNegative, hasPositive, columntitle, labelinterval, label, color, rid
monthday = 1

sqlstr = "SELECT region FROM buildings WHERE bldgnum='"&building&"'"
rst2.open sqlstr, cnn1
rid = trim(rst2("region"))
rst2.close

sqlstr="SELECT measure FROM tblutility WHERE utilityid="&utilityid
rst2.open sqlstr, getConnect(0,building,"billing")
columntitle = trim(rst2("measure"))
rst2.close

sqlstr= "SELECT convert(datetime,left(date,12),101) as d, avg(deg) as deg FROM dbo.deg_day dd WHERE region="&rid&" and dd.date>='"&startdate&"' and dd.date<dateadd(day,1,'"&enddate&"') GROUP BY convert(datetime,left(date,12),101) ORDER BY d"
rst2.open sqlstr, cnn1

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
cmd.Parameters("from")		= startdate
cmd.Parameters("to")		= enddate&" 23:59:00"
cmd.Parameters("code")		= "b"
cmd.Parameters("string")		= building
cmd.Parameters("utility")		= utilityid
cmd.Parameters("interval")		= 2

set rst1 = cmd.execute

dim usage, ypeak
ypeak = 0
labelinterval = round(datediff("d",startdate,enddate)/25)
if labelinterval<1 then labelinterval = 1
do until datediff("d",enddate,currentdate)>0
	if monthday mod labelinterval = 0 then label = month(currentdate)&"/"&day(currentdate) else label = ""
  if datediff("d",dutcdate,currentdate)=0 then color = rgb(100,100,255) else color = rgb(151,200,255)
  if not rst1.eof and uddusage then
		if year(currentdate)=year(rst1("date")) and month(currentdate)=month(rst1("date")) and day(currentdate)=day(rst1("date")) then
      usage = cdbl(rst1(columntitle))
			rst1.movenext
		else
      usage = 0
		end if
	else
    usage = 0
	end if
  Chart.Series(0).Add usage, label, color
  if usage > ypeak then ypeak = usage
	if not rst2.eof then
		if not(isnull(rst2("deg"))) and year(currentdate)=year(rst2("d")) and month(currentdate)=month(rst2("d")) and day(currentdate)=day(rst2("d")) then
			if not(clng(rst2("deg"))<0) then
        Chart.Series(1).Add clng(rst2("deg")), label, rgb(00,200,00)
				hasNegative = true
			else
        Chart.Series(1).Add abs(clng(rst2("deg"))), label, rgb(200,00,00)
				hasPositive = true
			end if
			rst2.movenext
		end if
	else
    Chart.Series(1).Add 20, label, rgb(00,00,200)
	end if 
	currentdate = dateadd("d",1,currentdate)
	monthday = monthday + 1
loop

if not udddegree then Chart.Series(1).Active = False

if lcase(columntitle)="tons" then columntitle = "Ton/hrs"

Chart.Tools.Add tcAnnotate
Chart.Tools.Items(0).asAnnotation.Text = startdate&" - "&enddate
Chart.Tools.Items(0).asAnnotation.Shape.Transparent = True
Chart.Tools.Items(0).asAnnotation.Shape.Font.Bold = True
Chart.Tools.Items(0).asAnnotation.Shape.Font.Color = rgb(100,100,100)
Chart.Tools.Items(0).asAnnotation.Position = ppRightTop

'SETTING
Chart.Series(0).Title = "Usage"
Chart.Series(0).Color = rgb(151,200,255)
Chart.Series(0).Marks.Visible = False

Chart.Series(1).Title = "Cooling Days"
Chart.Series(1).Color = rgb(00,200,00)
Chart.Series(1).VerticalAxis = aRightAxis
Chart.Series(1).asLine.Pointer.Visible = True
Chart.Series(1).asLine.LinePen.Visible = False

Chart.Series(2).Title = "Heating Days"
Chart.Series(2).Color = rgb(200,00,00)
Chart.Series(2).asLine.Pointer.Visible = True
Chart.Series(2).asLine.LinePen.Visible = False
Chart.Series(0).ShowInLegend = False
Chart.Series(1).ShowInLegend = False
Chart.Series(2).ShowInLegend = False

if ypeak = 0 then
  Chart.Axis.Left.AutomaticMaximum = False
  Chart.Axis.Left.Maximum = 1
end if
Chart.Axis.Left.AutomaticMinimum = False
Chart.Axis.Left.Minimum = 0
Chart.Axis.Left.AxisPen.Width = 1
Chart.Axis.Left.GridPen.Style = psSolid
Chart.Axis.Left.GridPen.color = rgb(100,100,100)
Chart.Axis.Left.MinorTicks.Visible = False
Chart.Axis.Left.Title.Caption = "Usage ("&columntitle&")"
Chart.Axis.Left.Title.Font.Bold = True
Chart.Axis.Right.AxisPen.Width = 1
Chart.Axis.Right.GridPen.Visible = False
Chart.Axis.Right.GridPen.color = rgb(100,100,100)
Chart.Axis.Right.MinorTicks.Visible = False
Chart.Axis.Right.Title.Caption = "Degree Days"
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
Chart.Header.Text.Add "Usage Versus Degree Days"

'setting up legend
Chart.Legend.Transparent = True
'Chart.Legend.Symbol.Width = 6
Chart.Series(0).ShowInLegend = False
Chart.Series(1).ShowInLegend = False
Chart.Series(2).ShowInLegend = False
Chart.AddSeries(scPoint)
Chart.AddSeries(scPoint)
Chart.AddSeries(scPoint)
Chart.Series(3).Title = "Usage"
Chart.Series(3).Color = rgb(151,200,255)
Chart.Series(4).Title = "Heating Days"
Chart.Series(4).Color = rgb(200,00,00)
Chart.Series(5).Title = "Cooling Days"
Chart.Series(5).Color = rgb(00,200,00)

Chart.Aspect.View3D = False
Chart.Panel.Color = vbWhite
Chart.Panel.BevelOuter = bvNone
Chart.Panel.BevelInner = bvNone
Chart.Panel.MarginBottom = 0
Chart.Panel.MarginLeft = 0
Chart.Panel.MarginRight = 0
Chart.Panel.MarginTop = 1
Chart.Width = 650
Chart.Height = 310
Response.BinaryWrite(Chart.Export.asGif.SaveToStream)
Set Chart=nothing
%>
