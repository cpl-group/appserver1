<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--METADATA NAME="TeeChart Pro v5 ActiveX Control" TYPE="TypeLib" UUID="{B6C10482-FB89-11D4-93C9-006008A7EED4}"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
server.ScriptTimeout = 210
dim groupname, startdate, enddate, columntitle, wday, label, ckwh, holiday, where, title, DBsupermodIP, groupid
startdate = request("startdate")
enddate = request("enddate")
wday = request("day")
holiday = request("holiday")
groupname = request("groupname")
groupid = request("groupid")
DBsupermodIP = ""

dim rst1, cnn1, cmd, prm, Chart
set cmd = server.createobject("ADODB.command")
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getViewConnect(groupid)
cnn1.CursorLocation = adUseClient

dim groupCol, hours, peakdemand, pdDate, loadFactor, avgTemp, peak, decimals
rst1.open "SELECT top 1 * FROM ["&groupname&"]", cnn1
groupCol = rst1.fields.Item(0).Name
rst1.close

Set Chart = CreateObject("TeeChart.TChart")
Chart.AddSeries(scline)

hours = 0
peak = 0
decimals = 0
if trim(holiday)<>"" or trim(wday)<>"" then
  if trim(holiday)<>"" then
    where = "convert(datetime,left(lmp.date,11)) in (SELECT date from "&DBsupermodIP&"holidaysch WHERE holiday='"& join(split(holiday,","),"' or holiday='") &"')"
    title = "Selected Holiday"
  else
    where = "(datepart(weekday,lmp.date)="& join(split(wday,",")," or datepart(weekday,lmp.date)=") &") and left(lmp.date,11) not in (SELECT DISTINCT left(date,11) FROM "&DBsupermodIP&"holidaysch)"
    if instr(wday,"1")<>0 or instr(wday,"7")<>0 then title = "Selected Weekend Days" else title = "Selected Weekdays"
'response.write where&"<br>"
  end if
  'get peakdemand
  'response.write "SELECT top 1 datepart(hour,date) as hours, datepart(minute,date) as minutes, avg(isnull(lmp.["&groupCol&"],0)) as lmp1, avg(isnull(lmp2.innerusage,0)) as lmp2, (avg(isnull(lmp.["&groupCol&"],0))+avg(isnull(lmp2.innerusage,0)))/2 as pd FROM ["&groupname&"] lmp LEFT JOIN (SELECT dateadd(n,-15,il.date) as innerdate, il.["&groupCol&"] as innerusage FROM ["&groupname&"] il WHERE date>'"&startdate&"' and date<'"&enddate&"') lmp2 ON lmp2.innerdate=lmp.date WHERE ("&where&") and lmp.date>'"&startdate&"' and lmp.date<'"&enddate&"' GROUP BY datepart(hour,date), datepart(minute,date) ORDER BY pd desc"&"<br>"& cnn1 
  'response.end
  rst1.open "SELECT top 1 datepart(hour,date) as hours, datepart(minute,date) as minutes, avg(isnull(lmp.["&groupCol&"],0)) as lmp1, avg(isnull(lmp2.innerusage,0)) as lmp2, (avg(isnull(lmp.["&groupCol&"],0))+avg(isnull(lmp2.innerusage,0)))/2 as pd FROM ["&groupname&"] lmp LEFT JOIN (SELECT dateadd(n,-15,il.date) as innerdate, il.["&groupCol&"] as innerusage FROM ["&groupname&"] il WHERE date>'"&startdate&"' and date<'"&enddate&"') lmp2 ON lmp2.innerdate=lmp.date WHERE ("&where&") and lmp.date>'"&startdate&"' and lmp.date<'"&enddate&"' GROUP BY datepart(hour,date), datepart(minute,date) ORDER BY pd desc", cnn1
  if not rst1.eof then
    if cDbl(rst1("pd"))<>0 then peakdemand = cDbl(rst1("pd")) else peakdemand = 1
    pdDate = rst1("hours")&":"&rst1("minutes")
  else
    peakdemand = 1
    pdDate = 0
  end if
  rst1.close
  'get load factor and average temperature
  rst1.open "SELECT case when datediff(hour,'"&startdate&"','"&enddate&"')=0 then 0 else (isnull(sum(["&groupCol&"])/4,0)/"&peakdemand&")/datediff(hour,'"&startdate&"','"&enddate&"') end as loadfactor, isnull(avg(isnull(dd.ou_t,0)),0) as avgTemp FROM ["&groupname&"] lmp FULL OUTER JOIN dbo.deg_day dd ON lmp.date=dd.date WHERE ("&where&") and lmp.date>'"&startdate&"' and lmp.date<'"&enddate&"'", cnn1
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
    Chart.Series(0).Add ckwh, hours, rgb(200,00,00)
    hours = hours + 1
  loop
end if
Chart.Tools.Add tcAnnotate
Chart.Tools.Items(0).asAnnotation.Text = "Peak Demand: "&formatnumber(peakdemand) &" "&groupCol&vbCr&_
                                         "Load Factor: "&formatpercent(loadFactor,2)&vbCr&_
                                         "Average Temp: "&formatnumber(avgTemp)
Chart.Tools.Items(0).asAnnotation.Shape.Transparent = False
Chart.Tools.Items(0).asAnnotation.Shape.Font.Bold = True
Chart.Tools.Items(0).asAnnotation.Shape.Top = 106
Chart.Tools.Items(0).asAnnotation.Shape.Left = 350
Chart.Tools.Items(0).asAnnotation.Shape.ShadowSize = 0

Chart.Aspect.View3D = False
Chart.Legend.visible = false
Chart.Panel.Color = vbWhite
Chart.Panel.BevelOuter = bvNone
Chart.Panel.BevelInner = bvNone

Chart.Axis.Left.AutomaticMinimum = False
Chart.Axis.Left.Minimum = 0
Chart.Axis.Left.AxisPen.Width = 1
Chart.Axis.Left.GridPen.Style = psSolid
Chart.Axis.Left.GridPen.color = vbBlack
Chart.Axis.Left.MinorTicks.Visible = False
Chart.Axis.Left.Title.Caption = "Usage (" & groupCol & ")"
Chart.Axis.Left.Title.Font.Bold = True
Chart.Axis.bottom.AutomaticMinimum = False
Chart.Axis.bottom.Minimum = 0
Chart.Axis.bottom.AxisPen.Width = 1
Chart.Axis.Bottom.GridPen.Visible = False
Chart.Axis.bottom.MinorTicks.Visible = False
Chart.Axis.Bottom.Title.Caption = "Hour"
Chart.Axis.Bottom.Title.Font.Bold = True
Chart.Axis.Bottom.Increment = 10

Chart.Header.Text.Clear
Chart.Header.Alignment = taRightJustify
Chart.Header.Font.Bold = True
Chart.Header.Font.color = rgb(0, 0, 0)
Chart.Header.Text.Add "Average Usage For " & title & " " & startdate &"-"&enddate

Chart.Panel.MarginBottom = 3
Chart.Panel.MarginLeft = 3
Chart.Panel.MarginRight = 3
Chart.Panel.MarginTop = 3
Chart.Width = 525
Chart.Height = 195
Set rst1=nothing
Set cnn1=nothing
Response.BinaryWrite(Chart.Export.asGif.SaveToStream)
Set Chart=nothing
%>