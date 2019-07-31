<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim startdate, enddate, currentdate, groupname, patrend, pascatter
startdate = request("startdate")
enddate = request("enddate")
groupname = request("groupname")
currentdate = startdate
if trim(request("patrend"))="True" then patrend = true else patrend = false
if trim(request("pascatter"))="True" then pascatter = true else pascatter = false

dim rst1, cnn1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")
'cnn1.CursorLocation = adUseClient

dim objChart
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

'getting columnnames
dim inputCol, outputCol
'rst1.open "SELECT top 1 * FROM ["&input&"]", cnn1
'inputCol = rst1.fields.Item(0).Name
'rst1.close
'rst1.open "SELECT top 1 * FROM ["&output&"]", cnn1
'outputCol = rst1.fields.Item(0).Name
'rst1.close

dim ypeak, xpeak, hasdata
hasdata = false
ypeak = 0
xpeak = 0
rst1.open "select y, x FROM ["&groupname&"] p WHERE p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"') and y>0 order by p.date", cnn1
do until rst1.eof
	if pascatter then
    objChart.AddData cdbl(rst1("x")), 0
  	objChart.AddData cdbl(rst1("y")), 1
  end if
 	if cdbl(rst1("x"))>xpeak then xpeak = cdbl(rst1("x"))
	if cdbl(rst1("y"))>ypeak then ypeak = cdbl(rst1("y"))
  hasdata = true
	rst1.movenext
loop
rst1.close

dim totalpoints, interval, trendpoints, minPointsPerAvg

rst1.open "SELECT isnull(count(x),0) as totalpoints FROM ["&groupname&"] p WHERE p.x>=0 and p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"') and y>0", cnn1
'SELECT x as x, y as y FROM [Plant_occc_lock] p WHERE p.x>=0 and p.date>='1/1/2003' and p.date<=dateadd(day,1,'1/7/2003') and y>0 ORDER BY x", cnn1
totalpoints = cInt(rst1("totalpoints"))
rst1.close

trendpoints = 15
minPointsPerAvg = 40
if totalpoints>0 then
  do while totalpoints/trendpoints<minPointsPerAvg and trendpoints<>0
    trendpoints = trendpoints - 1
  loop
  interval = totalpoints/trendpoints
else
  trendpoints = 0
end if
if trendpoints>1 then
dim xtemp, ytemp
rst1.open "SELECT isnull(x,0) as x, isnull(y,0) as y FROM ["&groupname&"] p WHERE p.x>=0 and p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"') and y>0 ORDER BY x", cnn1
do until rst1.eof
  xtemp=0
  ytemp=0
  trendpoints=0
  do until trendpoints>interval or rst1.eof
    xtemp = xtemp + cDbl(rst1("x"))
    ytemp = ytemp + cDbl(rst1("y"))
    rst1.movenext
    trendpoints = trendpoints + 1
  loop
  objChart.AddData cdbl(xtemp/trendpoints), 2
  objChart.AddData cdbl(ytemp/trendpoints), 3
loop
rst1.close
end if

if hasdata then
  dim xinterval, yinterval
  xpeak = round(xpeak)
  xinterval = clng((10^(len(xpeak)-1)))
  if ypeak>5 then 
    ypeak = round(ypeak)
    yinterval = clng((10^(len(ypeak)-1)))
    if ypeak/yinterval<4 then yinterval = yinterval/2
  end if
'  response.write ypeak
'  response.end
  if xpeak/xinterval<4 then xinterval = xinterval/2
  if yinterval>1 or yinterval=1 then ypeak = (yinterval - ypeak mod yinterval) + ypeak
  xpeak = (xinterval - xpeak mod xinterval) + xpeak
  xinterval = xpeak/10
  if pascatter then
    objChart.ChartArea(0).AddChart 3,0,1
    objChart.ChartArea(0).Transparent = true
    objChart.ChartArea(0).GridVEnabled = false
    objChart.ChartArea(0).SetPosition 70,30,550,270
    objChart.ChartArea(0).Axis(0).Maximum = yPeak
    if ypeak>20 or ypeak=20 then
      objChart.ChartArea(0).Axis(0).Interval = yinterval
    end if
    objChart.ChartArea(0).Axis(1).Maximum = xPeak
    objChart.ChartArea(0).Axis(3).Maximum = xPeak
    objChart.ChartArea(0).Axis(0).Minimum = 0
    objChart.ChartArea(0).Axis(2).Minimum = 0
    objChart.ChartArea(0).Axis(1).Minimum = 0
    objChart.ChartArea(0).Axis(3).Minimum = 0
    objChart.ChartArea(0).Axis(0).SetNumberFormat 1,2
    objChart.ChartArea(0).Axis(1).Interval = xinterval
    objChart.ChartArea(0).Axis(0).FontSize = 8
    objChart.ChartArea(0).Axis(1).FontSize = 8
    objChart.ChartArea(0).Axis(0).TruncatedLabels = true
'    objChart.ChartArea(0).Axis(0).SetNumberFormat 1, 1
    objChart.ChartArea(0).Axis(1).SetNumberFormat 1, 0
  end if
  
  objChart.SetSeriesColor 0, rgb(200,00,00)
  objChart.SetSeriesColor 2, rgb(00,00,200)

  if patrend then
    objChart.ChartArea(1).AddChart 2,2,3
    objChart.ChartArea(1).Transparent = true
    objChart.ChartArea(1).GridVEnabled = false
    objChart.ChartArea(1).SetPosition 70,30,550,270
    objChart.ChartArea(1).Axis(1).Maximum = xPeak
    objChart.ChartArea(1).Axis(0).Minimum = 0
    objChart.ChartArea(1).Axis(1).Minimum = 0
    objChart.ChartArea(1).Axis(0).SetNumberFormat 1,2
    objChart.ChartArea(1).Axis(0).Maximum = yPeak
    objChart.ChartArea(1).Axis(1).Maximum = xPeak
    if ypeak>20 or ypeak=20 then
      objChart.ChartArea(1).Axis(0).Interval = yinterval
    end if
    objChart.ChartArea(1).Axis(0).enabled = true
    objChart.ChartArea(1).Axis(1).enabled = true
    objChart.ChartArea(1).Axis(0).FontSize = 8
    objChart.ChartArea(1).Axis(1).FontSize = 8
    objChart.ChartArea(1).Axis(1).Interval = xinterval
'    objChart.ChartArea(1).Axis(0).SetNumberFormat 1, 1
    objChart.ChartArea(1).Axis(1).SetNumberFormat 1, 0
    objChart.ChartArea(1).LineWidth = 4
  end if
'  objChart.AntiAlias

  'objChart.Legend.Enabled = true
  'objChart.Legend.add "Cooling", rgb(00,200,00)
  'objChart.Legend.add "Heating", rgb(200,00,00)
  'objChart.Legend.SetPosition 550,265,600,300
  dim inputlabel, outputlabel
  inputlabel = "Input"
  Outputlabel = "Output"
  rst1.open "SELECT distinct measure FROM meters m INNER JOIN tblleasesutilityprices lup ON m.leaseutilityid=lup.leaseutilityid INNER JOIN tblutility u ON u.utilityid=lup.utility WHERE meterid in (SELECT typeid FROM groupitems WHERE groupid in (SELECT typeid FROM [group] g, groupitems gi WHERE g.id=gi.groupid AND g.groupname='"&groupname&"' AND typecode='i'))", cnn1
  do until rst1.eof
    if inputlabel<>"Input" and trim(rst1("measure"))<>inputlabel then inputlabel = "BTU" else inputlabel = trim(rst1("measure"))
    rst1.movenext
  loop
  if instr(lcase(inputlabel), "kwh")>0 then inputlabel = "kw"
  rst1.close
  rst1.open "SELECT distinct measure FROM meters m INNER JOIN tblleasesutilityprices lup ON m.leaseutilityid=lup.leaseutilityid INNER JOIN tblutility u ON u.utilityid=lup.utility WHERE meterid in (SELECT typeid FROM groupitems WHERE groupid in (SELECT typeid FROM [group] g, groupitems gi WHERE g.id=gi.groupid AND g.groupname='"&groupname&"' AND typecode='o'))", cnn1
  do until rst1.eof
    if Outputlabel<>"Output" and trim(rst1("measure"))<>Outputlabel then Outputlabel = "BTU" else Outputlabel = trim(rst1("measure"))
    rst1.movenext
  loop
  if instr(lcase(outputlabel), "kwh")>0 then outputlabel = "kw"
  rst1.close

  objChart.AddStaticText Inputlabel&"/"&Outputlabel,20,155,RGB(100,100,100),"Arial",8,1,2,90
  objChart.AddStaticText Outputlabel,300,295,RGB(100,100,100),"Arial",8,1,2
  objChart.AddStaticText "Efficiency Plot",70,12,RGB(100,100,100),"Arial",8,1,0
else
  objChart.AddStaticText "No data available for period " & startdate & " - " & enddate, 300, 155, rgb(200,200,200), "Arial", 8, 1, 2
end if
objChart.SendJPEG 600, 310
%>
