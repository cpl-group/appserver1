<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim startdate, enddate, currentdate, interval, total, groupname
startdate = request("startdate")
enddate = request("enddate")
groupname = request("groupname")
currentdate = startdate

dim rst1, cnn1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")
'cnn1.CursorLocation = adUseClient

dim objChart
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

interval = 250
'getting columnnames
'dim inputCol, outputCol
'rst1.open "SELECT top 1 * FROM ["&input&"]", cnn1
'inputCol = rst1.fields.Item(0).Name
'rst1.close
'rst1.open "SELECT top 1 * FROM ["&output&"]", cnn1
'outputCol = rst1.fields.Item(0).Name
'rst1.close

dim ypeak, xpeak, hasdata
hasdata = false
ypeak = 1
rst1.open "SELECT isnull(max(x),1) as xPeak FROM ["&groupname&"] p WHERE p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"')", cnn1
xpeak = cdbl(rst1("xpeak"))
xpeak = round(xpeak)
xinterval = clng((10^(len(xpeak)-1)))
if interval <> 0 then if xpeak/xinterval<4 then xinterval = xinterval/2
if xinterval > 1 then 
  xpeak = (xinterval - xpeak mod xinterval) + xpeak
  interval = round(xpeak/10^(len(round(xpeak))-2))*10^(len(round(xpeak))-2)
  interval = round(interval/10)
end if
rst1.close

rst1.open "SELECT isnull(sum(y),0) as total FROM ["&groupname&"] p WHERE p.y>=0 and p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"')", cnn1
if not rst1.eof then total = cdbl(rst1("total"))
rst1.close

rst1.open "select sum(y) as y, floor(x/"&interval&") as levels FROM ["&groupname&"] p WHERE p.y>=0 and p.x>=0 and p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"') GROUP BY floor(x/"&interval&")", cnn1
dim level
level = 0
do until level >= xpeak
  if not rst1.eof then
    if interval*cint(rst1("levels")) = level then
      objChart.AddData cdbl(rst1("y")), 0, interval*cint(rst1("levels"))&"-"&interval*(cint(rst1("levels"))+1)
    	if total <> 0 then objChart.AddData 0, 1, formatpercent(cdbl(rst1("y"))/total,0) else objChart.AddData 0, 1, "0%"
    	if cdbl(rst1("y"))>ypeak then ypeak = cdbl(rst1("y"))
      hasdata = true
    	rst1.movenext
    else
      objChart.AddData 0, 0, level&"-"&interval+level
    	objChart.AddData 0, 1, "0%"
    end if
  else
    objChart.AddData 0, 0, level&"-"&interval+level
  	objChart.AddData 0, 1, "0%"
  end if
  level = level + interval
loop
rst1.close

if hasdata then
  dim xinterval, yinterval
  ypeak = round(ypeak)
  yinterval = clng((10^(len(ypeak)-1)))
  if ypeak/yinterval<4 then yinterval = cdbl(yinterval)/2
  if yinterval<1 then yinterval = 1
  ypeak = (yinterval - ypeak mod yinterval) + ypeak
  
  objChart.ChartArea(0).AddChart 7,0,0
  objChart.ChartArea(0).Transparent = true
  objChart.ChartArea(0).GridVEnabled = false
  objChart.ChartArea(0).SetPosition 70,30,550,270
  objChart.ChartArea(0).Axis(0).Maximum = yPeak
  objChart.ChartArea(0).Axis(2).Maximum = yPeak
  objChart.ChartArea(0).Axis(0).SetNumberFormat 1, 0
  'objChart.ChartArea(0).Axis(0).Interval = yinterval
  objChart.ChartArea(0).Axis(0).FontSize = 8
  objChart.ChartArea(0).Axis(1).FontSize = 8
  objChart.ChartArea(0).Axis(0).TruncatedLabels = true
  objChart.ChartArea(1).AddChart 7,1,1
  objChart.ChartArea(1).Axis(3).FontSize = 8
  objChart.ChartArea(1).GridHEnabled = false
  objChart.ChartArea(1).GridVEnabled = false
  objChart.ChartArea(1).Transparent = true
  objChart.ChartArea(1).Axis(0).enabled=false
  objChart.ChartArea(1).Axis(1).enabled=false
  objChart.ChartArea(1).SetPosition 70,30,550,270
  objChart.ChartArea(1).Axis(3).enabled=true
  objChart.ChartArea(0).Axis(1).angle=90
  
  objChart.SetSeriesColor 0, rgb(200,00,00)
  objChart.SetSeriesColor 2, rgb(00,200,00)
  
  'objChart.Legend.Enabled = true
  'objChart.Legend.add "Cooling", rgb(00,200,00)
  'objChart.Legend.add "Heating", rgb(200,00,00)
  'objChart.Legend.SetPosition 550,265,600,300
  objChart.AddStaticText "Run Hours",20,162,RGB(100,100,100),"Arial",8,1,2,90
  dim Outputlabel
  Outputlabel = "Output"
  rst1.open "SELECT distinct measure FROM meters m INNER JOIN tblleasesutilityprices lup ON m.leaseutilityid=lup.leaseutilityid INNER JOIN tblutility u ON u.utilityid=lup.utility WHERE meterid in (SELECT typeid FROM groupitems WHERE groupid in (SELECT typeid FROM [group] g, groupitems gi WHERE g.id=gi.groupid AND g.groupname='"&groupname&"' AND typecode='o'))", cnn1
  do until rst1.eof
    if Outputlabel<>"Output" and trim(rst1("measure"))<>Outputlabel then Outputlabel = "BTU" else Outputlabel = trim(rst1("measure"))
    rst1.movenext
  loop
  rst1.close
  objChart.AddStaticText outputlabel,300,335,RGB(100,100,100),"Arial",8,1,2
  objChart.AddStaticText "Load Hour Report",70,2,RGB(100,100,100),"Arial",8,1,0

else
  objChart.AddStaticText "No data available for period " & startdate & " - " & enddate, 300, 155, rgb(200,200,200), "Arial", 8, 1, 2
end if
objChart.SendJPEG 600, 350
%>
