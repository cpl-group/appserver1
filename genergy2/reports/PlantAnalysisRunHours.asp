<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--METADATA NAME="TeeChart Pro v5 ActiveX Control" TYPE="TypeLib" UUID="{B6C10482-FB89-11D4-93C9-006008A7EED4}"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim startdate, enddate, currentdate, interval, total, groupname, groupid, DBsupermodIP
startdate = request("startdate")
enddate = request("enddate")
groupname = request("groupname")
groupid = request("groupid")
currentdate = startdate
DBsupermodIP = ""

dim rst1, cnn1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getViewConnect(groupid)
cnn1.CommandTimeout = 98000

dim Chart
Set Chart = CreateObject("TeeChart.TChart")
Chart.AddSeries(scBar)

interval = 250

dim ypeak, xpeak, hasdata, xinterval, yinterval
hasdata = false
ypeak = 1
rst1.open "SELECT isnull(max(x),1) as xPeak FROM ["&groupname&"] p WHERE p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"')", cnn1
xpeak = cdbl(rst1("xpeak"))
xpeak = round(xpeak)
xinterval = clng((10^(len(xpeak)-1)))
if interval <> 0 then if xpeak/xinterval<4 then xinterval = xinterval/2
if xinterval >= 1 then 
  xpeak = (xinterval - xpeak mod xinterval) + xpeak
  interval = round(xpeak/10^(len(round(xpeak))-2))*10^(len(round(xpeak))-2)
  interval = round(interval/10)
  if interval = 0 then interval = 1
end if
rst1.close
'response.write "SELECT isnull(sum(y),0) as total FROM ["&groupname&"] p WHERE p.y>=0 and p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"')"
'response.end
rst1.open "SELECT isnull(sum(y),0) as total FROM ["&groupname&"] p WHERE p.y>=0 and p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"')", cnn1
if not rst1.eof then total = cdbl(rst1("total"))
rst1.close

rst1.open "select sum(y) as y, floor(x/"&interval&") as levels FROM ["&groupname&"] p WHERE p.y>=0 and p.x>=0 and p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"') GROUP BY floor(x/"&interval&")", cnn1

dim level
level = 0
do until level >= xpeak
  if not rst1.eof then
    if interval*cint(rst1("levels")) = level then
      Chart.Series(0).Add cDbl(rst1("y")), interval*cint(rst1("levels"))&"-"&interval*(cint(rst1("levels"))+1), rgb(200,00,00)
    	if cdbl(rst1("y"))>ypeak then ypeak = cdbl(rst1("y"))
      hasdata = true
    	rst1.movenext
    else
    Chart.Series(0).Add 0, level&"-"&interval+level, rgb(200,00,00)
    end if
  else
    Chart.Series(0).Add 0, level&"-"&interval+level, rgb(200,00,00)
  end if
  level = level + interval
loop
rst1.close

if hasdata then
  dim Outputlabel, inputlabel
  Outputlabel = "Output"
  rst1.open "select (SELECT measure FROM "&DBsupermodIP&"[group] WHERE id in (SELECT typeid FROM "&DBsupermodIP&"[group] g, "&DBsupermodIP&"groupitems gi WHERE g.id=gi.groupid and typecode='i' and g.groupname='PA_Test2')) as input, (SELECT measure FROM "&DBsupermodIP&"[group] WHERE id in (SELECT typeid FROM "&DBsupermodIP&"[group] g, "&DBsupermodIP&"groupitems gi WHERE g.id=gi.groupid and typecode='o' and g.groupname='"&groupname&"')) as output", cnn1
  if not rst1.eof then
    inputlabel = rst1("input")
    Outputlabel = rst1("Output")
  end if
  rst1.close
else
'  objChart.AddStaticText "No data available for period " & startdate & " - " & enddate, 300, 155, rgb(200,200,200), "Arial", 8, 1, 2
end if

'SETTINGS
'Chart.Series(0).asBar.OffsetPercent = -50
Chart.Series(0).Marks.Style = smsPercent
Chart.Series(0).Marks.Frame.Visible = False
Chart.Series(0).Marks.BackColor = vbWhite
Chart.Series(0).Marks.ShadowSize = 0
Chart.Series(0).Marks.Arrow.Visible = False
Chart.Series(0).Marks.ArrowLength = 3

Chart.Aspect.View3D = False
Chart.Legend.visible = false
Chart.Panel.Color = vbWhite
Chart.Panel.BevelOuter = bvNone
Chart.Panel.BevelInner = bvNone

Chart.Axis.Left.AxisPen.Width = 1
Chart.Axis.Left.GridPen.Style = psSolid
Chart.Axis.Left.GridPen.color = vbBlack
Chart.Axis.Left.MinorTicks.Visible = False
Chart.Axis.Left.Title.Caption = Outputlabel
Chart.Axis.Left.Title.Font.Bold = True
Chart.Axis.bottom.AxisPen.Width = 1
Chart.Axis.Bottom.GridPen.Visible = False
Chart.Axis.bottom.MinorTicks.Visible = False
Chart.Axis.Bottom.Title.Caption = "Run Hours"
Chart.Axis.Bottom.Title.Font.Bold = True
Chart.Header.Text.Clear
Chart.Header.Alignment = taLeftJustify
Chart.Header.Font.Bold = True
Chart.Header.Font.color = rgb(0, 0, 0)
Chart.Header.Text.Add "Load Hour Report"

Chart.Panel.MarginBottom = 3
Chart.Panel.MarginLeft = 3
Chart.Panel.MarginRight = 3
Chart.Panel.MarginTop = 3
Chart.Width = 600
Chart.Height = 350
Set rst1=nothing
Set cnn1=nothing
Response.BinaryWrite(Chart.Export.asGif.SaveToStream)
Set Chart=nothing
%>
