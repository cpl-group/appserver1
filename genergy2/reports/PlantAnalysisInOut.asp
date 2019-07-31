<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--METADATA NAME="TeeChart Pro v5 ActiveX Control" TYPE="TypeLib" UUID="{B6C10482-FB89-11D4-93C9-006008A7EED4}"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim startdate, enddate, currentdate, groupname, groupid, patrend, pascatter, DBsupermodIP
startdate = request("startdate")
enddate = request("enddate")
groupname = request("groupname")
groupid = request("groupid")
DBsupermodIP = ""
currentdate = startdate
if trim(request("patrend"))="True" then patrend = true else patrend = false
if trim(request("pascatter"))="True" then pascatter = true else pascatter = false

dim rst1, cnn1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getViewConnect(groupid)
cnn1.CommandTimeout = 98000

dim Chart
Set Chart = CreateObject("TeeChart.TChart")
Chart.AddSeries(scPoint)
Chart.AddSeries(scline)

'GETTING COLUMNNAMES
dim inputCol, outputCol

dim ypeak, xpeak, hasdata
hasdata = false
ypeak = 0
xpeak = 0
'response.write "select y, x FROM ["&groupname&"] p WHERE p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"') and y>0 order by p.date<br>"&cnn1
'response.end
rst1.open "select y, x FROM ["&groupname&"] p WHERE p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"') and y>0 order by p.date", cnn1
do until rst1.eof
	Chart.Series(0).AddXY cDbl(rst1("x")), cDbl(rst1("y")),"",rgb(200,00,00)
 	if cdbl(rst1("x"))>xpeak then xpeak = cdbl(rst1("x"))
	if cdbl(rst1("y"))>ypeak then ypeak = cdbl(rst1("y"))
  hasdata = true
	rst1.movenext
loop
rst1.close

if hasdata then
  dim xinterval, yinterval
  xpeak = round(xpeak)
  xinterval = clng((10^(len(xpeak)-1)))
  if ypeak>5 then 
    ypeak = round(ypeak)
    yinterval = clng((10^(len(ypeak)-1)))
    if ypeak/yinterval<4 then yinterval = yinterval/2
  end if
  if xpeak/xinterval<4 then xinterval = xinterval/2
  if yinterval>1 or yinterval=1 then ypeak = (yinterval - ypeak mod yinterval) + ypeak
  xinterval = cint(xinterval)+1
  if xinterval<>0 then xpeak = (xinterval - xpeak mod xinterval) + xpeak
  xinterval = xpeak/10
  
  if pascatter then Chart.Series(0).Active = true else Chart.Series(0).Active = false
  if patrend then Chart.Series(1).Active = true else Chart.Series(1).Active = false
  dim inputlabel, outputlabel
  inputlabel = "Input"
  Outputlabel = "Output"
  rst1.open "select (SELECT isnull(measure,'NA') as measure FROM "&DBsupermodIP&"[group] WHERE id in (SELECT typeid FROM "&DBsupermodIP&"[group] g, "&DBsupermodIP&"groupitems gi WHERE g.id=gi.groupid and typecode='i' and g.groupname='"&groupname&"')) as input, (SELECT isnull(measure,'NA') as measure FROM "&DBsupermodIP&"[group] WHERE id in (SELECT typeid FROM "&DBsupermodIP&"[group] g, "&DBsupermodIP&"groupitems gi WHERE g.id=gi.groupid and typecode='o' and g.groupname='"&groupname&"')) as output", cnn1
  if not rst1.eof then
    inputlabel = rst1("input")
    Outputlabel = rst1("Output")
  end if
  if instr(lcase(outputlabel), "kwh")>0 then outputlabel = "kw"
  rst1.close
else
  'objChart.AddStaticText "No data available for period " & startdate & " - " & enddate, 300, 155, rgb(200,200,200), "Arial", 8, 1, 2
end if

'CURVE FUNCTION
Chart.Series(1).Color = rgb(00,00,200)
Chart.Series(1).SetFunction tfCurveFit
Chart.Series(1).FunctionType.PeriodStyle = psNumPoints
Chart.Series(1).FunctionType.PeriodAlign = paCenter
Chart.Series(1).FunctionType.Period = 50
Chart.Series(1).DataSource = "Series0"

'SETTINGS
Chart.Series(1).asLine.LinePen.Width = 2
Chart.Series(0).asPoint.Pointer.Pen.Visible = False
Chart.Series(0).asPoint.Pointer.HorizontalSize = 4
Chart.Series(0).asPoint.Pointer.VerticalSize = 4
Chart.Series(0).asPoint.Pointer.Style = psCircle

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
Chart.Axis.Left.Title.Caption = Inputlabel&"/"&Outputlabel
Chart.Axis.Left.Title.Font.Bold = True
Chart.Axis.bottom.AutomaticMinimum = False
Chart.Axis.bottom.AutomaticMaximum = False
Chart.Axis.bottom.Minimum = 0
Chart.Axis.bottom.Maximum = xPeak
Chart.Axis.bottom.AxisPen.Width = 1
Chart.Axis.Bottom.GridPen.Visible = False
Chart.Axis.bottom.MinorTicks.Visible = False
Chart.Axis.Bottom.Title.Caption = cstr(Outputlabel)
Chart.Axis.Bottom.Title.Font.Bold = True
Chart.axis.Bottom.Increment = 10
Chart.Header.Text.Clear
Chart.Header.Alignment = taLeftJustify
Chart.Header.Font.Bold = True
Chart.Header.Font.color = rgb(0, 0, 0)
Chart.Header.Text.Add "Efficiency Plot"

Chart.Panel.MarginBottom = 3
Chart.Panel.MarginLeft = 3
Chart.Panel.MarginRight = 3
Chart.Panel.MarginTop = 3
Chart.Width = 600
Chart.Height = 310
Set rst1=nothing
Set cnn1=nothing
Response.BinaryWrite(Chart.Export.asGif.SaveToStream)
Set Chart=nothing
%>
