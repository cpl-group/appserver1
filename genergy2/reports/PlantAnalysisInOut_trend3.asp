<%option explicit%>
<!--METADATA NAME="TeeChart Pro v5 ActiveX Control" TYPE="TypeLib" UUID="{B6C10482-FB89-11D4-93C9-006008A7EED4}"-->
<%
dim startdate, enddate, groupname
startdate = request("startdate")
enddate = request("enddate")
groupname = request("groupname")

dim rst1, cnn1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")
rst1.open "SELECT isnull(x,0) as x, isnull(y,0) as y FROM ["&groupname&"] p WHERE p.x>=0 and p.date>='"&startdate&"' and p.date<=dateadd(day,1,'"&enddate&"') and y>0 ORDER BY x", cnn1

dim chart, xpeak
Set Chart = CreateObject("TeeChart.TChart")
Chart.Aspect.View3D = False
Chart.Legend.visible = false
Chart.Panel.Color = vbWhite
Chart.Panel.BevelOuter = bvNone
Chart.Panel.BevelInner = bvNone
Chart.AddSeries(scPoint)
Chart.Axis.Left.AutomaticMinimum = False
Chart.Axis.Left.Minimum = 0
Chart.Axis.Left.AxisPen.Width = 1
Chart.Axis.Left.GridPen.Style = psSolid
Chart.Axis.Left.GridPen.color = vbBlack
Chart.Axis.Left.MinorTicks.Visible = False
Chart.Axis.bottom.AutomaticMinimum = False
Chart.Axis.bottom.Minimum = 0
Chart.Axis.bottom.AxisPen.Width = 1
Chart.Axis.Bottom.GridPen.Visible = False
Chart.Axis.bottom.MinorTicks.Visible = False
Chart.Header.Visible = False
Chart.Panel.MarginBottom = 3
Chart.Panel.MarginLeft = 7
Chart.Panel.MarginRight = 8
Chart.Panel.MarginTop = 3
Chart.Width = 600
Chart.Height = 310

Chart.Series(0).asPoint.Pointer.HorizontalSize = 2
Chart.Series(0).asPoint.Pointer.VerticalSize = 2
Chart.Series(0).asPoint.Pointer.Style = psRectangle

do until rst1.eof
  Chart.Series(0).AddXY cDbl(rst1("x")), cDbl(rst1("y")),"",rgb(200,00,00)
  if cint(rst1("x")) > xpeak then xpeak = cint(rst1("x"))
  rst1.movenext
loop

dim xinterval
xpeak = round(xpeak)
xinterval = clng((10^(len(xpeak)-1)))
if xpeak/xinterval<4 then xinterval = xinterval/2
xpeak = (xinterval - xpeak mod xinterval) + xpeak
xinterval = xpeak/10

Chart.Axis.bottom.AutomaticMaximum = False
Chart.Axis.bottom.Maximum = xpeak

Chart.AddSeries(scline)
Chart.Series(1).SetFunction tfCurveFit
Chart.Series(1).FunctionType.PeriodStyle = psNumPoints
Chart.Series(1).FunctionType.PeriodAlign = paCenter
Chart.Series(1).FunctionType.Period = 50
Chart.Series(1).DataSource = "Series0"
Chart.Series(1).Color = rgb(00,00,200)
Chart.Series(1).asLine.LinePen.Width = 2

rst1.close
cnn1.close
Set rst1=nothing
Set cnn1=nothing
Response.BinaryWrite(Chart.Export.asGif.SaveToStream)
Set Chart=nothing


%>
