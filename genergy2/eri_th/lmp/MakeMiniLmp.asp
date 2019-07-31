<%option explicit

dim data
dim dataset, mday, peak
peak = (clng(request.querystring("scale"))/1)+10
data = request.querystring("data")
mday = request.querystring("day")
dataset = split(data,",")

'make chart
dim objChart, bit
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
for each bit in dataset
	dim label, tempdata
	tempdata = 0
	label = ""
	objChart.AddData bit, 0, , RGB(00,153,255)
next
objChart.SetColorFromPoint (0)
objChart.ChartArea(0).AddChart 1,0,0
objChart.ChartArea(0).Axis(1).Angle = 0
objChart.ChartArea(0).Axis(0).enabled=false
objChart.ChartArea(0).Axis(1).enabled=false
objChart.ChartArea(0).Axis(2).enabled=false
objChart.ChartArea(0).Axis(0).color=RGB(255,255,255)
objChart.ChartArea(0).Axis(1).color=RGB(255,255,255)
objChart.ChartArea(0).Axis(2).color=RGB(255,255,255)
objChart.ChartArea(0).Axis(3).color=RGB(255,255,255)
if peak>10 then
	objChart.ChartArea(0).Axis(0).Maximum = Peak
	objChart.ChartArea(0).Axis(2).Maximum = Peak
end if
objChart.ChartArea(0).LineWidth = 3
objChart.ChartArea(0).GridVEnabled = false
objChart.ChartArea(0).GridHEnabled = false
objChart.AntiAlias()
objChart.ChartArea(0).BackgroundColor = RGB(255,255,255)
objChart.ChartArea(0).SetPosition 0,0,140,80
objChart.AddStaticText mday,0,48,RGB(00,00,00),"Arial",20,0,0
objChart.SendJpeg 140,80
%>
