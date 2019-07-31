<%option explicit
dim data
dim dataset
data = request.querystring("a")
dataset = split(data,",")

'make chart
dim objChart, i
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
for i=0 to 11
	dim label, tempdata
	tempdata = 0
	label = ""
	if i mod 2 = 0 then label = left(monthname(i+1),3)
	if trim(dataset(i))<>"" then tempdata=dataset(i)
'	if tempdata<0 then
		objChart.AddData tempdata, 0, label , RGB(255,0,0)
'	else
'		objChart.AddData tempdata, 0, label , RGB(0,255,0)
'	end if
next
objChart.SetColorFromPoint (0)
objChart.ChartArea(0).AddChart 1,0,0
objChart.ChartArea(0).Axis(1).Angle = 0
objChart.ChartArea(0).SetPosition 50,18,299,90
objChart.ChartArea(0).Axis(0).SetNumberFormat 1,0
objChart.ChartArea(0).Axis(2).SetNumberFormat 1,0
objChart.AddStaticText "Net Gain",50,4,RGB(100,100,100),"Arial",8,1
objChart.SendJpeg 300,110
%>
