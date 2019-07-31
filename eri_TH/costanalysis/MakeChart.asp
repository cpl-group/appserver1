<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE file="revFunctions.asp"-->
<%
dim date1, date2, b, utype
b = request.querystring("b")
utype = request.querystring("utype")
date1 = request.querystring("date1")
date2 = request.querystring("date2")

'get building name (cnn1 is defined in reFunctions.asp)
dim rst2, bname
set rst2 = server.createobject("ADODB.recordset")
rst2.open "SELECT strt from buildings WHERE bldgNum='"& b &"'", cnn1
if not(rst2.eof) then
	bname = rst2("strt")
else
	bname = ""
end if
rst2.close

dim ArrDataSeriesERI(12)
dim ArrDataSeriesSubMetered(12)
dim ArrDataSeriesExpenses(12)
dim ArrDataSeriesUnreportedExp(12)
dim ArrDataSeriesUnreportedRev(12)
dim ArrDataSeriesMac(12)
dim ArrDataSeriesPLP(12)
dim ArrDataSeriesNet(12)
Dim ArrPieRevenue(3)
Dim ArrPieExpenses(3)
dim peak, chartH
chartH = 180
checkprefs()


call getdataSets(date1, b, utype, "0")', ArrDataSeriesERI, ArrDataSeriesSubMetered, ArrDataSeriesExpenses, ArrDataSeriesUnreportedExp, ArrDataSeriesUnreportedRev, ArrPieRevenue, ArrPieExpenses)


'make chart
dim objChart, i
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
for i = 1 to 12
	objChart.AddData abs(ArrDataSeriesExpenses(i)), 0, , RGB(255,0,0)
	objChart.AddData abs(ArrDataSeriesUnreportedExp(i)), 1, , RGB(150,0,0)
	objChart.AddData ArrDataSeriesERI(i), 4, left(monthname(i),3), RGB(0,0,255)
	objChart.AddData ArrDataSeriesSubMetered(i), 5, , RGB(51,51,200)
	objChart.AddData ArrDataSeriesUnreportedRev(i), 6, , RGB(103,103,150)
	objChart.AddData ArrDataSeriesMac(i), 7, , RGB(103,150,150)
	objChart.AddData ArrDataSeriesPLP(i), 8, , RGB(150,103,150)
next
objChart.SetColorFromPoint (0)
objChart.SetColorFromPoint (1)
objChart.SetColorFromPoint (4)
objChart.SetColorFromPoint (5)
objChart.SetColorFromPoint (6)
objChart.SetColorFromPoint (7)
objChart.SetColorFromPoint (8)
objChart.ChartArea(0).AddChart 7,4,8
objChart.ChartArea(1).AddChart 7,0,1
objChart.ChartArea(1).Transparent = true
objChart.ChartArea(0).GridVEnabled = false
objChart.ChartArea(0).GridHEnabled = false
objChart.ChartArea(1).GridVEnabled = false
objChart.ChartArea(1).GridHEnabled = false
objChart.ChartArea(0).Axis(0).Enabled=true
objChart.ChartArea(1).Axis(2).Enabled=true
objChart.ChartArea(0).Axis(2).Enabled=false
objChart.ChartArea(1).Axis(0).Enabled=false
objChart.ChartArea(1).Axis(1).Enabled=false
objChart.ChartArea(0).SetPosition 40,20,460,150
objChart.ChartArea(1).SetPosition 50,20,470,150
objChart.AddStaticText "Revenue Profile for "&date1,40,6,RGB(100,100,100),"Arial",8,1
objChart.AddStaticText bname,469,6,RGB(100,100,100),"Arial",8,1,1
objChart.AddStaticText "Dollar in Thousands",0,chartH/2,RGB(100,100,100),"Arial",8,1,2,90
peak = findpeak()
objChart.ChartArea(0).Axis(0).Maximum = Peak
objChart.ChartArea(1).Axis(2).Maximum = Peak
objChart.ChartArea(0).Axis(2).Maximum = Peak
objChart.ChartArea(1).Axis(0).Maximum = Peak
dim dpoints
if peak < 6 then dpoints = 1 else dpoints = 0
objChart.ChartArea(1).Axis(0).SetNumberFormat 1,dpoints
objChart.ChartArea(0).Axis(0).SetNumberFormat 1,dpoints
objChart.ChartArea(0).Axis(2).SetNumberFormat 1,dpoints
objChart.ChartArea(1).Axis(2).SetNumberFormat 1,dpoints


if trim(date2)<>"" then
	call getdataSets(date2, b, utype, "0")
	for i = 1 to 12
		objChart.AddData abs(ArrDataSeriesExpenses(i)), 2,, RGB(255,0,0)
		objChart.AddData abs(ArrDataSeriesUnreportedExp(i)), 3,, RGB(150,0,0)
		objChart.AddData ArrDataSeriesERI(i), 9, left(monthname(i),3), RGB(0,0,255)
		objChart.AddData ArrDataSeriesSubMetered(i), 10,, RGB(51,51,200)
		objChart.AddData ArrDataSeriesUnreportedRev(i), 11,, RGB(103,103,150)
		objChart.AddData ArrDataSeriesMac(i), 12,, RGB(103,150,150)
		objChart.AddData ArrDataSeriesPLP(i), 13,, RGB(150,103,150)
	next
	
	objChart.ChartArea(2).AddChart 7,9,13
	objChart.ChartArea(3).AddChart 7,2,3
	
	objChart.SetColorFromPoint (2)
	objChart.SetColorFromPoint (3)
	objChart.SetColorFromPoint (9)
	objChart.SetColorFromPoint (10)
	objChart.SetColorFromPoint (11)
	objChart.SetColorFromPoint (12)
	objChart.SetColorFromPoint (13)
	objChart.ChartArea(3).Transparent = true
	objChart.ChartArea(2).GridVEnabled = false
	objChart.ChartArea(2).GridHEnabled = false
	objChart.ChartArea(3).GridVEnabled = false
	objChart.ChartArea(3).GridHEnabled = false
	objChart.ChartArea(2).Axis(0).Enabled=true
	objChart.ChartArea(3).Axis(2).Enabled=true
	objChart.ChartArea(2).Axis(2).Enabled=false
	objChart.ChartArea(3).Axis(0).Enabled=false
	objChart.ChartArea(3).Axis(1).Enabled=false
	objChart.ChartArea(2).SetPosition 30,180,450,300
	objChart.ChartArea(3).SetPosition 40,180,460,300
	objChart.AddStaticText "Revenue Profile for "&date2,30,166,RGB(100,100,100),"Arial",8,1
	peak = findpeak()
	objChart.ChartArea(2).Axis(0).Maximum = Peak
	objChart.ChartArea(3).Axis(2).Maximum = Peak
	objChart.ChartArea(2).Axis(2).Maximum = Peak
	objChart.ChartArea(3).Axis(0).Maximum = Peak
  if peak < 6 then dpoints = 1 else dpoints = 0
	objChart.ChartArea(3).Axis(0).SetNumberFormat 1,0
	objChart.ChartArea(2).Axis(0).SetNumberFormat 1,0
	objChart.ChartArea(2).Axis(2).SetNumberFormat 1,0
	objChart.ChartArea(3).Axis(2).SetNumberFormat 1,0
	chartH = 320
end if
objChart.Legend.Enabled = true
objChart.Legend.FontSize = 8
if ArrPrefs(exps) then objChart.Legend.Add "Expenses", RGB(255,0,0)
'if ArrPrefs(aexp) then objChart.Legend.Add "Adjusted Expenses", RGB(150,0,0)
if ArrPrefs(eri) then objChart.Legend.Add "ERI", RGB(0,0,255)
if ArrPrefs(subm) then objChart.Legend.Add "Submetered", RGB(51,51,200)
if ArrPrefs(urae) then objChart.Legend.Add "Expense Adjustments", RGB(150,0,0)
if ArrPrefs(urar) then objChart.Legend.Add "Revenue Adjustments", RGB(103,103,150)
if ArrPrefs(mac) then objChart.Legend.Add "Mac Adjustments", RGB(103,150,150)
if ArrPrefs(plp) then objChart.Legend.Add "PLP", RGB(150,103,150)
objChart.Legend.SetPosition 500,15,600,150 
objChart.SendJpeg 640,chartH,50

function findpeak()
	dim peak
	peak = 1
	for i = 1 to 12
		dim tempExp, tempRev
		tempExp = abs(ArrDataSeriesExpenses(i))+abs(ArrDataSeriesUnreportedExp(i))
		if tempExp>peak then peak = tempExp
		tempRev = ArrDataSeriesSubMetered(i)+ArrDataSeriesUnreportedRev(i)+ArrDataSeriesMac(i)+ArrDataSeriesPLP(i)+ArrDataSeriesERI(i)
		if tempRev>peak then peak = tempRev
	next
	if peak > 50 then peak = 50 - (peak mod 50) + cint(peak) else peak = peak + 1 
	findpeak = peak
end function

%>

