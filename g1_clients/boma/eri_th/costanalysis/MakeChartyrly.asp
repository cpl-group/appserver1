<%@Language=VBScript%>
<%Option Explicit

Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!-- #INCLUDE FILE ="ChartConst.inc" -->
<!--#include file="adovbs.inc"-->

<%
dim objChart			'Dundas Chart 2D object
Dim ArrRevenue(12)
Dim ArrExpenses(12)
Dim Bldgname

dim ctr					'loop counter

Dim Title(12) 'Pair title
Title(1) = "Jan"
Title(2) = "Feb"
Title(3) = "Mar"
Title(4) = "Apr"
Title(5) = "May"
Title(6) = "Jun"
Title(7) = "Jul"
Title(8) = "Aug"
Title(9) = "Sep"
Title(10) = "Oct"
Title(11) = "Nov"
Title(12) = "Dec"

'ArrDataSeries0= Array(12, 15, 20, 9, 13)
'ArrDataSeries1= Array(13, 10, 15, 10, 15)

Dim cnn1
Dim rst1
Dim strsql
Dim bldg
Dim year

year=request.querystring("year")
bldg=request.querystring("bldgnum")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"


strsql = "SELECT eri_data.dbo.eri_total.eri_amt AS ERI_rev, UtilityBill.TotalBillAmt AS Expenses, SUM(tblBillByPeriod.TotalAmt) AS SubMetered, BillYrPeriod.BillPeriod, tblRPentries_sum.Expr1 AS UnReportedAmts FROM UtilityBill FULL OUTER JOIN BillYrPeriod FULL OUTER JOIN tblRPentries_sum ON  BillYrPeriod.BillPeriod = tblRPentries_sum.period AND BillYrPeriod.BldgNum = tblRPentries_sum.bldgnum AND tblRPentries_sum.userid = 'cotto' AND BillYrPeriod.BillYear = tblRPentries_sum.year ON UtilityBill.ypId = BillYrPeriod.ypId FULL OUTER JOIN eri_data.dbo.eri_total ON BillYrPeriod.BldgNum = eri_data.dbo.eri_total.bldg_no AND BillYrPeriod.BillPeriod = eri_data.dbo.eri_total.BillPeriod AND BillYrPeriod.BillYear = eri_data.dbo.eri_total.BillYear FULL OUTER JOIN tblBillByPeriod ON  UtilityBill.ypId = tblBillByPeriod.ypId WHERE (BillYrPeriod.BldgNum = '"& bldg &"') AND (BillYrPeriod.BillYear = '"& year &"') GROUP BY eri_data.dbo.eri_total.eri_amt, UtilityBill.TotalBillAmt,  BillYrPeriod.BillPeriod, tblRPentries_sum.Expr1 ORDER BY BillYrPeriod.BillPeriod"


rst1.Open strsql, cnn1, adOpenStatic


if not rst1.eof then 
Dim numRecords
numRecords = rst1.RecordCount
Dim unreprev, unrepexp, submetered, eri_rev, expenses
Dim index
	
	for index=0 to NumRecords-1
		unrepexp =0
		unreprev =0

		if IsNull(rst1("Submetered")) then 
		submetered = 0
		else 
		submetered = clng(rst1("Submetered"))
		end if

		if not IsNull(rst1("UnreportedAmts")) then 
			if rst1("UnreportedAmts") < 0 then 
				unrepexp =(ccur(rst1("UnreportedAmts"))) * -1
			else
				unreprev =(ccur(rst1("UnreportedAmts")))
			end if 
		end if
		eri_rev = (rst1("eri_rev"))
		Expenses = (rst1("Expenses"))
		
		ArrExpenses(index) = Expenses + UnRepExp
		ArrRevenue(index) = eri_rev + submetered + unreprev
		
		rst1.movenext
	next
	
rst1.close

set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

objChart.ChartArea(0).SetPosition 65, 60, 400,200

'Step 2: Add ALL data to be sued by our chart
for ctr=1 to numRecords
	objChart.AddData ArrExpenses(ctr-1), 0	'Add data to Data Series 0 
	objChart.AddData ArrRevenue(ctr-1),  1	, title(ctr)	'and assign a label to this data
next



'Step 3: Use data in Data Series 0 and 1 to make a Column chart, then
'		 add this chart to ChartArea 0. The constant "STACKED_COLUMN_CHART" has been
'		 defined in ChartConst.inc file.

objChart.ChartArea(0).AddChart 10, 1,1,,2
objChart.ChartArea(0).AddChart Line_CHART, 0,0
objChart.ShowPieLabels "Arial",8
objChart.ChartArea(0).Axis(0).SetNumberFormat 2,0
objChart.ChartArea(0).Axis(0).TruncatedLabels=True
objChart.ChartArea(0).Axis(0).Angle=35
objChart.ChartArea(0).Axis(1).Angle=45
objChart.ChartArea(0).GridHEnabled = false
objChart.ChartArea(0).GridVEnabled = false
objChart.ChartArea(0).SetShadow 

objChart.AddStaticText "Amount $$$", 1,175,RGB(0,0,0),"Arial",12,0,0,90

objChart.AddStaticText "Revenue vs. Expenses for " & year,65,40,RGB(0,0,0),"Arial",9,0,0
'--------------------------------------------
' setup the legend
'--------------------------------------------
objChart.Legend.Enabled = true  'enable the legend (it is disabled by default)
objChart.Legend.FontSize = 8
'setup the labels for each series, these names will then appear in the legend
'Note: if using a legend with a pie chart the labels are retrieved from the charts
'data points (see the Chart Object's AddData member for more details) instead of
' from data series.
'Note: you can also specify custom legend entries by calling the Add method of
' the legend object.

objChart.Legend.Add "Expenses", RGB(255,0,0)
objChart.Legend.Add "Revenue",  RGB(40,120,255)

objChart.Legend.SetPosition 420,60,500,120

'optional legend settings
objChart.Legend.BorderColor = RGB(110,0,0)
objChart.Legend.BackgroundColor = RGB(230,230,230)
objChart.Legend.FontColor = RGB(0,0,110)
objChart.Legend.Transparent = false 'set to false, so that the background color
                                    ' can be seen
'--------------------------------------------
' finished setting up the legend
'--------------------------------------------

'setup the colors for each series (we display data elements using their series color,
' we could display them using their individual colors)
objChart.SetSeriesColor 0, RGB(255,0,0) 'first series
objChart.SetSeriesColor 1, RGB(40,120,255)    'second series

objChart.SendJpeg 500,250
end if
set objChart = nothing
%>