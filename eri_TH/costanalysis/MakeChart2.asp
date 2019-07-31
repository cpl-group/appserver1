<%@Language=VBScript%>
<%Option Explicit

Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!-- #INCLUDE FILE ="ChartConst.inc" -->
<!--#include file="adovbs.inc"-->

<%
dim objChart			'Dundas Chart 2D object
dim ArrDataSeriesERI(12)		'Array of first set of data 
dim ArrDataSeriesSubMetered(12)
dim ArrDataSeriesUnreportedExp(12)
dim ArrDataSeriesUnreportedRev(12)
Dim ArrPieRevenue(3)
Dim ArrPieExpenses(2)
Dim Bldgname
dim ArrDataSeriesExpenses(12)		'Array of second set of data

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
'response.write strsql
'response.end

rst1.Open strsql, cnn1, adOpenStatic


if not rst1.eof then 
Dim numRecords
numRecords = rst1.RecordCount

Dim index
	
	for index=0 to NumRecords-1
		if isnull(rst1("eri_rev")) then 
			ArrDataSeriesERI(index)=0
		else
			ArrDataSeriesERI(index)=rst1("eri_rev")/1000
		end if
		if isnull(rst1("Expenses")) then 
			ArrDataSeriesExpenses(index)=0
		else
			ArrDataSeriesExpenses(index)=rst1("Expenses")/1000
		end if
		if IsNull(rst1("Submetered")) then 
			ArrDataSeriesSubmetered(index)=0		
		else
			ArrDataSeriesSubmetered(index)=clng(rst1("Submetered"))/1000
		end if

		if IsNull(rst1("UnreportedAmts")) then
				ArrDataSeriesUnreportedExp(index)=0
				ArrDataSeriesUnreportedRev(index)=0
		else
			if rst1("UnreportedAmts") < 0 then 
				ArrDataSeriesUnreportedExp(index)=(clng(rst1("UnreportedAmts"))/1000) * -1
		ArrPieExpenses(2)=ArrPieExpenses(1) + ((clng(ArrDataSeriesUnreportedExp(index))) * -1)
				ArrDataSeriesUnreportedRev(index)=0
			else
				ArrDataSeriesUnreportedRev(index)=(clng(rst1("UnreportedAmts"))/1000)
		ArrPieRevenue(3)=ArrPieRevenue(3) + (clng(ArrDataSeriesUnreportedRev(index)))
				ArrDataSeriesUnreportedExp(index) = 0
			end if 
		end if
		
		ArrPieRevenue(1)=ArrPieRevenue(1) + (ArrDataSeriesERI(index))  
		ArrPieRevenue(2)=ArrPieRevenue(2) + (clng(ArrDataSeriesSubmetered(index)))
		ArrPieExpenses(1)=ArrPieExpenses(1) + (ArrDataSeriesExpenses(index))
		rst1.movenext
	next
rst1.close

'Step 1: Create a Dundas Chart 2D object
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
'Mark the top left corner of the JPEG to (0,0) and the 
'bottom right corner to (1000, 1000) regardless the JPEG size.
objChart.Rectangle3DEffect()

'All arguments will be in relative coordinates
'The ChartArea occupies the whole JPEG.
objChart.ChartArea(0).SetPosition 30, 50, 500,200
objChart.ChartArea(1).SetPosition 125,270,225,330
objChart.ChartArea(2).SetPosition 305,270,405,330


'Step 2: Add ALL data to be sued by our chart
for ctr=1 to numRecords
	objChart.AddData ArrDataSeriesExpenses(ctr-1), 0, title(ctr)	'Add data to Data Series 0 
	objChart.AddData ArrDataSeriesUnreportedExp(ctr-1),  1		'and assign a label to this data
	objChart.AddData ArrDataSeriesERI(ctr-1),	2			'Add data to Data Series 1
	objChart.AddData ArrDataSeriesSubMetered(ctr-1),3 
	objChart.AddData ArrDataSeriesUnreportedRev(ctr-1),  4
next
		objChart.AddData ArrPieExpenses(1),  5, "Expenses",RGB(255,0,0)
		
		objChart.AddData ArrPieExpenses(2),  5, "Adj Exp" ,RGB(255,96,0)
	if ArrPieRevenue(1) > 0 then 
		objChart.AddData ArrPieRevenue(1),  6, "ERI Revenue",RGB(0,204,254)
	end if
	if ArrPieRevenue(2) > 0 then 
		objChart.AddData ArrPieRevenue(2),  6, "Submeter Revenue",RGB(0,126,255)
	end if
	if ArrPieRevenue(3) > 0 then 
		objChart.AddData ArrPieRevenue(3),  6, "Adj Rev", RGB(0,64,128)
	end if
'Step 3: Use data in Data Series 0 and 1 to make a Column chart, then
'		 add this chart to ChartArea 0. The constant "STACKED_COLUMN_CHART" has been
'		 defined in ChartConst.inc file.
objChart.ChartArea(0).AddChart 10, 0,1,,2
objChart.ChartArea(0).AddChart STACKED_COLUMN_CHART, 2,4

'Publish Pie Charts for Expenses and Revenue
if ArrPieExpenses(1) <> 0 and ArrPieExpenses(2) <> 0   then 
	objChart.ChartArea(1).AddChart PIE_CHART, 5,5
	if ArrPieExpenses(2) <> 0 then 
		objChart.SetExploded 5, 1
	end if

end if
if (ArrPieRevenue(1) > 0 and ArrPieRevenue(2) > 0) or  (ArrPieRevenue(1) > 0 and ArrPieRevenue(3) > 0) or (ArrPieRevenue(2) > 0 and ArrPieRevenue(3) > 0)    then
	objChart.ChartArea(2).AddChart PIE_CHART, 6,6
	if (ArrPieRevenue(3) <> 0 and (ArrPieRevenue(2) <> 0 and ArrPieRevenue(1) <> 0)) then 
		objChart.SetExploded 6, 2
	else
		if (ArrPieRevenue(3) <> 0 and (ArrPieRevenue(2) <> 0 or ArrPieRevenue(1) <> 0)) then 
			objChart.SetExploded 6, 1
		end if
	end if
	
end if




objChart.ShowPieLabels "Arial",8


'setup the colors for each series (we display data elements using their series color,
' we could display them using their individual colors)
objChart.SetSeriesColor 0, RGB(255,0,0) 'first series
objChart.SetSeriesColor 1, RGB(255,96,0)    'second series
objChart.SetSeriesColor 2, RGB(0,204,254)    'third series
objChart.SetSeriesColor 3, RGB(0,126,255)'fourth
objChart.SetSeriesColor 4, RGB(0,64,128)'fourth
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

if ArrPieExpenses(1) <> 0 then 
objChart.Legend.Add "Expenses", RGB(255,0,0)
end if
if ArrPieExpenses(2) <> 0 then 
objChart.Legend.Add "Adjusted Expenses",RGB(255,96,0)
end if
if ArrPieRevenue(1) <> 0 then 
	objChart.Legend.Add "ERI Revenue",  RGB(0,204,254)
end if
if ArrPieRevenue(2) <> 0 then 
	objChart.Legend.Add "Submeter Revenue", RGB(0,126,255)
end if
if ArrPieRevenue(3) <> 0 then 
objChart.Legend.Add "Adjusted Revenue", RGB(0,64,128)
end if

'set the legend position.  Note that if you specify relative coordinates to be used via
' the SetCoordinates method then you would not have to reposition the legend if the size
'  of the output JPEG is changed (see the documentation for more details concerning the
'  SetRelativeCoordinates method)
'set the position of the legend
objChart.Legend.SetPosition 520,50,635,120

'optional legend settings
objChart.Legend.BorderColor = RGB(110,0,0)
objChart.Legend.BackgroundColor = RGB(230,230,230)
objChart.Legend.FontColor = RGB(0,0,110)
objChart.Legend.Transparent = false 'set to false, so that the background color
                                    ' can be seen
'--------------------------------------------
' finished setting up the legend
'--------------------------------------------
'objChart.AddStaticText "Revenue Profile for " & bldgname,0,0,RGB(0,0,0),"Arial",14

objChart.SendJpeg 640,380
end if
set objChart = nothing
%>