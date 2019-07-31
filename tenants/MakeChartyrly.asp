<%@Language=VBScript%>
<%Option Explicit

Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!-- #INCLUDE FILE ="ChartConst.inc" -->
<!--#include file="adovbs.inc"-->

<%
dim objChart			'Dundas Chart 2D object
Dim InvoiceGraphKWH()
Dim InvoiceGraphKW()
dim title()

dim ctr					'loop counter

Dim cnn1
Dim rst1
Dim strsql
Dim billyear, billperiod
Dim leaseid

leaseid=request.querystring("lid")
billyear=request.querystring("by")
billperiod=request.querystring("bp")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"


strsql = "SELECT BillPeriod, SUM(KWHUsed) AS KWH, SUM(OnPeak) AS OnPeak, SUM(OffPeak) AS OffPeak, LeaseUtilityId,billyear  FROM tblMetersByPeriod where leaseutilityid=" & leaseid & " and billyear = " & billyear & " and billperiod <= " & billperiod & " GROUP BY BillPeriod, LeaseUtilityId,billyear "

rst1.Open strsql, cnn1, adOpenStatic


if not rst1.eof then 
Dim numRecords
numRecords = rst1.RecordCount
ReDim InvoiceGraphKWH(numRecords, 3)
ReDim title(numRecords)

Dim index
	
	for index=1 to NumRecords
		InvoiceGraphKWH(index, 1) = rst1("KWH")
		InvoiceGraphKWH(index, 2) = rst1("OnPeak")
		InvoiceGraphKWH(index, 3) = rst1("OffPeak")
		title(index) = billyear & "/" & rst1("billperiod")		
		rst1.movenext
	next
	
rst1.close
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

objChart.ChartArea(0).SetPosition 65, 15, 300,100
objChart.ChartArea(1).SetPosition 340, 15, 575,100

'Step 2: Add ALL data to be sued by our chart
for ctr=1 to numRecords
	objChart.AddData InvoiceGraphKWH(ctr, 1), 0, title(ctr) 'Add data to Data Series 0 
	objChart.AddData InvoiceGraphKWH(ctr, 2), 1 	'Add data to Data Series 0 
	objChart.AddData InvoiceGraphKWH(ctr, 3), 2	'Add data to Data Series 0 
next

strsql = "SELECT BillPeriod, SUM(Demand_P) AS KW, LeaseUtilityId,billyear  FROM tblMetersByPeriod where leaseutilityid=" & leaseid & " and billyear = " & billyear & " and billperiod <= " & billperiod & " GROUP BY BillPeriod, LeaseUtilityId,billyear "

rst1.Open strsql, cnn1, adOpenStatic


if not rst1.eof then 
numRecords = rst1.RecordCount
ReDim InvoiceGraphKW(numRecords)
ReDim title(numRecords)
	
	for index=1 to NumRecords
		InvoiceGraphKW(index) = rst1("kw")
		title(index) = billyear & "/" & rst1("billperiod")		

		rst1.movenext
	next
	
rst1.close

'Step 2: Add ALL data to be sued by our chart
for ctr=1 to numRecords
	objChart.AddData InvoiceGraphKW(ctr), 3, title(ctr)	'Add data to Data Series 0 
next

end if

'Step 3: Use data in Data Series 0 and 1 to make a Column chart, then
'		 add this chart to ChartArea 0. The constant "STACKED_COLUMN_CHART" has been
'		 defined in ChartConst.inc file.

objChart.ChartArea(0).AddChart Column_CHART, 0,2
objChart.ChartArea(1).AddChart Column_CHART, 3,3

'objChart.ChartArea(0).Axis(0).SetNumberFormat 2,0
objChart.ChartArea(0).Axis(0).TruncatedLabels=True
'objChart.ChartArea(0).Axis(0).Angle=35
objChart.ChartArea(0).Axis(1).FontSize=8
objChart.ChartArea(1).Axis(1).FontSize=8

'objChart.ChartArea(0).GridHEnabled = false
'objChart.ChartArea(0).GridVEnabled = false
'objChart.ChartArea(0).SetShadow 

objChart.AddStaticText "KWH",64,127,RGB(0,0,0),"Arial",9,1
objChart.AddStaticText "On Peak",104,127,RGB(128,128,128),"Arial",9,1
objChart.AddStaticText "Off Peak",164,127,RGB(192,192,192),"Arial",9,1
objChart.AddStaticText "KW",340,127,RGB(0,128,128),"Arial",9,1
objChart.AddStaticText "Historical Energy Usage (KWH)",64,0,RGB(0,0,0),"Arial",9,1
objChart.AddStaticText "Historical Demand (KW)",340,0,RGB(0,0,0),"Arial",9,1

'setup the colors for each series (we display data elements using their series color,
' we could display them using their individual colors)
objChart.SetSeriesColor 0, RGB(0,0,0) 'first series
objChart.SetSeriesColor 1, RGB(128,128,128)    'second series
objChart.SetSeriesColor 2, RGB(192,192,192)    'second series
objChart.SetSeriesColor 3, RGB(0,128,128)    'second series

objChart.SendJpeg 600,175
end if
set objChart = nothing
%>