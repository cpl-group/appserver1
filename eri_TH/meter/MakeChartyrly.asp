<%@Language=VBScript%>
<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!--#INCLUDE VIRTUAL="/genergy2/ChartConst.inc" -->
<%
dim hidedemand
hidedemand = request("hidedemand")
dim objChart			'Dundas Chart 2D object
Dim InvoiceGraphKWH()
Dim InvoiceGraphKW()
dim title()

dim ctr					'loop counter

Dim cnn1, rst1, strsql, billyear, billperiod
Dim leaseid, unittype, unitc, unitd, isOUC

leaseid=request.querystring("lid")
billyear=request.querystring("by")
billperiod=request.querystring("bp")
unittype=request("unittype")
isOUC = request("isOUC")
if trim(unittype)="tons" then
	unitc = "Ton Hours"
	unitd = "Tons"
else
	unitc = "KWH"
	unitd = "KW"
end if

set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

if trim(request("genergy2"))<>"" then
  cnn1.Open application("cnnstr_genergy2")
  strsql = "SELECT BillPeriod, SUM(Used) AS KWH, SUM(OnPeak) AS OnPeak, SUM(OffPeak) AS OffPeak, LeaseUtilityId,billyear  FROM tblMetersByPeriod where leaseutilityid=" & leaseid & " and billyear = " & billyear & " and billperiod <= " & billperiod & " GROUP BY BillPeriod, LeaseUtilityId,billyear "
else
  cnn1.Open application("cnnstr_genergy1")
  strsql = "SELECT BillPeriod, SUM(KWHUsed) AS KWH, SUM(OnPeak) AS OnPeak, SUM(OffPeak) AS OffPeak, LeaseUtilityId,billyear  FROM tblMetersByPeriod where leaseutilityid=" & leaseid & " and billyear = " & billyear & " and billperiod <= " & billperiod & " GROUP BY BillPeriod, LeaseUtilityId,billyear "
end if

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
		if trim(isOUC)="true" then title(index) = left(monthname(rst1("billperiod")),1) else title(index) = rst1("billperiod")
		rst1.movenext
	next
	
rst1.close

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
		if trim(isOUC)="true" then title(index) = left(monthname(rst1("billperiod")),1) else title(index) = rst1("billperiod")
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
if trim(hidedemand)="" then objChart.ChartArea(1).AddChart Column_CHART, 3,3

'objChart.ChartArea(0).Axis(0).SetNumberFormat 2,0
objChart.ChartArea(0).Axis(0).TruncatedLabels=True
'objChart.ChartArea(0).Axis(0).Angle=35
objChart.ChartArea(0).Axis(1).FontSize=8
objChart.ChartArea(1).Axis(1).FontSize=8

'objChart.ChartArea(0).GridHEnabled = false
'objChart.ChartArea(0).GridVEnabled = false
'objChart.ChartArea(0).SetShadow 

objChart.AddStaticText unitc,64,142,RGB(0,0,0),"Arial",9,1
objChart.AddStaticText "On Peak",134,142,RGB(128,128,128),"Arial",9,1
objChart.AddStaticText "Off Peak",194,142,RGB(192,192,192),"Arial",9,1
objChart.AddStaticText billyear & " Historical Usage ("&unitc&")",64,0,RGB(0,0,0),"Arial",9,1
if trim(hidedemand)="" then objChart.AddStaticText unitd,340,142,RGB(0,128,128),"Arial",9,1
if trim(hidedemand)="" then objChart.AddStaticText billyear & " Historical Demand ("&unitd&")",340,0,RGB(0,0,0),"Arial",9,1

'setup the colors for each series (we display data elements using their series color,
' we could display them using their individual colors)
objChart.SetSeriesColor 0, RGB(0,0,0) 'first series
objChart.SetSeriesColor 1, RGB(128,128,128)    'second series
objChart.SetSeriesColor 2, RGB(192,192,192)    'second series
objChart.SetSeriesColor 3, RGB(0,128,128)    'second series
objChart.ChartArea(0).Axis(0).SetNumberFormat 1, 0
objChart.ChartArea(0).Axis(1).angle = 0
objChart.ChartArea(1).Axis(1).angle = 0

objChart.SendJpeg 600,175
end if
set objChart = nothing
%>