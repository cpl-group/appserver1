<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
on error resume next
Response.Buffer = true  'enable buffering so that ALL browsers will save image as a JPEG when
						'  a user right-clicks over it and saves it to disk
%>
<!--#INCLUDE VIRTUAL="/INCLUDES/ChartConst.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim objChart			'Dundas Chart 2D object
Dim InvoiceGraphKWH()
Dim InvoiceGraphKW()
dim title()

dim ctr					'loop counter

Dim cnn1, rst1, strsql, billyear, billperiod, building, hidedemand, includepeaks, calcintpeak, billid
Dim leaseid, unittype, unitc, unitd, isOUC, extusage, isGVA,summary
Dim rst2, strsql2
building = request("building")
leaseid=request.querystring("lid")
billyear=request.querystring("by")
billperiod=request.querystring("bp")
extusage = request.querystring("extusg")
billid = request.querystring("billid")
summary = request.QueryString("summary")

if summary = "" then summary = "false"

isOUC = request("isOUC")
isGVA = request("isGVA")
if lcase(trim(request("hidedemand")))="true" then hidedemand = true else hidedemand = false
if lcase(trim(request("includepeaks")))="true" then includepeaks = true else includepeaks = false
if lcase(trim(request("calcintpeak")))="true" then calcintpeak = true else calcintpeak = false

set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst2 = server.CreateObject("ADODB.recordset")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(building)

if extusage <> "" then 
	strsql = "SELECT * FROM (SELECT top 13 m.BillPeriod, SUM(used+usedint+usedoff) AS KWH, SUM(used+usedint) AS OnPeak, SUM(usedoff) AS OffPeak, m.LeaseUtilityId, m.billyear, l.utility FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l, billyrperiod b WHERE bbp.id=m.bill_id and bbp.reject=0 and b.ypid=m.ypid and m.leaseutilityid=l.leaseutilityid and m.leaseutilityid=" & leaseid & " and b.datestart<=(SELECT bb.datestart FROM billyrperiod bb WHERE bb.billyear=" & billyear & " and bb.billperiod=" & billperiod & " and bb.BldgNum=b.BldgNum and bb.utility=b.utility) GROUP BY m.BillPeriod, m.LeaseUtilityId, m.billyear, l.utility, b.datestart ORDER BY b.datestart desc) h ORDER BY billyear, billperiod"
	if summary = "true" then
	strsql = "SELECT * FROM (SELECT top 13 m.BillPeriod, SUM(used+usedint+usedoff) AS KWH, SUM(used+usedint) AS OnPeak, SUM(usedoff) AS OffPeak, m.billyear, l.utility FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l, billyrperiod b WHERE bbp.id=m.bill_id and bbp.reject=0 and b.ypid=m.ypid and m.leaseutilityid=l.leaseutilityid and b.bldgnum='" & building & "' and b.datestart<=(SELECT bb.datestart FROM billyrperiod bb WHERE bb.billyear=" & billyear & " and bb.billperiod=" & billperiod & " and bb.BldgNum=b.BldgNum and bb.utility=b.utility) GROUP BY m.BillPeriod, m.billyear, l.utility, b.datestart ORDER BY b.datestart desc) h ORDER BY billyear, billperiod"
	end if	
else
	strsql = "SELECT * FROM (SELECT top 13 m.BillPeriod, SUM(used) AS KWH, SUM(OnPeak) AS OnPeak, SUM(Offpeak) AS OffPeak, m.LeaseUtilityId, m.billyear, l.utility FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l, billyrperiod b WHERE b.ypid=m.ypid and bbp.id=m.bill_id and bbp.reject=0 and m.leaseutilityid=l.leaseutilityid and m.leaseutilityid=" & leaseid & " and b.datestart<=(SELECT bb.datestart FROM billyrperiod bb WHERE bb.billyear=" & billyear & " and bb.billperiod=" & billperiod & " and bb.BldgNum=b.BldgNum and bb.utility=b.utility) GROUP BY m.BillPeriod, m.LeaseUtilityId, m.billyear, l.utility, b.datestart ORDER BY b.datestart desc) h ORDER BY billyear, billperiod"
	'KCheng added 6/3/2009 
	strsql2 = "Select * from tblleasespecificmeasure where LeaseutilityId="&leaseid
	if summary = "true" then
	strsql = "SELECT * FROM (SELECT top 13 m.BillPeriod, SUM(used) AS KWH, SUM(OnPeak) AS OnPeak, SUM(Offpeak) AS OffPeak, m.billyear, l.utility FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l, billyrperiod b WHERE b.ypid=m.ypid and bbp.id=m.bill_id and bbp.reject=0 and m.leaseutilityid=l.leaseutilityid and b.bldgnum='" & building & "' and b.datestart<=(SELECT bb.datestart FROM billyrperiod bb WHERE bb.billyear=" & billyear & " and bb.billperiod=" & billperiod & " and bb.BldgNum=b.BldgNum and bb.utility=b.utility) GROUP BY m.BillPeriod, m.billyear, l.utility, b.datestart ORDER BY b.datestart desc) h ORDER BY billyear, billperiod"
	end if	

end if



rst1.Open strsql, cnn1, adOpenStatic
'response.write strsql
'response.end
if not rst1.eof then unittype=cint(rst1("utility"))

dim usagedivisor
usagedivisor = 1.0
if trim(unittype)="CCF" or trim(unittype)="3" or trim(unittype)="4" or trim(unittype)="10" then
	unitc = "CCF Used"
	unitd = ""
	usagedivisor = 100.0
elseif trim(unittype)="tons" or trim(unittype)=6 then
	unitc = "Ton Hours"
	unitd = "Tons"
elseif unittype=21 then
    unitc = "Ton Hours"
	unitd = "Tons"
elseif trim(unittype)="1" then
	unitc = "M#"
	unitd = "M#/Hr"
elseif trim(unittype) = "4" then 
	unitc = "CF Used"
elseif trim(unittype)="18" or trim(unittype)="19" or trim(unittype)="20" then  '4/28/2009 N.Ambo added this line for new chiller utilities
	unitc = "MMBTU"
	unitd = ""
else
	unitc = "Kilowatt Hours"
	unitd = "Kilowatt of Demand"
end if

if ((unittype = 6 OR unittype = 21) and building = "1992B") then unitc = "KBTU" end if 'rsm
if ((unittype = 6 OR unittype = 21) and building = "1992B") then unitd = "KBTU" end if 'rsm
 
'KCheng added 6/3/2009 
if (unittype = 6 OR unittype = 21) then
    rst2.open strsql2, cnn1
     if (NOT rst2.EOF) then
        if ( rst2("ConsumptionMeasure") <> "default") then
            unitc = rst2("ConsumptionMeasure")
        end if
        if (rst2("DemandMeasure") <> "default") then
            unitd = rst2("DemandMeasure")
        end if
    end if
    
    rst2.close
end if


if not rst1.eof then 
Dim numRecords
numRecords = rst1.RecordCount
ReDim InvoiceGraphKWH(numRecords, 3)
ReDim title(numRecords)

Dim index

	
	for index=1 to NumRecords
		InvoiceGraphKWH(index, 1) = cdbl(rst1("KWH")) / usagedivisor
		InvoiceGraphKWH(index, 2) = rst1("OnPeak")
		InvoiceGraphKWH(index, 3) = rst1("OffPeak")
		if trim(isOUC)="true" then 
			title(index) = left(monthname(rst1("billperiod")),1) 
		elseif trim(isGVA)="true" then
			title(index) = rst1("billperiod") & "/" & rst1("billyear")
		else
			title(index) = rst1("billyear") & "/" & rst1("billperiod")
		end if
		rst1.movenext
	next
	
rst1.close

objChart.ChartArea(0).SetPosition 65, 15, 300,100
objChart.ChartArea(1).SetPosition 340, 15, 575,100

'Step 2: Add ALL data to be used by our chart
for ctr=1 to numRecords
	objChart.AddData InvoiceGraphKWH(ctr, 1), 0, title(ctr) 'Add data to Data Series 0 
	objChart.AddData InvoiceGraphKWH(ctr, 2), 1 	'Add data to Data Series 0 
	objChart.AddData InvoiceGraphKWH(ctr, 3), 2	'Add data to Data Series 0 
next

if 1=0 then
	strsql = "SELECT * FROM (SELECT top 13 bbp.BillPeriod, SUM(total_mlbs) AS KW, bbp.LeaseUtilityId, b.billyear, l.utility FROM custom_bchbill_water_steam m, tblbillbyperiod bbp, tblleasesutilityprices l, billyrperiod b WHERE bbp.id=m.bill_id and bbp.reject=0 and b.ypid=m.ypid and m.leaseutilityid=l.leaseutilityid and m.leaseutilityid="&leaseid&" and b.datestart<=(SELECT bb.datestart FROM billyrperiod bb WHERE bb.billyear="&billyear&" and bb.billperiod="&billperiod&" and bb.BldgNum=b.BldgNum and bb.utility=b.utility) GROUP BY bbp.BillPeriod, bbp.LeaseUtilityId, b.billyear, l.utility, b.datestart ORDER BY b.datestart desc) h ORDER BY billyear, billperiod"
else
	strsql = "SELECT * FROM (SELECT top 13 m.BillPeriod, SUM(Demand_P) AS KW, m.LeaseUtilityId, m.billyear FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l, billyrperiod b where m.leaseutilityid=l.leaseutilityid and m.ypid=b.ypid and b.datestart<=(SELECT bb.datestart FROM billyrperiod bb WHERE bb.billyear=" & billyear & " and bb.billperiod=" & billperiod & " and bb.BldgNum=b.BldgNum and bb.utility=b.utility) and bbp.id=m.bill_id and bbp.reject=0 and m.leaseutilityid=" & leaseid & " GROUP BY m.BillPeriod, m.LeaseUtilityId, m.billyear, b.datestart ORDER BY b.datestart desc) h ORDER BY billyear, billperiod"
end if

if summary = "true" then
   strsql = "SELECT * FROM (SELECT top 13 m.BillPeriod, SUM(Demand_P) AS KW, m.billyear FROM tblMetersByPeriod m, tblbillbyperiod bbp, tblleasesutilityprices l, billyrperiod b where m.leaseutilityid=l.leaseutilityid and m.ypid=b.ypid and b.datestart<=(SELECT bb.datestart FROM billyrperiod bb WHERE bb.billyear=" & billyear & " and bb.billperiod=" & billperiod & " and bb.BldgNum=b.BldgNum and bb.utility=b.utility) and bbp.id=m.bill_id and bbp.reject=0 and b.bldgnum='" & building & "' GROUP BY m.BillPeriod, m.billyear, b.datestart ORDER BY b.datestart desc) h ORDER BY billyear, billperiod"
end if
'response.write strsql
'response.end
rst1.Open strsql, cnn1, adOpenStatic


if not rst1.eof then 
numRecords = rst1.RecordCount
ReDim InvoiceGraphKW(numRecords)
ReDim title(numRecords)
	
	for index=1 to NumRecords
		InvoiceGraphKW(index) = rst1("kw")
		if trim(isOUC)="true" then
			title(index) = left(monthname(rst1("billperiod")),1)
		elseif trim(isGVA)="true" then
			title(index) = rst1("billperiod") & "/" & rst1("billyear")
		else
			title(index) = rst1("billyear") & "/" & rst1("billperiod")
		end if
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
if includepeaks and not(calcintpeak) and unittype<>3 and unittype <> 10 then 'don't show the onpeak and offpeak
  objChart.ChartArea(0).AddChart Column_CHART, 0,2
  objChart.AddStaticText "On Peak",149,142,RGB(128,128,128),"Arial",9,1
  objChart.AddStaticText "Off Peak",209,142,RGB(192,192,192),"Arial",9,1
else
  objChart.ChartArea(0).AddChart Column_CHART, 0,0 
end if
if not hidedemand then objChart.ChartArea(1).AddChart Column_CHART, 3,3

'objChart.ChartArea(0).Axis(0).SetNumberFormat 2,0
objChart.ChartArea(0).Axis(0).TruncatedLabels=True
'objChart.ChartArea(0).Axis(0).Angle=35
objChart.ChartArea(0).Axis(1).FontSize=8
objChart.ChartArea(1).Axis(1).FontSize=8
objChart.ChartArea(1).Axis(0).SetNumberFormat 1,1

'objChart.ChartArea(0).GridHEnabled = false
'objChart.ChartArea(0).GridVEnabled = false
'objChart.ChartArea(0).SetShadow 

objChart.AddStaticText billyear & " Historical Usage ("&unitc&")",64,0,RGB(0,0,0),"Arial",9,1
if not hidedemand then objChart.AddStaticText billyear & " Historical Demand ("&unitd&")",340,0,RGB(0,0,0),"Arial",9,1
if calcintpeak then
  unitc = "Total Usage"
  unitd = "On"
end if
objChart.AddStaticText unitc,64,142,RGB(0,0,0),"Arial",9,1
if not hidedemand then objChart.AddStaticText unitd,340,142,RGB(0,128,128),"Arial",9,1

'setup the colors for each series (we display data elements using their series color,
' we could display them using their individual colors)
objChart.SetSeriesColor 0, RGB(0,0,0) 'first series
objChart.SetSeriesColor 1, RGB(128,128,128)    'second series
objChart.SetSeriesColor 2, RGB(192,192,192)    'second series
objChart.SetSeriesColor 3, RGB(0,128,128)    'second series
objChart.ChartArea(0).Axis(0).SetNumberFormat 1, 0

'error checking code
if err then
	dim errChart
	set errChart = Server.CreateObject("Dundas.ChartServer2D.2")
	errChart.AddStaticText err.Description,0,0,RGB(0,0,0),"Arial",8,1
	errChart.SendJpeg 600,175
	response.end
end if

objChart.SendJpeg 600,175
end if
set objChart = nothing
%>