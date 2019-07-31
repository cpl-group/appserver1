<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim dateone, b
dateone = request.querystring("dateone")
b = request.querystring("b")

dim cnn1, rst1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"


strsql = "SELECT isnull(eri_data.dbo.eri_total.eri_amt,0) AS ERI_rev, isnull(UtilityBill.TotalBillAmt,0) AS Expenses, SUM(isnull(tblBillByPeriod.TotalAmt,0)) AS SubMetered, BillYrPeriod.BillPeriod, isnull(tblRPentries_sum.Expr1,0) AS UnReportedAmts FROM UtilityBill FULL OUTER JOIN BillYrPeriod FULL OUTER JOIN tblRPentries_sum ON  BillYrPeriod.BillPeriod = tblRPentries_sum.period AND BillYrPeriod.BldgNum = tblRPentries_sum.bldgnum AND tblRPentries_sum.userid = 'cotto' AND BillYrPeriod.BillYear = tblRPentries_sum.year ON UtilityBill.ypId = BillYrPeriod.ypId FULL OUTER JOIN eri_data.dbo.eri_total ON BillYrPeriod.BldgNum = eri_data.dbo.eri_total.bldg_no AND BillYrPeriod.BillPeriod = eri_data.dbo.eri_total.BillPeriod AND BillYrPeriod.BillYear = eri_data.dbo.eri_total.BillYear FULL OUTER JOIN tblBillByPeriod ON  UtilityBill.ypId = tblBillByPeriod.ypId WHERE (BillYrPeriod.BldgNum = '"& b &"') AND (BillYrPeriod.BillYear = '"& dateone &"') GROUP BY eri_data.dbo.eri_total.eri_amt, UtilityBill.TotalBillAmt,  BillYrPeriod.BillPeriod, tblRPentries_sum.Expr1 ORDER BY BillYrPeriod.BillPeriod"
response.write strsql
response.end

rst1.Open strsql, cnn1, adOpenStatic

dim ArrDataSeriesERI(12)
dim ArrDataSeriesSubMetered(12)
dim ArrDataSeriesUnreportedExp(12)
dim ArrDataSeriesUnreportedRev(12)
Dim ArrPieRevenue(3)
Dim ArrPieExpenses(2)
Dim recordnum
recordnum=0
response.write rst1.eof
do until rst1.eof
	recordnum=recordnum+1
	ArrDataSeriesERI(recordnum)=rst1("eri_rev")/1000
	ArrDataSeriesExpenses(recordnum)=rst1("Expenses")/1000
	ArrDataSeriesSubmetered(recordnum)=clng(rst1("Submetered"))/1000
	if rst1("UnreportedAmts") < 0 then 
		ArrDataSeriesUnreportedExp(recordnum)=(clng(rst1("UnreportedAmts"))/1000) * -1
		ArrPieExpenses(2)=ArrPieExpenses(1) + ((clng(ArrDataSeriesUnreportedExp(index))) * -1)
		ArrDataSeriesUnreportedRev(recordnum)=0
	else
		ArrDataSeriesUnreportedRev(recordnum)=(clng(rst1("UnreportedAmts"))/1000)
		ArrPieRevenue(3)=ArrPieRevenue(3) + (clng(ArrDataSeriesUnreportedRev(index)))
		ArrDataSeriesUnreportedExp(recordnum) = 0
	end if 
	ArrPieRevenue(1)=ArrPieRevenue(1) + (ArrDataSeriesERI(recordnum))  
	ArrPieRevenue(2)=ArrPieRevenue(2) + (clng(ArrDataSeriesSubmetered(recordnum)))
	ArrPieExpenses(1)=ArrPieExpenses(1) + (ArrDataSeriesExpenses(recordnum))
	rst1.movenext
loop
rst1.close


'make chart
dim objChart, i
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
response.write recordnum
for i = 1 to recordnum
	response.write ArrDataSeriesERI(i)
	'objChart.AddData (Data As Double, Series As Long, [DataLabel As String=""], [Color As Long = &HFFFFF])
next

%>

hey