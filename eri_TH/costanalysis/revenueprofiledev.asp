<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="adovbs.inc"-->
<html>
<head>
<%

		
Dim bldg
bldg=Request.QueryString("bldgnum")
Dim year
year=Request.QueryString("year")
Dim userid
userid=Session("loginemail")
Dim Bldgname
if bldg<>"" then 		
%>

<title>Revenue Profile</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function viewrevprof(bldg, year, userid) {
	var temp
		temp="revenueprofiledev.asp?bldgnum=" + bldg +"&year=" + year +"&userid="+userid
		document.frames.admin.location=temp
} 
function loadypidlist(bldg,pid) {
	var temp = "revbldglistdev.asp?bldg=" + bldg + "&pid="+pid
	document.location = temp
}
function bldglist(pid){
document.location="revbldglistdev.asp?pid=" + pid
}
function unreported(bldg, year, userid){
	var temp = "unreported.asp?bldg=" + bldg + "&year="+year+"&userid="+userid
 	 window.open(temp,"", "scrollbars=no, width=500, height=600, resizeable, status")

}
</script>
<style>
 .pagebreak {page-break-before: always}
</style>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<%
Dim cnn1
Dim rst1
Dim strsql

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
strsql="Select strt from buildings where bldgnum = '" & Bldg &"'"

rst1.Open strsql, cnn1, adOpenStatic
if not rst1.EOF then 
	bldgname=rst1("strt")
end if
rst1.close
%>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000">
<div align="center"> 
  <table border="1" cellspacing="0" cellpadding="0" bordercolor="#333333" align="center" width="100%">
    <tr> 
      <td height="11" valign="top" bgcolor="#000000"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF"><%=bldgname%> 
        for <%=year%></font></b></font> </td>
    </tr>
    <tr> 
      <td height="100%"> 
        <div align="center"><img src="<%="makechart.asp?bldgNUM=" & bldg & "&year=" & year %>"></div>
      </td>
    </tr>
    <tr>
      <td height="2" bgcolor="#000000"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>MONTHLY 
        DETAILS</b></font></td>
    </tr>
    <tr> 
      <td height="474" valign="top"> 
          <%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"


strsql = "SELECT eri_data.dbo.eri_total.eri_amt AS ERI_rev, UtilityBill.TotalBillAmt AS Expenses, SUM(tblBillByPeriod.TotalAmt) AS SubMetered, BillYrPeriod.BillPeriod, tblRPentries_sum.Expr1 AS UnReportedAmts FROM UtilityBill FULL OUTER JOIN BillYrPeriod FULL OUTER JOIN tblRPentries_sum ON  BillYrPeriod.BillPeriod = tblRPentries_sum.period AND BillYrPeriod.BldgNum = tblRPentries_sum.bldgnum AND tblRPentries_sum.userid = 'cotto' AND BillYrPeriod.BillYear = tblRPentries_sum.year ON UtilityBill.ypId = BillYrPeriod.ypId FULL OUTER JOIN eri_data.dbo.eri_total ON BillYrPeriod.BldgNum = eri_data.dbo.eri_total.bldg_no AND BillYrPeriod.BillPeriod = eri_data.dbo.eri_total.BillPeriod AND BillYrPeriod.BillYear = eri_data.dbo.eri_total.BillYear FULL OUTER JOIN tblBillByPeriod ON  UtilityBill.ypId = tblBillByPeriod.ypId WHERE (BillYrPeriod.BldgNum = '"& bldg &"') AND (BillYrPeriod.BillYear = '"& year &"') GROUP BY eri_data.dbo.eri_total.eri_amt, UtilityBill.TotalBillAmt,  BillYrPeriod.BillPeriod, tblRPentries_sum.Expr1 ORDER BY BillYrPeriod.BillPeriod"


rst1.Open strsql, cnn1, adOpenStatic


if not rst1.eof then
Dim Values(12, 6) 
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

Dim TableLabels(6)
tablelabels(1) = "Expenses"
tablelabels(2) = "Expense Adjustment"
tablelabels(3) = "Submeter"
tablelabels(4) = "ERI"
tablelabels(5) = "Revenue Adjustment"
tablelabels(6) = "Net" 

Dim i,x
Dim numRecords
numRecords = rst1.RecordCount
if numRecords > 12 then numRecords=12 end if
for x = 1 to numRecords
	if rst1("Expenses")<>"" then Values(x,1) = clng(rst1("Expenses")) else Values(x,1) = 0 end if
	if rst1("UnreportedAmts") < 0 then
		Values(x,2) = clng(rst1("UnreportedAmts")) * -1
		Values(x,5) = 0
	else
		if  rst1("UnreportedAmts") <> "" then Values(x,5) = clng(rst1("UnreportedAmts")) else Values(x,5) =0 end if
		Values(x,2) = 0
	end if
	
	if  rst1("eri_rev") <> "" then Values(x,3) = clng(rst1("eri_rev")) else Values(x, 3) = 0 end if 
	if rst1("SubMetered")<>"" then Values(x,4) = clng(rst1("SubMetered")) else Values(x,4) = 0 end if
	Values(x,6) = (clng(Values(x,3)) + clng(Values(x,4)) + clng(Values(x,5))) - (clng(Values(x,2)) + clng(Values(x,1)))
	
rst1.MoveNext
next
%>
          <p>&nbsp;</p>
      <table width="810" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC" align="center">
        <tr> 
              <td bordercolor="#FFFFFF"><font face="Arial, Helvetica, sans-serif"></font> 
              </td>
              <td> </td>
              <td> </td>
              <% for i = 1 to numRecords %>
              <td> 
                <div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=Title(i)%></font> 
                </div>
              </td>
              <% next %>
            </tr>
            <% For x = 1 to 6 %>
            <tr valign="bottom"> 
              <td><font face="Arial, Helvetica, sans-serif" size="1"><%=TableLabels(x)%></font> 
              <td> 
              <td> </td>
              <% for i = 1 to numRecords %>
              <td> 
                <div align="right"><font size="1" face="Arial, Helvetica, sans-serif" <% if values(i,x) < 0 then %>color="#FF0000" <%end if%>><%=FormatCurrency(Values(i,x)/1000,1)%> 
                  k</font> </div>
              </td>
              <% next %>
            </tr>
            <% next %>
          </table>
          
      <table width="810" border="0" cellspacing="0" cellpadding="0" align="center" height="200">
        <tr> 
              
          <td height="33"> 
            <div align="left"><img src="<%="makechartyrly.asp?bldgNUM=" & bldg & "&year=" & year %>"></div>
          </td>
            </tr>
          </table>
      </td>
    </tr>
  </table>
</body>
<%end if
end if%>
</html>