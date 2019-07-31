<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->

<html>
<head>
<title>Building Bill History</title>
<script>
function loadentry(building,billyear,billperiod)
{	var temp= "../pdfMaker/pdfBatchPrint.asp?building=" + building + "&byear=" + billyear + "&bperiod=" + billperiod
	alert(temp);
	window.open(temp,'','statusbar=no, menubar=no,scrollbars=yes, HEIGHT=800, WIDTH=700')
}
</script>
</head>
<body bgcolor="#FFFFFF"onLoad="settonull()">
<%
dim bldg, leaseid
bldg = request.querystring("bldg")
leaseid = request.querystring("leaseid")

dim cnn1, cmd1, rst1, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy1")

sql="SELECT top 12 sum(TotalAmt) as TotalAmt,b.BillYear, b.BillPeriod, sum(m.KWHUsed) as KWH, sum(Demand_P) as demand_P, b.DateStart, b.DateEnd FROM tblmetersbyperiod m INNER JOIN tblbillbyperiod b on b.ypid=m.ypid and b.LeaseUtilityId=m.LeaseUtilityId WHERE b.bldgnum='"&bldg&"' and b.posted = 1 GROUP BY b.BillYear, b.BillPeriod, b.DateStart, b.DateEnd ORDER BY b.BillYear desc, b.BillPeriod desc"
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
'response.write sql
'response.end
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Building Details - </b>Bill History</font></td>
  </tr>
  <tr> 
    <td width="46%">&nbsp;</td>
  </tr>
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
	<tr bgcolor="#0099FF"> 
		<td width="25%" align="center"><b><font size="1" face="Arial">Bill Period</font></b></td>
		<td width="25%" align="center"><b><font size="1" face="Arial">From</font></b></td>
		<td width="25%" align="center"><b><font size="1" face="Arial">To</font></b></td>
		<td width="25%" align="center"><b><font size="1" face="Arial">Total Amount</font></b></td>
	</tr>
</table>

<%while not rst1.eof%>

<table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
	<tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=bldg%>',<%=trim(rst1("billyear"))%>,<%=trim(rst1("billperiod"))%>)"> 
		<td width="25%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("billyear")%>/<%=rst1("billperiod")%></font></b></td>
		<td width="25%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("datestart")%></font></b></td>
		<td width="25%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("dateend")%></font></b></td>
		<td width="25%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatcurrency(rst1("TotalAmt"))%></font></b></td>
	</tr>
</table>


<%
rst1.movenext
wend
set cnn1 = nothing
%>
<div align="center">
  <p>&nbsp;</p>
  <p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click Any Bill Period Row To View All Bills For the Building</i></b></font></p>
</div>
</body>

</html>
