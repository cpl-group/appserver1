<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Meter Details - Usage History</title>
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

<script> 
function loadentry(luid, ypid){

	var temp = 'meterbillsummaryPDF.asp?l=' +luid+'&Y='+ypid+'&building=<%=server.urlencode(request("b"))%>'
	parent.document.frames.panel_2.location = temp
}
function settonull(){
	var temp = 'null.htm'
	parent.document.frames.panel_2.location = temp
}
</script>
<body bgcolor="#FFFFFF"onLoad="settonull()">
<%
bldg = request.querystring("B")
leaseid = request.querystring("leaseid")

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy1")
if leaseid<>0 then
	title = "Tenant"
	sql="SELECT top 12 TotalAmt,b.BillYear, b.BillPeriod, sum(m.KWHUsed) as KWH, sum(Demand_P) as demand_P, m.ypid, m.leaseutilityid FROM tblmetersbyperiod m INNER JOIN tblbillbyperiod b on b.ypid=m.ypid and b.LeaseUtilityId=m.LeaseUtilityId WHERE meterid in (select meterid from meters where leaseutilityid =" & leaseid & ") and b.bldgnum='" & bldg & "' and b.posted = 1 GROUP BY b.BillYear, b.BillPeriod, m.ypid, m.leaseutilityid, TotalAmt order by m.ypid desc"
else
	title = "Building"
	sql="SELECT top 12 sum(TotalAmt) as TotalAmt, BillYear, BillPeriod, sum(energy) as KWH, sum(Demand) as demand_P, ypid, 0 as leaseutilityid FROM tblbillbyperiod WHERE bldgnum='" & bldg & "' and posted = 1 GROUP BY BillYear, BillPeriod, ypid order by ypid desc"
end if
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><%=title%> Details - </b>Usage History</font></td>
  </tr>
  <tr> 
    <td width="46%">&nbsp;</td>
  </tr>
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
	<tr bgcolor="#0099FF"> 
		<td width="25%" align="center"><b><font size="1" face="Arial">Bill Period</font></b></td>
		<td width="25%" align="center"><b><font size="1" face="Arial">Kwhr</font></b></td>
		<td width="25%" align="center"><b><font size="1" face="Arial">Demand</font></b></td>
		<td width="25%" align="center"><b><font size="1" face="Arial">Total Amount</font></b></td>
	</tr>
</table>

<%while not rst1.eof%>

<table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
	<tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=rst1("leaseutilityid")%>','<%=rst1("ypid")%>')"> 
		<td width="25%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("billyear")%>/<%=rst1("billperiod")%></font></b></td>
		<td width="25%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("kwh"),0)%></font></b></td>
		<td width="25%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst1("demand_P"),2)%></font></b></td>
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
  <p align="left"><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click 
    Any Bill Period Row To View Meter Details Below</i></b></font></p>
</div>
</body>

</html>
