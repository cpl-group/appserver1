<%option explicit

dim date1, date2, b, utype, pid, adjtype
b = request.querystring("b")
pid = request.querystring("pid")
date1 = request.querystring("date1")
date2 = request.querystring("date2")
utype = request.querystring("utype")
adjtype = request.querystring("adjtype")
dim title, sqlsign
if adjtype="exp" then
	title = "Expense"
	sqlsign = "<"
else
	title = "Revenue"
	sqlsign = ">"
end if

dim rst1, cnn1, sql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy1")
sql = "select *, convert(varchar,entrydate,101) as date from tblRPentries where pid='"&pid&"' and bldgnum='"&b&"' and year ='"&date1&"' and amt"&sqlsign&"0 ORDER BY period"
rst1.open sql, cnn1%>
<html>
<head><title></title>
<style type="text/css">
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
</head>
<body bgcolor="#FFFFFF" text="#000000" onload="parent.closeLoadBox('loadFrame2')" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="706" border="0" cellspacing="0" cellpadding="0">
	<tr><td bgcolor="#000000" width="50%"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"><b><%=title%> Adjustment Breakdown</b></font></td>
		<td bgcolor="#000000" width="50%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location.href='monthlyDetails.asp?b=<%=b%>&pid=<%=pid%>&date1='+ parent.document.forms['form1'].date1.value +'&date2='+ parent.document.forms['form1'].date2.value +'&utype='+ parent.document.forms['form1'].utype.value" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Monthly Details</a></b></font><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"></font></td>
	</tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
	<td align="center">Date</td>
	<td align="center">Description</td>
	<td align="center">Period</td>
	<td align="center">Total Amount</td>
</tr>
<%do until rst1.eof%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
		<td><%=rst1("date")%></td>
		<td align="right"><%=rst1("description")%></td>
		<td align="right"><%=rst1("period")%></td>
		<td align="right"><%=FormatCurrency(rst1("amt"))%></td>
	</tr>
	<%rst1.movenext
loop%>
</table>