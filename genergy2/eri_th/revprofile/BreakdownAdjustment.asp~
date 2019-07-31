<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim date1, date2, bldg, utype, pid, adjtype, utilitywhere, utildisplay(15)
bldg = request.querystring("bldg")
pid = request.querystring("pid")
date1 = request.querystring("date1")
date2 = request.querystring("date2")
utype = request.querystring("utype")
adjtype = request.querystring("adjtype")
dim title, sqlsign
if adjtype="exp" then
	title = "Expense"
	sqlsign = "0"
else
	title = "Revenue"
	sqlsign = "1"
end if
if utype<>0 then utilitywhere = " and utility="&utype

dim rst1, cnn1, sql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)

rst1.open "SELECT * FROM tblutility", getConnect(pid,bldg,"billing")
do until rst1.eof
  utildisplay(cint(rst1("utilityid"))) = rst1("utilitydisplay")
  rst1.movenext
loop
rst1.close

sql = "select *, convert(varchar,entrydate,101) as date from tblRPentries where bldgnum='"&bldg&"' and year ='"&date1&"' AND type="&sqlsign&" "&utilitywhere&" ORDER BY period"
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
		<td bgcolor="#000000" width="50%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location.href='monthlyDetails.asp?bldg=<%=bldg%>&pid=<%=pid%>&date1='+ parent.document.forms['form1'].date1.value +'&date2='+ parent.document.forms['form1'].date2.value +'&utype='+ parent.document.forms['form1'].utype.value" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Monthly Details</a></b></font><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
	</tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
	<td align="center">Date</td>
	<td align="center">Description</td>
	<td align="center">Period</td>
	<td align="center">Total Amount</td>
	<td align="center">Utility</td>
</tr>
<%do until rst1.eof%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
		<td aling="center"><%=rst1("date")%></td>
		<td><%=rst1("description")%></td>
		<td align="center"><%=rst1("period")%></td>
		<td align="right"><%=FormatCurrency(rst1("amt"))%></td>
		<td align="right"><%=utildisplay(cint(rst1("utility")))%></td>
	</tr>
	<%rst1.movenext
loop%>
</table>
