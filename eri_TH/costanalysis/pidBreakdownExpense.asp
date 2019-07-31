<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim date1, date2, b, utype, pid, i,pidAccountlist
b = request.querystring("b")
pid = request.querystring("pid")
utype = request.querystring("utype")
date1 = request.querystring("date1")
date2 = request.querystring("date2")

dim cnn1, rst1, sql
set cnn1 = server.createobject("ADODB.Connection")
set rst1 = server.createobject("ADODB.recordset")
set pidAccountlist = server.createobject("ADODB.recordset")
cnn1.Open application("Cnnstr_genergy1")
sql = "SELECT DISTINCT AcctID FROM tblAcctSetup [as] join buildings b on [as].bldgnum=b.bldgnum where b.portfolioid = '"&pid&"' order by acctid"
pidAccountlist.open sql, cnn1,1,1
%>
<html>
<head>
<title></title>
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

<body bgcolor="#FFFFFF" text="#000000" onload="parent.closeLoadBox('loadFrame2')" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="706" border="0" cellspacing="0" cellpadding="0">
	<tr><td bgcolor="#000000" width="50%"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"><b>Expense Break Down</b></font></td>
		<td bgcolor="#000000" width="50%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location.href='monthlyDetails.asp?b=<%=b%>&pid=<%=pid%>&date1='+ parent.document.forms['form1'].date1.value +'&date2='+ parent.document.forms['form1'].date2.value +'&utype='+ parent.document.forms['form1'].utype.value" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Monthly Details</a></b></font><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
	</tr>
</table>


<table border="0" cellspacing="0" cellpadding="0">
<tr><td valign="top">
<table width="155" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
<td>Acount</td></tr>
<%
while not pidAccountlist.eof
sql ="SELECT [as].acctid,[as].bldgnum, billperiod ,TotalBillAmt from utilitybill ub join BillYrPeriod bp on ub.ypid=bp.ypid join tblAcctSetup [as] on [as].bldgnum=bp.bldgnum WHERE [as].acctid='"&pidAccountlist("acctid")&"' and billyear='"&date1&"' order by billperiod"
rst1.open sql, cnn1
%>
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
<td><nobr><%if not(rst1.eof) then response.write rst1("acctid") & " ("& rst1("bldgnum")&")"%></nobr></td>
<%
pidAccountlist.movenext
rst1.close
wend
pidAccountlist.movefirst

%>
</table>
</td><td valign="top"><div style="width:550; overflow:auto; height: <%=pidAccountlist.recordcount*15+33%>;">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
<%for i = 1 to 12%>
	<td align="center" ><%=left(monthname(i),3)%></td>
<%next%>
</tr>
<%
while not pidAccountlist.eof
sql ="SELECT [as].acctid,[as].bldgnum, billperiod ,TotalBillAmt from utilitybill ub join BillYrPeriod bp on ub.ypid=bp.ypid join tblAcctSetup [as] on [as].bldgnum=bp.bldgnum WHERE [as].acctid='"&pidAccountlist("acctid")&"' and billyear='"&date1&"' order by billperiod"
rst1.open sql, cnn1
%>
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
<%
do until rst1.eof
	response.write "<td align='left'>"& formatcurrency(rst1("TotalBillAmt")) &"</td>"
	rst1.movenext
loop
rst1.close
pidAccountlist.movenext
wend
pidAccountlist.movefirst
%>

</tr>
</table>
</div>
</td>
</tr>
</table>
</body>
</html>