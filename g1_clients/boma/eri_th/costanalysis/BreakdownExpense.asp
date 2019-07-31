<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim date1, date2, b, utype, pid, i
b = request.querystring("b")
pid = request.querystring("pid")
utype = request.querystring("utype")
date1 = request.querystring("date1")
date2 = request.querystring("date2")

dim cnn1, rst1, sql, cmd, prm
set cnn1 = server.createobject("ADODB.Connection")
set rst1 = server.createobject("ADODB.recordset")
set cmd = server.createobject("ADODB.Command")
cnn1.Open application("Cnnstr_genergy1")

cmd.CommandType = adCmdStoredProc
cmd.ActiveConnection = cnn1
cmd.CommandText = "sp_ExpBreakdown"
'cmd.CursorType = adOpenDynamic

Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("util", adVarChar, adParamInput, 30)
cmd.Parameters.Append prm

cmd.Parameters("bldg") = b
cmd.Parameters("by") = date1
cmd.Parameters("util") = utype
set rst1 = cmd.execute
'rst1.CursorType = adOpenDynamic
'sql = "SELECT ub.Acctid, billperiod ,TotalBillAmt from utilitybill ub join BillYrPeriod bp on ub.ypid=bp.ypid join Buildings b on b.bldgnum=bp.bldgnum WHERE bp.bldgnum='"&b&"' and billyear='"&date1&"' order by billperiod"
'response.write sql
'rst1.open sql, cnn1
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
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Acount</td></tr>
<%
dim acctid, acctnumber
do until rst1.eof%>
	<%if not(rst1.eof) then 
		if acctid<>trim(rst1("Acctid")) then%>
		<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td><nobr>
			<%response.write rst1("acctid")
			acctid = trim(rst1("Acctid"))
			acctnumber = acctnumber + 1
		end if
		%></nobr></td></tr>
	<%end if
	rst1.movenext
loop%>
</table>
</td><td valign="top"><div style="width:550; overflow:auto; height: <%=acctnumber*15+33%>;">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
<%for i = 1 to 12%>
	<td align="center"><%=left(monthname(i),3)%></td>
<%next%>
</tr>
<%
dim monthi
rst1.close
acctid = ""
set rst1 = cmd.execute
do until rst1.eof
%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><%
	monthi = 1
	acctid = trim(rst1("Acctid"))
	do until rst1.eof
		if acctid<>trim(rst1("Acctid")) or (isnull(acctid) <> isnull(rst1("Acctid"))) or monthi>12 then exit do
		if cint(rst1("billperiod")) = monthi then
			if cdbl(rst1(3)) <> 0 then 
				response.write "<td>"& formatcurrency(rst1(3)) &"</td>"
			else
				response.write "<td>0</td>"
			end if 
			rst1.movenext
		else
			response.write "<td>0</td>"
		end if
		monthi = monthi+1
	loop
%></tr><%
loop
%>
</table></div>
</td></tr></table>
</body>
</html>
