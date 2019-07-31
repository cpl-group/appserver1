<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim building, pid, utility, byear, bperiod
byear = request("byear")
bperiod = request("bperiod")

dim rst1, cnn1, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

if trim(byear)<>"" and trim(bperiod)<>"" then
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "sp_oucbcp"
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	cmd.parameters("by") = byear
	cmd.parameters("bp") = bperiod
	cmd.execute
end if
%>
<html><head><title>OUC Report</title>
<link rel="Stylesheet" href="../setup/setup.css" type="text/css">
</head>
<body link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td bgcolor="#6699cc"><span class="standardheader">Download OUC Output File</span></td>
</tr>
</table>
<form action="oucOutputReport.asp" method="post">
<select name="byear">
<option value="">Select Bill Year</option>
<%rst1.open "SELECT Distinct BillYear FROM BillYrPeriod byp, buildings b, portfolio p WHERE byp.bldgnum=b.bldgnum and b.portfolioid=p.id and p.portfolio='OUC' order by billyear", cnn1
do until rst1.eof
	%><option value="<%=rst1("Billyear")%>"<%if trim(rst1("billyear"))=trim(byear) then response.write " SELECTED"%>><%=rst1("Billyear")%></option><%
	rst1.movenext
loop
rst1.close
%>
</select>
<select name="bperiod">
<option value="">Select Bill Period</option>
<%rst1.open "SELECT Distinct billperiod, billyear FROM BillYrPeriod byp, buildings b, portfolio p WHERE byp.bldgnum=b.bldgnum and b.portfolioid=p.id and p.portfolio='OUC' order by billyear, billperiod", cnn1
do until rst1.eof
	%><option value="<%=rst1("billperiod")%>"<%if trim(rst1("billperiod"))=trim(bperiod) then response.write " SELECTED"%>><%=rst1("BillPeriod")%></option><%
	rst1.movenext
loop
rst1.close
%>
</select>
<input type="submit" value="View Download"><br>
</form>
<%if trim(byear)<>"" and trim(bperiod)<>"" then%>
Download Files:<br>
<a href="/eri_TH/sqldownload/OUC/ouc<%=byear%><%=bperiod%>.zip">OUC Totals</a>
<%end if%>
</body>
</html>
