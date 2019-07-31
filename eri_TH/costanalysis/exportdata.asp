<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim date1, date2, b, utype, pid, ddate, datearray, dateST, dateEN, periodST, periodEN, downloadtype
b = request.querystring("b")
pid = request.querystring("pid")
utype = request.querystring("utype")
date1 = request.querystring("date1")
date2 = request.querystring("date2")
dateST = request("dateST")
dateEN = request("dateEN")
periodST = request("periodST")
periodEN = request("periodEN")
downloadtype = request("downloadtype")
dim tempdate
if dateEN<dateST then
	tempdate = dateEN
	dateEN = dateST
	dateST = tempdate
	tempdate = periodEN
	periodEN = periodST
	periodST = tempdate
end if
if dateEN=dateST and periodEN<periodST then
	tempdate = periodEN
	periodEN = periodST
	periodST = tempdate
end if
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
<%
if request.form("submit")="Download" then

dim cnn1, rst1, cmd, prm
Set rst1 = Server.CreateObject("ADODB.recordset")
set cnn1 = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rst1 = server.createobject("ADODB.Recordset")
cnn1.Open application("Cnnstr_genergy1")
cnn1.CursorLocation = adUseClient
cmd.CommandType = adCmdStoredProc
cmd.Name = "getdata"
Set cmd.ActiveConnection = cnn1
if trim(downloadtype)="submetered" then
	cmd.CommandText = "sp_subdownload"
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("from", adInteger, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("files", adVarChar, adParamOutput, 50)
	cmd.Parameters.Append prm
	cnn1.getdata b, dateST
else
	cmd.CommandText = "sp_revdownload"
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("from", adInteger, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("fromp", adInteger, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("to", adInteger, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("top", adInteger, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("pid", adVarChar, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("files", adVarChar, adParamOutput, 50)
	cmd.Parameters.Append prm
	cnn1.getdata b, dateST, periodST, dateEN, periodEN, pid
end if
'sp_subdownload @bldg,@from,@file output


%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td bgcolor="#000000" align="center"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Download Data for <%=request.form("ddate")%></b></font></td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td bgcolor="#0099FF" align="center"><a href="https://appserver1.genergy.com/eri_TH/sqldownload/<%=cmd.Parameters("files")%>" style="text-decoration:none;color:black" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'white'"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>Click Here to Download Data File</b></font></a></td></tr>
<tr><td>&nbsp;</td></tr>
<!-- <tr><td height="18" align="center"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location.back()" style="text-decoration:none;color:black" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'black'">Return Export Data</a></b></font></td></tr>
<tr><td height="18"><div align="center"><hr width="100"></div></td></tr> -->
<tr><td height="18" align="center"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:parent.loadoptions()" style="text-decoration:none;color:black" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'black'">Return Options</a></b></font></td></tr>
</table>
<%else%>
<form name="form1" method="post" action="">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td bgcolor="#000000" width="50%"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"><b>Export Data</b></font></td>
		<td bgcolor="#000000" width="50%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:parent.loadoptions()" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Options</a></b></font><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"></font></td>
	</tr>
</table>
<span style="font-family:Arial, Helvetica, sans-serif; font-size:12">
Select a beginning and ending for the data you wish to download.<br>&nbsp;<br>
<table border="0" cellspacing="0" cellpadding="0">
<tr style="font-family:Arial, Helvetica, sans-serif; font-size:12"><td>
Begining:<br>
<select name="dateST">
	<%
	dim rstmin
	set rstmin = server.createobject("ADODB.recordset")
	dim cyear,myear,i
	cyear = year(date())
	myear = cyear-5
	rstmin.open "SELECT TOP 1 BillYear FROM BillYrPeriod WHERE BldgNum='" & b & "' and billyear >= 2002 ORDER BY BillYear", application("cnnstr_genergy1")
	if not(rstmin.EOF) then myear=trim(rstmin("BillYear"))
	rstmin.close
	for i = myear to cyear
		response.write "<option value="""& i &""""
		if i=cint(date1) then response.write " SELECTED"
		response.write ">"& i &"</option>"
	next
	%>
</select>
<select name="periodST" id="beginningPeriod" style="display: inline;">
<%
for i = 1 to 12
	response.write "<option value="""&i&""">"& i &"</option>"
next
%>
</select><br>
</td><td>&nbsp;&nbsp;</td><td>
<div id="ending" style="display: inline;">
Ending:<br>
<select name="dateEN">
	<%
	set rstmin = server.createobject("ADODB.recordset")
	cyear = year(date())
	myear = cyear-5
	rstmin.open "SELECT TOP 1 BillYear FROM BillYrPeriod WHERE BldgNum='" & b & "' and billyear >= 2002 ORDER BY BillYear", application("cnnstr_genergy1")
	if not(rstmin.EOF) then myear=trim(rstmin("BillYear"))
	rstmin.close
	for i = myear to cyear
		response.write "<option value="""& i &""""
		if i=cint(date1) then response.write " SELECTED"
		response.write ">"& i &"</option>"
	next
	%>
</select>
<select name="periodEN">
<%
for i = 1 to 12
	response.write "<option value="""&i&""">"& i &"</option>"
next
%>
</select><br>
</td>
</tr><tr style="font-family:Arial, Helvetica, sans-serif; font-size:12">
<td colspan="3" valign="top"><input type="radio" name="downloadtype" value="all" onclick="document.all['ending'].style.display='inline';document.all['beginningPeriod'].style.display='inline';" checked>&nbsp;Download&nbsp;All<br><input type="radio" name="downloadtype" value="submetered" onclick="document.all['ending'].style.display='none';document.all['beginningPeriod'].style.display='none';">&nbsp;Download&nbsp;Submetered&nbsp;</td>
</tr></table>
<br><input type="submit" name="submit" value="Download" onclick="parent.openLoadBox('loadFrame2')">
</div>
</font>
<%end if%>
</body>
</html>
