<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim  b, date1, date2, utype, user, pid
b = request.querystring("b")
pid = request.querystring("pid")
utype = request.querystring("utype")
date1 = request.querystring("date1")
date2 = request.querystring("date2")

dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_security")
%>
<html>
<head>
<title></title>
<script>
function sendnewdates(date1, date2)
{	parent.document.forms['form1'].date1.value = date1;
	parent.document.forms['form1'].date2.value = date2;
	parent.loadchart();
	parent.loadoptions();
}

function nullfunction()//for null href with onclick actions
{
}

</script>
</head>
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

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000" onload="parent.closeLoadBox('loadFrame2');">
<table width="706" border="0" cellspacing="0" cellpadding="0">
	<tr><td bgcolor="#000000" width="50%"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"><b>Revenue Profile Options</b></font></td>
		
    <td bgcolor="#000000" width="50%" align="right"> 
    </td>
	</tr>
</table>
&nbsp;<br>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
  <tr>
    <td width="51%" valign="top"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
<%
'addons available
rst1.open "SELECT tbladdons.SID, Label, Link, Target, Active FROM tbladdonlinks JOIN tbladdons on tbladdons.SID=tbladdonlinks.SID WHERE CSID=12 and userid='"&session("userid")&"' ORDER BY listorder", cnn1
if rst1.eof then response.write "Client has no options"
'response.write "SELECT tbladdons.SID, Label, Link, Target, Active FROM tbladdonlinks JOIN tbladdons on tbladdons.SID=tbladdonlinks.SID WHERE CSID=12 and userid='"&session("userid")&"' ORDER BY listorder"
'response.end
do while not(rst1.eof)
	if trim(rst1("SID"))=11 then
		response.write "<a href=""javascript:parent.loadoptions()"" onclick=""javascript:window.open('"&rst1("Link")&"?building="&b&"&pid="&pid&"&date1='+ parent.document.forms['form1'].date1.value +'&date2='+ parent.document.forms['form1'].date2.value +'&utility='+ parent.document.forms['form1'].utype.value,'', 'scrollbars=no,width=450, height=480, status=no' );"" onMouseOver=""this.style.color='gray'"" onMouseOut=""this.style.color='Black'"" onclick=""parent.openLoadBox('loadFrame2')"">"&rst1("Label")&"</a><br>"
	elseif trim(rst1("SID"))=9 then
		response.write "<a href=""javascript:parent.loadoptions()"" onclick=""javascript:window.open('"&rst1("Link")&"?building="&b&"&date1='+ parent.document.forms['form1'].date1.value +'&date2='+ parent.document.forms['form1'].date2.value +'&utility='+ parent.document.forms['form1'].utype.value,'', 'scrollbars=yes,width=800, height=600, status=no' );"" onMouseOver=""this.style.color='gray'"" onMouseOut=""this.style.color='Black'"" onclick=""parent.openLoadBox('loadFrame2')"">"&rst1("Label")&"</a><br>"
	elseif trim(rst1("SID"))=8 then
		response.write "<a href=""javascript:nullfunction()"" onclick=""javascript:document.all['comparison'].style.visibility='visible'"" onMouseOver=""this.style.color='gray'"" onMouseOut=""this.style.color='Black'"" onclick=""parent.openLoadBox('loadFrame2')"">"&rst1("Label")&"</a><br>"
	else	
		response.write "<a href=""javascript:document.location.href='"&rst1("Link")&"?b="&b&"&pid="&pid&"&date1='+ parent.document.forms['form1'].date1.value +'&date2='+ parent.document.forms['form1'].date2.value +'&utype='+ parent.document.forms['form1'].utype.value"" onMouseOver=""this.style.color='gray'"" onMouseOut=""this.style.color='Black'"" onclick=""parent.openLoadBox('loadFrame2')"">"&rst1("Label")&"</a><br>"
	end if
    rst1.movenext
loop
rst1.close
'https://appserver1.genergy.com/um/um/invoiceview4del.asp?ypid=3625&bldg=011&luid=524
'addons NOT available
rst1.open "SELECT Label FROM tbladdons WHERE SID not in (SELECT SID FROM tbladdonlinks WHERE userid='" &session("userid")& "' and active=1) AND CSID=12 ORDER BY listorder", cnn1
do while not(rst1.eof)
    response.write "<li style=""color:cccccc"">" &rst1("Label")& "</li>"
    rst1.movenext
loop
%>
</p><div id="comparison" style="width:200;height:100;overflow:hidden;visibility:hidden;position:absolute;left:300;top:30;>
<form method="get" name="form1">
<table border="0" cellspacing="0" cellpadding="0">
<tr style="color:black; font-family: Arial, Helvetica, sans-serif; font-size: 12;"><td>View 1<br>
<select name="date1">
	<%
	dim rstmin
	set rstmin = server.createobject("ADODB.recordset")
	dim cyear,myear,i
	cyear = year(date())
	myear = cyear-5
	rstmin.open "SELECT TOP 1 BillYear FROM BillYrPeriod WHERE BldgNum='" & b & "' and billyear >= 2000 ORDER BY BillYear", application("cnnstr_genergy1")
	if not(rstmin.EOF) then myear=trim(rstmin("BillYear"))
	rstmin.close
	for i = myear to cyear
		response.write "<option value="""& i &""""
		if i=cint(date1) then response.write " SELECTED"
		response.write ">"& i &"</option>"
	next
	%>
</select>
</td>
<td>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>View 2<br>
<select name="date2">
	<option value="">none</option>
	<%
	cyear = year(date())
	for i = myear to cyear
		response.write "<option value="""& i &""""
		if date2<>"" then if i=cint(date2) then response.write " SELECTED"
		response.write ">"& i &"</option>"
	next
	%>
</select>
</td>
</tr>
</table>&nbsp;<br>
<input type="button" onclick="sendnewdates(date1.value, date2.value)" value="Compare">
</form>
</div>