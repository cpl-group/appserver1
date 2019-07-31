<%option explicit

dim  b, date1, date2, utype, user
b = request.querystring("b")
utype = request.querystring("utype")
date1 = request.querystring("date1")
date2 = request.querystring("date2")
%>
<html>
<head>
<title></title>
<script>
function sendnewdates()
{	parent.document.forms['form1'].date1.value = document.forms['form1'].date1.value;
	parent.document.forms['form1'].date2.value = document.forms['form1'].date2.value;
	parent.loadchart();
	parent.loadoptions();
}

</script>
</head>
<body>

<body bgcolor="#FFFFFF" onload="parent.closeLoadBox('loadFrame2')" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#000000" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Yearly Comparison</b></font></td>
    <td bgcolor="#000000" width="46%" height="2">
      <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:parent.loadoptions()" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Options</a></b></font></div>
    </td>
  </tr>
  <tr>
    <td width="46%">&nbsp;</td>
    <td width="46%">&nbsp;</td>
  </tr>
</table>
<form method="get" name="form1">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr style="color:black; font-family: Arial, Helvetica, sans-serif; font-size: 12;"><td align="center">View 1<br>
<select name="date1">
	<%
	dim rstmin
	set rstmin = server.createobject("ADODB.recordset")
	dim cyear,myear,i
	cyear = year(date())
	myear = cyear-5
	rstmin.open "SELECT TOP 1 BillYear FROM BillYrPeriod WHERE BldgNum=33 ORDER BY BillYear", application("cnnstr_genergy1")
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
<td align="center">View 2<br>
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
</table>
<input type="button" onclick="sendnewdates()" value="Submit">
</form>
