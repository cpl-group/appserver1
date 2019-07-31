<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")

dim rid
rid = request("rid")

%>
<html>
<head>
<title>Portfolio View</title>
<script>
function holidayEdit(hid)
{	document.location = 'holidayEdit.asp?rid=<%=rid%>&hid='+hid
}
</script>
</head>
<body>
<FORM>
<table width="100%" border="0" bgcolor="#FFFFFF"><tr><td bgcolor="#3399CC" align="center"><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Region View</font></b></td></tr></table>
<%
Response.Write "<input type=""button"" value=""Add Holiday"" onclick=""holidayEdit('');"">"
rst1.Open "SELECT * FROM rateholiday where regionid="&rid&" order by date desc", cnn1
if not rst1.EOF then%>
	<BR><table width="100%" border="0" cellpadding="3" cellspacing="1">
	<tr bgcolor="#cccccc">
		<td><span class="standard"><b>Holiday Name</b></span></td>
		<td><span class="standard"><b>Date</b></span></td>
	</tr>

	<%do until rst1.EOF%>
	<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="holidayEdit('<%=rst1("id")%>');">
		<td><span class="standard"><%=rst1("holiday")%></span></td>
		<td><span class="standard"><%=rst1("date")%></span></td>
	</tr>
	
	<%rst1.movenext
	loop%>
</table>
<%
else
	Response.Write "<BR>There are no holidays setup for this region."
end if
rst1.close
%>
</FORM>
</body>
</html>
