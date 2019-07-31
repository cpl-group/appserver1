<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")


%>
<html>
<head>
<title>Building View</title>
<script>
function editPoint(pid)
{	document.location.href = 'premiseEdit.asp?pid='+pid;
}
</script>
</head>

<body>
<form name="form2" method="post" action="buildingsave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#0099ff">
	<td colspan="2" style="font-family:Arial;color:#FFFFFF;font-size:12px"><b>OUC Premise Listing</b></td>
</tr>
</table>
<input type="button" value="Add Data Point Reference" onclick="editPoint('');">

<%
	rst1.Open "SELECT * FROM oucDataKey", cnn1
	if not rst1.EOF then%>
		List of Tenants:<BR><table width="100%" border="0" cellpadding="3" cellspacing="1">
		<tr bgcolor="#cccccc">
			<td><span class="standard"><b>Premise Name</b></span></td>
			<td><span class="standard"><b>Point Name</b></span></td>
			<td><span class="standard"><b>Meter ID</b></span></td>
			<td><span class="standard"><b>Column Name</b></span></td>
		</tr>

		<%do until rst1.EOF%>
		<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="editPoint('<%=rst1("id")%>');">
			<td><span class="standard"><%=rst1("premise")%></span></td>
			<td><span class="standard"><%=rst1("pointName")%></span></td>
			<td><span class="standard"><%=rst1("meterid")%></span></td>
			<td><span class="standard"><%=rst1("keyname")%></span></td>
		</tr>
		
		<%rst1.movenext
		loop%>
	</table>
	<%
	else
		Response.Write "<BR>There are no premises set up."
	end if
	rst1.close
%>

</form>
</body>
</html>