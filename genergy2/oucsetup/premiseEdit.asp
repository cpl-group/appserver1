<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")

dim pid, premise, pointName, meterid, ptype, keyName
pid = request("pid")

if trim(pid)<>"" then
	rst1.Open "SELECT * FROM oucDataKey WHERE id='"&pid&"'", cnn1
	
	if not rst1.EOF then
		premise = rst1("premise")
		pointName = rst1("pointName")
		ptype = rst1("type")
		keyname = rst1("keyname")
		meterid = rst1("meterid")
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Point Setup</title>
</head>

<body>
<form name="form2" method="post" action="premiseSave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#0099ff" style="font-family:Arial;color:#FFFFFF;font-size:12px">
	<td colspan="2"><span class="standardheader">
		<b><%if trim(pid)<>"" then%>
			Update Point
		<%else%>
			Add New Point
		<%end if%></b>
	</span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Premise</span></td> 
	<td><input type="text" name="premise" value="<%=Premise%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Point Name</span></td>
	<td><input type="text" name="pointName" value="<%=PointName%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Type</span></td>
	<td><select name="ptype">
			<option value="decimal(18,4)" <%if trim(ptype)="decimal(18,4)" then response.write "SELECTED"%>>decimal(18,4)</option>
			<option value="integer" <%if trim(ptype)="integer" then response.write "SELECTED"%>>integer</option>
			<option value="varchar(50)" <%if trim(ptype)="varchar(50)" then response.write "SELECTED"%>>varchar(50)</option>
			<option value="datetime" <%if trim(ptype)="datetime" then response.write "SELECTED"%>>datetime</option>
		</select>
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Meterid</span></td>
	<td>
	<select name="meterid" onchange="loadbuilding()">
	<%
	rst1.open "SELECT * FROM meters ORDER BY meternum", cnn1
	do until rst1.eof%>
		<option value="<%=trim(rst1("meterid"))%>"<%if trim(rst1("meterid"))=trim(meterid) then response.write " SELECTED"%>><%=rst1("meternum")%>/<%=rst1("meterid")%></option>
	<%	rst1.movenext
	loop
	rst1.close
	%>
	</select>	
	</td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Key Name</span></td>
	<td><input type="text" name="keyname" value="<%=KeyName%>"></td>
</tr>
<tr bgcolor="#cccccc"> 
	<td><span class="standard">&nbsp;</span></td>
	
	<td>
		<%if trim(pid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard"> <input type="submit" name="action" value="Delete" class="standard">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard">
		<%end if%>
	</td>
</tr>
</table>
<input type="hidden" name="oldkey" value="<%=KeyName%>">
<input type="hidden" name="oldmeterid" value="<%=meterid%>">
<input type="hidden" name="pid" value="<%=pid%>">
</form>
</body>
</html>
