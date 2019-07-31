<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")

dim rid, hid
rid = request("rid")
hid = request("hid")

dim holiday, holidaydate, action
if trim(hid)<>"" then
	rst1.Open "SELECT * FROM rateholiday WHERE id="&hid, cnn1
	if not rst1.EOF then
		holiday = rst1("holiday")
		holidaydate = rst1("date")
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Building View</title>
</head>

<body>
<form name="form2" method="post" action="holidaySave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#0099ff">
	<td colspan="2"><span class="standardheader">
		<%if trim(hid)<>"" then%>
			Update Region
		<%else%>
			Add New Region
		<%end if%>
	</span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Holiday</span></td> 
	<td><input type="text" name="holiday" value="<%=holiday%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Date</span></td>
	<td><input type="text" name="holidaydate" value="<%=holidaydate%>"></td>
</tr>
<tr bgcolor="#cccccc"> 
	<td><span class="standard">&nbsp;</span></td>
	
	<td>
		<%if trim(hid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard">
		<%end if%>
	</td>
</tr>
</table>
<input type="hidden" name="rid" value="<%=rid%>">
<input type="hidden" name="hid" value="<%=hid%>">
</form>
</body>
</html>
