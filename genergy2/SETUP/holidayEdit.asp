<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, hid
rid = secureRequest("rid")
hid = secureRequest("hid")

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
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body>
<form name="form2" method="post" action="holidaySave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#3399cc">
	<td colspan="2"><span class="standardheader">
		<%if trim(hid)<>"" then%>
			Update Holiday
		<%else%>
			Add New Holiday
		<%end if%>
	</span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Holiday</span></td> 
	<td><input type="text" name="holiday" value="<%=holiday%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Date</span></td>
	<td><span class="standard"><input type="text" name="holidaydate" value="<%=holidaydate%>"> DD/MM/YYYY</span></td>
</tr>
<tr bgcolor="#cccccc"> 
	<td><span class="standard">&nbsp;</span></td>
	
	<td>
		<%if trim(hid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
			<input type="button" name="action" value="Cancel" onclick="location='holidayView.asp?rid=<%=rid%>';" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
			<input type="button" name="action" value="Cancel" onclick="location='holidayView.asp?rid=<%=rid%>';" class="standard" style="background-color:ccf3cc;border-top:2px solid #ddffdd;border-left:2px solid #ddffdd;">
		<%end if%>
	</td>
</tr>
</table>
<input type="hidden" name="rid" value="<%=rid%>">
<input type="hidden" name="hid" value="<%=hid%>">
</form>
</body>
</html>
