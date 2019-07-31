<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid
rid = secureRequest("rid")

dim city
if trim(rid)<>"" then
	rst1.Open "SELECT * FROM regions WHERE id='"&rid&"'", cnn1
	if not rst1.EOF then
		city = rst1("city")
	end if
	rst1.close
end if

%>
<html>
<head>
<title>Regional Holidays</title>
<script>
function holidayEdit(hid)
{	document.location = 'holidayEdit.asp?rid=<%=rid%>&hid='+hid
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<body>
<FORM>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#000000" colspan="2">
<%
dim showWeirdBlackBar
showWeirdBlackBar = false
if allowGroups("Genergy Users") AND showWeirdBlackBar then
%>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td><span class="standardheader"><a href="index.asp" target="main" class="breadcrumb" style="text-decoration:none;"><img src="images/aro-left-000.gif" align="left" width="13" height="13" border="0">Utility Manager Setup</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="frameset.asp" target="main" class="breadcrumb" style="text-decoration:none;">Update Meters</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="portfolioview.asp" target="main" class="breadcrumb" style="text-decoration:none;">Set Up Portfolios</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="regionView.asp" target="main" class="breadcrumb" style="text-decoration:none;">Set Up Rates</a></span></td>
  </tr>
  </table>
<%end if%>
  </td>
</tr>
<tr bgcolor="#3399cc">
	<td colspan="2"><span class="standardheader">
	Manage Holidays | <a href="regionedit.asp?rid=<%=rid%>" style="color:#ffffff;font-weight:normal;"><%=city%> Region</a>
	</span></td>
</tr>
<%
rst1.Open "SELECT * FROM rateholiday where regionid="&rid&" order by date desc", cnn1
if not rst1.EOF then%>
	<tr bgcolor="#cccccc">
		<td width="30%"><span class="standard"><b>Holiday Name</b></span></td>
		<td width="70%"><span class="standard"><b>Date</b></span></td>
	</tr>

	<%do until rst1.EOF%>
	<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="holidayEdit('<%=rst1("id")%>');">
		<td><span class="standard"><%=rst1("holiday")%></span></td>
		<td><span class="standard"><%=rst1("date")%></span></td>
	</tr>
	
	<%rst1.movenext
	loop%>
<%
else%>
<tr><td colspan="2"><span class="standard">There are no holidays setup for this region.</span></td></tr>
<%end if
rst1.close
%>
<tr><td colspan="2"><input type="button" value="Add Holiday" onclick="holidayEdit('');"><input type="button" value="Cancel" onclick="document.location='seasonView.asp?rid=<%=rid%>';"></td></tr>
</table>
</FORM>
</body>
</html>
