<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg
pid = secureRequest("pid")
bldg = secureRequest("bldg")
'dim DBmainmodIP

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
if trim(bldg)<>"" then
  cnn1.open getLocalConnect(bldg) 
 ' DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."
else 
  cnn1.open getMainConnect(pid)
end if

dim bldgname, portfolioname, breadcrumbtrail
if trim(bldg)<>"" then 
	rst1.Open "SELECT * FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	bldgname = rst1("bldgname")
	portfolioname = rst1("name")
	rst1.close
	breadcrumbtrail = "<a href='portfolioeditG1.asp?pid=" & pid & "' style='color:#ffffff;'>" & portfolioname & "</a> &gt; <a href='buildingeditG1.asp?pid=" & pid & "&bldg=" & bldg & "' style='color:#ffffff'>" & bldgname & "</a>"
else
  rst1.open "select name from portfolio where id='" & pid & "'", cnn1
  portfolioname = rst1("name")
	breadcrumbtrail = "<a href=""portfolioeditG1.asp?pid=" & pid & """ style=""color:#ffffff;"">" & portfolioname & "</a>"
  rst1.close
end if
%>
<html>
<head>
<title>Contact View</title>
<script>
function contactEdit(ctid)
{	document.location = 'contacteditG1.asp?pid=<%=pid%>&bldg=<%=bldg%>&ctid='+ctid
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
	Contacts | <span style="font-weight:normal;"><%=breadcrumbtrail%></span>
	</span></td>
</tr>
<%
rst1.Open "SELECT * FROM contacts WHERE cid='"&pid&"' AND bldgnum='"&bldg&"'", cnn1
if not rst1.EOF then%>
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
	<tr bgcolor="#cccccc">
		<td><span class="standard"><b>Contact Name</b></span></td>
		<td><span class="standard"><b>Address</b></span></td>
		<td><span class="standard"><b>Phone</b></span></td>
		<td><span class="standard"><b>Fax</b></span></td>
		<td><span class="standard"><b>Email</b></span></td>
	</tr>

	<%do until rst1.EOF%>
	<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="contactEdit('<%=rst1("id")%>');">
		<td><span class="standard"><%=rst1("Name")%></span></td>
		<td><span class="standard"><%=rst1("address")%>, <%=rst1("city")%>&nbsp;<%=rst1("state")%>, <%=rst1("zip")%> </span></td>
		<td><span class="standard"><%=rst1("phone")%></span></td>
		<td><span class="standard"><%=rst1("fax")%></span></td>
		<td><span class="standard"><%=rst1("email")%></span></td>
	</tr>
	
	<%rst1.movenext
	loop%>
</table>
<%
else
%>
<table border=0 cellpadding="3" cellspacing="0">
<tr>
  <td>There are no contacts set up.</td>
</tr>
</table>
<%
end if
rst1.close
%>
<table border=0 cellpadding="3" cellspacing="0">
<tr>
  <td><%if not(isBuildingOff(bldg)) then%><input type="button" value="Add Contact" onclick="contactEdit('');" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;margin:3px;"><%end if%>&nbsp;</td>
</tr>
</table>
</FORM>
</body>
</html>
