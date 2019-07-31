<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, level, isportfolio
pid = secureRequest("pid")
tid = secureRequest("tid")
bldg = secureRequest("bldg")
if bldg="0" then bldg = ""
if trim(bldg)="" then isportfolio = 1 else isportfolio = 0

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
if trim(bldg)<>"" then
  cnn1.open getLocalConnect(bldg) 
else 
  cnn1.open getMainConnect(pid)
end if


'where build
dim where, bldgname, portfolioname, breadcrumbtrail
if trim(bldg)<>"" then 
  where = where&" and g.bldgnum='"&trim(bldg)&"'"

	rst1.Open "SELECT * FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.eof then
    bldgname = rst1("bldgname")
  	portfolioname = rst1("name")
  end if
	rst1.close
	breadcrumbtrail = "<a href='portfolioedit.asp?pid=" & pid & "' style='color:#ffffff;'>" & portfolioname & "</a> &gt; <a href='buildingedit.asp?pid=" & pid & "&bldg=" & bldg & "' style='color:#ffffff'>" & bldgname & "</a>"
else
  rst1.open "select name from portfolio where id='" & pid & "'", cnn1
  portfolioname = rst1("name")
	breadcrumbtrail = "<a href=""portfolioedit.asp?pid=" & pid & """ style=""color:#ffffff;"">" & portfolioname & "</a>"
  rst1.close
end if
if trim(tid)<>"" and trim(tid)<>"0" then where = where&" and g.tenant='"&trim(tid)&"'"


%>
<html>
<head>
<title>Portfolio View</title>
<script>
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}
function groupEdit(groupname)
{	document.location = 'groupEdit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&groupname='+groupname;
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<body>
<FORM>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
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
  <td>
  <span class="standardheader">
  Manage Groups | <span style="font-weight:normal;"><%=breadcrumbtrail%></span>
  </td>
   <td align="right"><button id="qmark2" onclick="openCustomWin('help.asp?page=groupview','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) Quick Help</button></td>
</tr>
<tr bgcolor="#eeeeee">
  <td><b>Groups</b></td>
  <td align="right"><%if not(isBuildingOff(bldg)) then%><input type="button" value="Add group" onclick="groupEdit('');" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><%end if%>&nbsp;</td>
</tr>
</table>

<%
DIM SQLVAR

sqlvar = "SELECT * FROM [group] g INNER JOIN grouptype gt ON gt.id=g.type WHERE groupname in (SELECT [name] FROM sysobjects) and clientid='"&pid&"' and portfolio="&isportfolio&" "&where&" ORDER BY grouplabel desc"
rst1.Open sqlvar, cnn1

if not rst1.EOF then%>
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
	<tr bgcolor="#dddddd">
		<td width="35%"><span class="standard"><b>Group Name</b></span></td>
		<td width="65%"><span class="standard"><b>Type</b></span></td>
	</tr>

	<%do until rst1.EOF%>
	<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="groupEdit('<%=rst1("groupname")%>');">
		<td><span class="standard"><%=rst1("grouplabel")%></span></td>
		<td><span class="standard"><%=rst1("type")%></span></td>
	</tr>
	
	<%rst1.movenext
	loop%>
	<tr><td bgcolor="#dddddd" colspan="2"><%if not(isBuildingOff(bldg)) then%><input type="button" value="Add group" onclick="groupEdit('');" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><%end if%>&nbsp;</td></tr>
</table>
<%
else
%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr><td><span class="standard">There are no groups set up for this section.</span></td></tr>
<tr><td bgcolor="#dddddd"><%if not(isBuildingOff(bldg)) then%><input type="button" value="Add group" onclick="groupEdit('');" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><%end if%>&nbsp;</td></tr>
</table>
<%
end if
rst1.close
%>
</FORM>
</body>
</html>
