<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'3/4/2008 N.Ambo added functionality to grey out portfolios which are offline

if not(allowGroups("Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, ccolor
pid = secureRequest("pid")
bldg = secureRequest("bldg")
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,bldg,"dbCore")


tid = secureRequest("tid")

dim utility, lid
if trim(lid)<>"" then
	rst1.Open "SELECT * FROM tblleasesutilityprices WHERE leaseutilityid='"&lid&"'", cnn1
	if not rst1.EOF then
		utility = rst1("utility")
	end if
	rst1.close
end if

%>
<html>
<head>
<title>Portfolio View</title>
<script>
function loadPortfolio(pid)
{	document.location = "toolbar.asp?pid="+pid;
//parent.contentfrm.location = "groupView.asp?pid="+pid;
}
function portfolioEdit(pid){
  document.location = "portfolioedit.asp?pid="+pid;
}
function groupView()
{	parent.contentfrm.location = 'groupView.asp?pid=<%=pid%>';
}

function loadbottomframes(){
  if (parent.contentfrm.length < 1) {
    parent.contentfrm.location = "contentfrm.asp";
  }
}

</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form>
<table border=0 cellpadding="5" cellspacing="0" width="100%">
<tr>
  <td bgcolor="#000000">
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
</table>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr><td colspan="3" bgcolor="#3399cc"><span class="standardheader">Portfolio Setup</span></td></tr>
<tr bgcolor="#dddddd">
  <td width="5%"><span class="standard"><b>id</b></span></td>
  <td width="15%"><span class="standard"><b>Portfolio Number</b></span></td>
  <td><span class="standard"><b>Portfolio Name</b></span></td>
</tr>
<%
rst1.open "SELECT * FROM portfolio WHERE id != 150 ORDER BY name", cnn1
do until rst1.eof
	ccolor = ""
	if rst1("offline") then ccolor="class=""grayout"""	
  %>
  <tr onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="portfolioEdit(<%=rst1("id")%>);">
     <td <%=ccolor%>> <%=rst1("id")%></td>
    <td <%=ccolor%>> <%=rst1("portfolio")%></td>
    <td <%=ccolor%>> <%=rst1("name")%></td>
  </tr>
  <%
  rst1.MoveNext
loop
rst1.close
%>
<tr><td colspan="3" height="10"></td></tr>
<tr>
  <td colspan="3" bgcolor="#dddddd"><input type="button" value="Add Portfolio" onclick="portfolioEdit('');" id=1 name=1 class="standard"></td>
</tr>
  
</table>

</form>
</body>
</html>
