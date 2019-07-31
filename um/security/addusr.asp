<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'Response.Redirect "http://www.genergyonline.com"
		else
			if Session("eri") < 5 or Session("um") < 5 or Session("opslog") < 5 or Session("ts") < 5 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if
		
		end if		
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Add User</title>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>
<body bgcolor="#eeeeee">
<form name="form1" method="post" action="add.asp">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td colspan="2"><span class="standardheader">Add User</span></td>
</tr>
<tr valign="top">
  <td width="30%">
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td colspan="2"><b>User Info</b></td>
  </tr>
  <tr> 
    <td>Login</td>
    <td><input type="text" name="login" size="15"></td>
  </tr>
  <tr> 
    <td>Password</td>
    <td><input type="password" name="password" size="15"></td>
  </tr>
  <tr> 
    <td>Name</td>
    <td><input type="text" name="name" size="15" maxlength="25"></td>
  </tr>
  </table>
  </td>
  <td>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td colspan="2"><b>Security Levels</b></td>
  </tr>
  <tr> 
    <td>UM</td>
    <td><input type="text" name="um" size="15"></td>
  </tr>
  <tr> 
    <td>ERI</td>
    <td><input type="text" name="eri" size="15"></td>
  </tr>
  <tr>
    <td>OPSLog</td>
    <td><input type="text" name="opslog" size="15"></td>
  </tr>
  <tr> 
    <td>TS</td>
    <td><input type="text" name="ts" size="15"></td>
  </tr>
  <tr> 
    <td>CORP</td>
    <td><input type="text" name="corp" size="15"></td>
  </tr>
  <tr> 
    <td>ADMIN</td>
    <td><input type="text" name="admin" size="15"></td>
  </tr>
  </table>
  </td>
</tr>
<tr bgcolor="#dddddd">
  <td>&nbsp;&nbsp;<input type="submit" name="Submit" value="Submit" style="border:1px outset #ddffdd;background-color:ccf3cc;"></td>
  <td>&nbsp;</td>
</tr>
</table>
  
</form>
