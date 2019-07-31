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
</head>
<body bgcolor="#FFFFFF">
<form name="form1" method="post" action="add.asp">
  <div align="center">
    <p><font face="Arial, Helvetica, sans-serif"><b>ADD USER</b></font></p>
    <table width="31%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
      <tr> 
        <td width="38%" bgcolor="#66CCFF" > 
          <div align="left"><font face="Arial, Helvetica, sans-serif">Login</font></div>
        </td>
        <td width="62%" > 
          <div align="center"> 
            <input type="text" name="login" size="15">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%" bgcolor="#66CCFF"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif">Password</font></div>
        </td>
        <td width="62%"> 
          <div align="center"> 
            <input type="password" name="password" size="15">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%" bgcolor="#66CCFF"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif">Name</font></div>
        </td>
        <td width="62%"> 
          <div align="center"> 
            <input type="text" name="name" size="15" maxlength="25">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%" bgcolor="#66CCFF"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif">UM</font></div>
        </td>
        <td width="62%"> 
          <div align="center"> 
            <input type="text" name="um" size="15">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%" bgcolor="#66CCFF" height="37"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif">ERI</font></div>
        </td>
        <td width="62%" height="37"> 
          <div align="center"> 
            <input type="text" name="eri" size="15">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%" bgcolor="#66CCFF" height="34"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif">OPSLog</font></div>
        </td>
        <td width="62%" height="34"> 
          <div align="center"> 
            <input type="text" name="opslog" size="15">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%" bgcolor="#66CCFF"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif">TS</font></div>
        </td>
        <td width="62%"> 
          <div align="center"> 
            <input type="text" name="ts" size="15">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif">CORP</font></td>
        <td width="62%"> 
          <div align="center"> 
            <input type="text" name="corp" size="15">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif">ADMIN</font></td>
        <td width="62%"> 
          <div align="center"> 
            <input type="text" name="admin" size="15">
          </div>
        </td>
      </tr>
      <tr> 
        <td width="38%"> 
          <div align="left"> <font face="Arial, Helvetica, sans-serif"> 
            <input type="submit" name="Submit" value="Submit">
            </font></div>
        </td>
        <td width="62%"> 
          <div align="center"></div>
        </td>
      </tr>
    </table>
  </div>
</form>
