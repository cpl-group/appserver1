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
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function openpopup(){
//configure "Open Logout Window

parent.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}
</script>
<STYLE>
<!--
A.ssmItems:link		{color:black;text-decoration:none;}
A.ssmItems:hover	{color:black;text-decoration:none;}
A.ssmItems:active	{color:black;text-decoration:none;}
A.ssmItems:visited	{color:black;text-decoration:none;}
//-->
</STYLE>
<base target="main">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" height="100%" align="center">
  <tr>
    <td valign="bottom" height="2" bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Employee 
        Administration</b></font></div>
    </td>
  </tr>
  <tr> 
    <td valign="top"> 
      <p><a href="usrlist.asp" target="main">User List</a> | <a href="addusr.asp" target="main">Add 
        User</a></p>
      <p>&nbsp;</p><IFRAME name="main" width="100%" height="500" src="usrlist.asp" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME> 
    </td>
  </tr>
</table>
<p>&nbsp;</p>

<p>&nbsp; </p>
</body>
</html>
