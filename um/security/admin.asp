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
<base target="cframe">
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>
<body bgcolor="#eeeeee" text="#000000">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td bgcolor="#666699"><span class="standardheader">Employee Administration</span></td>
  </tr>
  <tr> 
    <td valign="top" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
    <img src="/images/intranet/aro-rt.gif" align="absmiddle" border="0">&nbsp;<a href="usrlist.asp" target="cframe">User List</a> | <img src="/images/intranet/aro-rt.gif" align="absmiddle" border="0">&nbsp;<a href="addusr.asp" target="cframe">Add User</a></p>
    </td>
  </tr>
</table>
<IFRAME name="cframe" width="100%" height="86%" src="usrlist.asp" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 

</body>
</html>
