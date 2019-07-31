<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'Response.Redirect "../index.asp"
		else
			if  Session("ts") < 5 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if		
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Client Administration</span></td>
</tr>
</table>
<IFRAME name="user" width="100%" height="215" src="usrdetail.asp" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 
<IFRAME name="site" width="100%" height="500" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 
</body>
</html>
