<%@Language="VBScript"%>
<%
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		cnn1.Open application("CnnStr_security")
								
		strsql = "UPDATE employees SET status=0 where login = '" & Session("login") & "'"
		cnn1.execute strsql
		set cnn1 = nothing
		Session.Abandon
%>
<script>
setTimeout("self.close()",1000);
</script>
<link rel="Stylesheet" href="styles.css" type="text/css">
<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center">
  <tr>
    <td>
      <div align="center"><% Response.write "You have been successfully logged off" %></div>
      <script>opener.location = "http://www.genergyonline.com";</script>
    </td>
  </tr>
</table>