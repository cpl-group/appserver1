<%@Language="VBScript"%>
<%
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
								
		strsql = "UPDATE employees SET status=0 where login = '" & Session("login") & "'"
		cnn1.execute strsql
		set cnn1 = nothing
		Session.Abandon
%>
<script>
setTimeout("self.close()",1000);
</script>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center">
  <tr>
    <td>
      <div align="center"><% Response.write "User has been logged off" %></div>
    </td>
  </tr>
</table>
