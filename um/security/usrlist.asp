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
<title>Genergy Users</title>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="#eeeeee">
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")
strsql = "SELECT * FROM employees"
rst1.Open strsql, cnn1, 0, 1, 1

dim scolor

if not rst1.eof then
%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#dddddd" style="font-weight:bold;"> 
    <td style="border-bottom:1px solid #ffffff;">Status</td>
    <td style="border-bottom:1px solid #ffffff;">Login</td>
    <td style="border-bottom:1px solid #ffffff;">Password</td>
    <td style="border-bottom:1px solid #ffffff;">User Name</td>
    <td style="border-bottom:1px solid #ffffff;">UM</td>
    <td style="border-bottom:1px solid #ffffff;">ERI</td>
    <td style="border-bottom:1px solid #ffffff;">Ops Log</td>
    <td style="border-bottom:1px solid #ffffff;">TS</td>
    <td style="border-bottom:1px solid #ffffff;">CORP</td>
    <td style="border-bottom:1px solid #ffffff;">IT</td>
    <td style="border-bottom:1px solid #ffffff;">ADMIN</td>
    <td style="border-bottom:1px solid #ffffff;">&nbsp;</td>
  </tr>
  <tr valign="middle"> 
    <%
Do While Not rst1.EOF
 %>
    <form name="form1" method="post" action="updateusr.asp">
      <% 
	  	if rst1("status")=1 then
		    scolor = "#66ee66"
		  else
		    scolor = "#ff3300"
		  end if
  		%>
      <td align="center"><div style="width:16px;height:16px;background-color:<%=scolor%>;">&nbsp;</div></td>
      <td><input type="text" name="login" value="<%=rst1("login")%>" size="15"></td>
      <td><input type="password" name="password" value="<%=rst1("password")%>" size="10"></td>
      <td><input type="text" name="name" value="<%=rst1("name")%>" size="15"></td>
      <td><input type="text" name="um" value="<%=rst1("um")%>" size="3"></td>
      <td><input type="text" name="eri" value="<%=rst1("eri")%>" size="3"></td>
      <td><input type="text" name="opslog" value="<%=rst1("opslog")%>" size="3"></td>
      <td><input type="text" name="ts" value="<%=rst1("ts")%>" size="3"></td>
      <td><input type="text" name="corp" value="<%=rst1("corp")%>" size="3"></td>
      <td><input type="text" name="it" value="<%=rst1("it")%>" size="3"></td>
      <td><input type="text" name="admin" value="<%=rst1("admin")%>" size="3"></td>
      <td valign="bottom"> 
        <input type="hidden" name="id" value="<%=rst1("employeeid") %>">
        <input type="submit" name="Submit" value="Update" style="border:1px outset #ddffdd;background-color:ccf3cc;">
		</td>
    </form>
  </tr>
  <%
rst1.MoveNext  
Loop

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing


End if
 %>
</table>
<br><br>

</body>
</html>
