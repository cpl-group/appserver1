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
<title>Tenant Selection</title>
</head>

<body bgcolor="#FFFFFF">
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
strsql = "SELECT * FROM employees"
rst1.Open strsql, cnn1, 0, 1, 1

if not rst1.eof then
%>
<table border="1" width="100%" cellpadding="0" cellspacing="0" bordercolor="#000000" height="46" align="center">
  <tr> 
    <td width="18%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Login</font></td>
    <td width="35%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Password</font></td>
    <td width="35%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">User 
      Name </font></td>
    <td width="11%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">UM</font></td>
    <td width="8%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">ERI</font></td>
    <td width="13%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Ops 
      Log</font></td>
    <td width="4%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">TS</font></td>
    <td width="11%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" size="2">CORP</font></td>
    <td width="11%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" size="2">IT</font></td>
    <td width="11%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" size="2">ADMIN</font></td>
    <td width="11%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" size="2">STATUS</font></td>
  </tr>
  <tr valign="middle"> 
    <%
Do While Not rst1.EOF
 %>
    <form name="form1" method="post" action="updateusr.asp">
      <td width="16%" align="left" height="0"><font face="Arial"><i> 
        <input type="text" name="login" value="<%=rst1("login")%>" size="15">
        </i></font></td>
      <td width="16%" align="center" height="0"><font face="Arial"><i> 
        <input type="password" name="password" value="<%=rst1("password")%>" size="10">
        </i></font></td>
      <td width="16%" align="center" height="0"><font face="Arial"><i> 
        <input type="text" name="name" value="<%=rst1("name")%>" size="15">
        </i></font></td>
      <td width="11%" align="center" height="0"><font face="Arial"><i> 
        <input type="text" name="um" value="<%=rst1("um")%>" size="3">
        </i></font></td>
      <td width="11%" align="center" height="0"><font face="Arial"><i> 
        <input type="text" name="eri" value="<%=rst1("eri")%>" size="3">
        </i></font></td>
      <td width="11%" align="center" height="0"><font face="Arial"><i> 
        <input type="text" name="opslog" value="<%=rst1("opslog")%>" size="3">
        </i></font></td>
      <td width="11%" align="center" height="0"><font face="Arial"><i> 
        <input type="text" name="ts" value="<%=rst1("ts")%>" size="3">
        </i></font></td>
      <td width="11%" align="center" height="0"><font face="Arial"><i>
        <input type="text" name="corp" value="<%=rst1("corp")%>" size="3">
        </i></font></td>
      <td width="11%" align="center" height="0"><font face="Arial"><i> 
        <input type="text" name="it" value="<%=rst1("it")%>" size="3">
        </i></font></td>
      <td width="11%" align="center" height="0"><font face="Arial"><i> 
        <input type="text" name="admin" value="<%=rst1("admin")%>" size="3">
        </i></font></td>
      <% 
	  	if rst1("status")=1 then
		
		%>
      <td width="11%" align="center" height="0" bgcolor="#66FF00"></td>
      <%Else %>
      <td width="11%" align="center" height="0" bgcolor="#FF0000"></td>
      <%
	    
		end if %>
      <td width="11%" align="center" height="0" valign="bottom"> 
        <input type="hidden" name="id" value="<%=rst1("employeeid") %>">
        <input type="submit" name="Submit" value="Update">
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
</body>

</html>
