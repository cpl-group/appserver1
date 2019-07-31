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
			if  Session("ts") < 5 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if
		
		end if		
username=Request("username")
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script>
function fillup(name){
    document.location="usrdetail.asp?username="+name
	parent.frames.site.location="usrsite.asp?username="+name
}
</script>
<title>Client</title>
</head>

<body bgcolor="#FFFFFF">
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
strsql = "SELECT * FROM clients"
rst1.Open strsql, cnn1, 0, 1, 1


%>
<table border="1" width="100%" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" height="40" align="center">
<tr> 
    <td> 
      <div align="center">
        
    <select name="list" size="8" onClick=fillup(this.value)>
      <%
  do until rst1.eof
      if username = Trim(rst1("username")) then
  %>
      <option value="<%=rst1("username")%>" selected><%=rst1("username")%></option>
  <% 
      else
  %>
	      <option value="<%=rst1("username")%>"><%=rst1("username")%></option>
          <%
	  end if
  rst1.movenext
  loop
  rst1.close
  %>
        </select>
      </div>
	<input type="button" name="submit" value="Add Client" onclick='javascript:document.location="usrdetail.asp";parent.frames.site.location="null.htm"'>
    </td>
<td>
<div align="center"> 
<form method="post" action="usrmodify.asp">
<table width="465">
  <%
if username = "" then
%>
  <tr> 
    <td> 
      <div align="left"><font face="Arial, Helvetica, sans-serif" size="2">User</font></div>
    </td>
    <td><font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="user" size="30">
      </font></td>
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td><font face="Arial, Helvetica, sans-serif" size="2">Passwd</font></td>
    <td> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="passwd" size="30">
      </font></td>
    <td><font face="Arial, Helvetica, sans-serif" size="2">Name</font></td>
    <td> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="name" size="30">
      </font></td>
  </tr>
  <tr> 
    <td><font face="Arial, Helvetica, sans-serif" size="2">Telephone</font></td>
    <td> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="telephone" size="30">
      </font></td>
    <td><font face="Arial, Helvetica, sans-serif" size="2">Email</font></td>
    <td> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="email" size="30">
      </font></td>
  </tr>
  <tr> 
    <td><font face="Arial, Helvetica, sans-serif" size="2">Company</font></td>
    <td> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="company" size="30">
      </font></td>
    <td><font face="Arial, Helvetica, sans-serif" size="2">Region Count</font></td>
    <td> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="regioncount" size="30">
      </font></td>
  </tr>
  <tr> 
    <td><font face="Arial, Helvetica, sans-serif" size="2">Initial Page</font></td>
    <td> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="initial_page" size="46">
      </font></td>
    <td colspan="2">
      <div align="center"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="submit" name="choice" value="Save">
        </font></div>
    </td>
  </tr>
  <%
else
    strsql = "SELECT * FROM clients where username ='"& username&"'"
	rst1.Open strsql, cnn1, 0, 1, 1
%>
  <div align="center"> 
    <input type="hidden" name="key" value="<%=rst1("clientkey")%>">
    <table width="387">
      <tr> 
        <td> 
          <div align="left">User</div>
        </td>
        <td> 
          <input type="text" name="user" value="<%=username%>" size="30">
        </td>
        <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="2"><a href="revtrack.asp?userid=<%=username%>&pid=<%=rst1("portfolio_id")%>" target="_blank">Revisions</a></font></td>
      </tr>
      <tr> 
        <td>Passwd</td>
        <td> 
          <input type="text" name="passwd" value="<%=rst1("paswd")%>" size="30">
        </td>
        <td>Name</td>
        <td> 
          <input type="text" name="name" value="<%=rst1("name")%>" size="30">
        </td>
      </tr>
      <tr> 
        <td>Telephone</td>
        <td> 
          <input type="text" name="telephone" value="<%=rst1("telephone")%>" size="30">
        </td>
        <td>Email</td>
        <td> 
          <input type="text" name="email" value="<%=rst1("email")%>" size="30">
        </td>
      </tr>
      <tr> 
        <td>Company</td>
        <td> 
          <input type="text" name="company" value="<%=rst1("company")%>" size="30">
        </td>
        <td>Region Count</td>
        <td> 
          <input type="text" name="regioncount" value="<%=rst1("regioncount")%>" size="30">
        </td>
      </tr>
      <tr> 
        <td>Initial Page</td>
        <td> 
          <input type="text" name="initial_page" value="<%=rst1("initial_page")%>" size="46">
        </td>
        <td colspan="2"> 
          <div align="center"> 
            <input type="submit" name="choice" value="Update">
            <input type="submit" name="choice" value="Delete">
            <input type="button" name="choice" value="Add Building" onClick='javascript:parent.frames.site.location="usrsite.asp?flag=<%=username%>"'>
          </div>
        </td>
      </tr>
      <%
end if
%>
    </table></form>
  </div></td></tr>
</table>

</body>

</html>
