<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
    if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
      'Response.Redirect "../index.asp"
    end if    
username=Request("username")
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->

<title>Tenant Selection</title>
<%
dim framesrc
if username<>"" then
  framesrc = "usrsite.asp?username=" & username
else
  framesrc = "null.htm"
end if
%>
<script language="JavaScript" type="text/javascript">
function fillup(name){
  document.location="usrdetail.asp?username="+name
  document.site.src="usrsite.asp?username="+name
  //alert(document.site.src);
}
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
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")


%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#666699">
  <td><span class="standardheader">Client Administration</span></td>
</tr>
</table>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
    <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">&nbsp;</td>
</tr>
<tr valign="top"> 
    <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> <br>
  </td>
  <form method="post" action="usrmodify.asp">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
 
  <table border=0 cellpadding="3" cellspacing="0"> 
  <%
  if username = "" then
  %>
  <tr>
    <td align="right">User</td>
    <td><input type="text" name="user" size="30"></td>
    <td align="right">Name</td>
    <td><input type="text" name="name" size="30"></td>
  </tr>
  <tr> 
    <td align="right">Passwd</td>
    <td><input type="text" name="passwd" size="30"></td>
    <td align="right">Email</td>
    <td><input type="text" name="email" size="30"></td>
  </tr>
  <tr> 
    <td align="right">Telephone</td>
    <td><input type="text" name="telephone" size="30"></td>
    <td align="right">Region Count</td>
    <td><input type="text" name="regioncount" size="30"></td>
  </tr>
  <tr>
    <td align="right">Company</td>
    <td><input type="text" name="company" size="30"></td>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr> 
    <td align="right">Initial Page</td>
    <td><input name="initial_page" type="text" value="http://appserver1.genergy.com/g1_clients/index.asp" size="46"></td>
    <td colspan="2"><input type="submit" name="choice" value="Save" style="border:1px outset #ddffdd;background-color:ccf3cc;"></td>
  </tr>
  <%
  else
      strsql = "SELECT * FROM clients where username ='"& username&"'"
    rst1.Open strsql, cnn1, 0, 1, 1
    if not rst1.eof then
  %>
     
  <input type="hidden" name="key" value="<%=rst1("clientkey")%>">
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td align="right">User </td>
    <td><input type="text" name="user" value="<%=username%>" size="30"></td>
    <td align="right">Name</td>
    <td><input type="text" name="name" value="<%=rst1("name")%>" size="30"> </td>
   </tr>
   <tr>
    <td align="right">Passwd</td>
    <td><input type="text" name="passwd" value="<%=rst1("paswd")%>" size="30"> </td>
    <td align="right">Email</td>
    <td><input type="text" name="email" value="<%=rst1("email")%>" size="30"> </td>
   </tr>
   <tr>
    <td align="right">Telephone</td>
    <td><input type="text" name="telephone" value="<%=rst1("telephone")%>" size="30"> </td>
    <td align="right">Region Count</td>
    <td><input type="text" name="regioncount" value="<%=rst1("regioncount")%>" size="30"> </td>
   </tr>
   <tr>
    <td align="right">Company</td>
    <td><input type="text" name="company" value="<%=rst1("company")%>" size="30"> </td>
    <td align="right">Initial Page</td>
    <td><input type="text" name="initial_page" value="<%=rst1("initial_page")%>" size="46"> </td>
   </tr>
   <tr>
    <td>&nbsp;</td>
    <td colspan="3"><input type="submit" name="choice" value="Update" style="border:1px outset #ddffdd;background-color:ccf3cc;"> <input type="submit" name="choice" value="Delete" style="border:1px outset #ddffdd;background-color:ccf3cc;"> <input type="button" name="choice" value="Add Building" onClick='document.site.location="usrsite.asp?flag=<%=username%>"' style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
   </tr>
   <%
   end if
  end if
  %>
  </table>
  </td>
  </form>
</tr>
</table>
<IFRAME name="site" width="100%" height="400" src="<%=framesrc%>" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 

</body>

</html>
