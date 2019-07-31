<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "index.asp"
		end if		
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript" src="./sniffer.js"></script>
<script language="javascript1.2" src="./custom.js"></script>
<script language="javascript1.2" src="./style.js"></script>
<script>
window.name="init"
function openpopup(user){
//configure "Open Logout Window
//window.open("logout.asp","","width=300,height=338")
user=document.forms[0].user.value
//alert(self.location.href)
alert(document.getSelection())
<%
  if isempty(Session("name")) then
%>
  window.open("logout.asp","","width=300,height=338")
<%
  else
%>
  alert(user)
<%
  end if
%>

}
function loadpopup(){
openpopup()
}
</script>
</head>

<%
user=Session("login")
%>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" 
marginheight="0" onunload="openpopup()">
<p>
  <script language="javascript1.2" src="./menu.js"></script>
</p>
<p>&nbsp;</p>
<form name="form1" method="post" action="">
<input type="hidden" name="user" value="<%=user%>">
</form>
<table width="100%" border="0" height="100%">
  <tr> 
    <td valign="top" height="28"> 
      <div align="right"><font face="Arial, Helvetica, sans-serif">Welcome <% Response.write Session("name")%></font></div>
    </td>
  </tr>
  <tr> 
    <td valign="top">
      <div align="center"><font face="Arial, Helvetica, sans-serif"><i>&nbsp; 
        <% Response.Write Session("fmessage") %>
        </i></font></div>
    </td>
  </tr>
</table>
</body>
</html>
