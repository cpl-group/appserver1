<%@Language="VBScript"%>
<%
    if isempty(Session("name")) then
      Response.Redirect "http://www.genergyonline.com"
    else
      if Session("admin") < 5 then 
        Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

        Response.Redirect "../main.asp"
      end if  
    end if    
poid=Request.Querystring("poid")

%>

<html>
<head>
<title>Administrative Comment</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="processporeject.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#6699cc"><span class="standardheader">Administrative comment for PO #<%=Request.Querystring("ponum")%></span></td>
</tr>
<tr valign="top"> 
  <td><%=Request.Querystring("pocomment")%><br>&nbsp;</td>
</tr>
<tr valign="top"> 
  <td><input type="button" name="Submit2" value="Close Window" onclick="javascript:window.close()"></td>
</tr>
</table>
</form>
</body>
</html>
