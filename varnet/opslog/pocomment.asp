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
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="processporeject.asp">
  <table width="90%" border="0" align="center">
  <tr>
      <td bgcolor="#3399CC"><b><font face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">ADMINISTRATIVE 
        COMMENT FOR PO NUMBER <%=Request.Querystring("ponum")%></font></font></b></td>
  </tr>
</table>
  <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
    
    <tr valign="top"> 
      <td width="36%"><font face="Arial, Helvetica, sans-serif"><%=Request.Querystring("pocomment")%></font></td>
    </tr>
   
    <tr valign="top" bgcolor="#3399CC"> 
      <td width="36%">
        <div align="center"><font face="Arial, Helvetica, sans-serif"> 
          <input type="button" name="Submit2" value="Close Window" onclick="javascript:window.close()">
          </font></div>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
