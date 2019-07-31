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
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr>
    <td bgcolor="#3399CC">
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Client 
        Administration</font></b></font></div>
    </td>
  </tr>
</table>
<IFRAME name="user" width="100%" height="215" src="usrdetail.asp" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
<IFRAME name="site" width="100%" height="500" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
</body>
</html>
