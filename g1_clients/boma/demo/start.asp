
<%@Language="VBScript"%>
<%
		if isempty(Session("loginemail")) then
			Response.Redirect "http://appserver1.genergy.com/eri_th/login.asp"	
		end if		
		
%>
<html>
<head>
<title>GenergyOne</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="235,*" rows="*" border="0" framespacing="0" frameborder="NO"> 
  <frame src="MyNavigator.htm" scrolling="NO" noresize frameborder="NO" name="nav">
  <frame  src="mymain.htm" noresize frameborder="NO" name="main">
</frameset>
<noframes>
</noframes> 
</html>
