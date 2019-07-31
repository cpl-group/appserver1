<%@Language="VBScript"%>
<%
		if isempty(Session("loginemail")) then
			Response.Redirect "http://appserver1.genergy.com/eri_th/login.asp"	
		end if		
		
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>

function openpopup(){
var popurl="start.asp"
winpops=window.open(popurl,"","width=1024,height=768,")
}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" onload="openpopup()">
<font face="Arial, Helvetica, sans-serif" size="3">User <%=Session("loginemail") %> 
successfully logged on </font> 
</body>
</html>
