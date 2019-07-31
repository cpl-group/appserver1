<%@Language="VBScript"%>
<%
		if isempty(Session("loginemail")) then
			Response.Redirect "https://appserver1.genergy.com/eri_th/login.asp"	
		end if		
		
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>

function openpopup(){
var popurl="g1navigation.asp"
winpops=window.open("https://appserver1.genergy.com/g1_clients/g1navigation.asp","GenergyOne","resizable,status")
}

</script>


</head>

<body bgcolor="#FFFFFF" text="#000000" onload="openpopup()">
<div align="center"><font face="Arial, Helvetica, sans-serif" size="3">User <%=Session("loginemail") %> 
  successfully logged on </font> </div>
</body>
</html>
