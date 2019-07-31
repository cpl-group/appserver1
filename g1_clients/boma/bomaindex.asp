<%@Language="VBScript"%>
<%
		if isempty(Session("loginemail")) then
			Response.Redirect "http://www.genergyonline.com"	
		end if		
		
%>
<html>
<head>
<title>Login Processing...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>

function openpopup(){
     // read resolution and set two variables
	var w = 1024
	var h = 768
	var winl = (screen.width - w) / 2;
    var wint = (screen.height - h) / 2;

    winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',status=yes,scrollbars='+scroll+',resizable=no'

     // open new window and use the variables to position it
	popupwin=window.open("bomag1navigation.asp","GenergyOne",winprops)
	popupwin.focus('GenergyOne')
}
function closeme(){ window.close() }
closeme()
</script>


</head>

<body bgcolor="#FFFFFF" text="#000000" onload="openpopup()">
<div align="center"><font face="Arial, Helvetica, sans-serif" size="3">User <%=Session("loginemail") %> 
  successfully logged on </font> </div>
</body>
</html>
