 
<%@Language="VBScript"%>
<%
		if isempty(Session("loginemail")) then
			Response.Redirect "https://appserver1.genergy.com/eri_th/login.asp"	
		end if		
		
%>
<html>
<head>
<title>GenergyOne</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
var h;
var w;
function maxm()
{
w=screen.availWidth;
h=screen.availHeight;
//self.moveTo(0,0);
//self.window.resizeTo(w,h);
}
function logoff(){
	parent.opener.location.href="https://appserver1.genergy.com/eri_th/login.asp"
}

</script>
</head>
<frameset onload="maxm()" onunload="logoff()" cols="235,*" rows="*" border="0" framespacing="0" frameborder="NO"> 
  <frame src="g1navigation.asp" scrolling="NO" noresize frameborder="NO" name="nav">
  <frame  src="mymain.htm" noresize frameborder="NO" name="main">
</frameset>
<noframes> 
</noframes> 
</html>
