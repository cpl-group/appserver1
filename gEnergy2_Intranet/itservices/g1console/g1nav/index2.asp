<%@Language="VBScript"%>
<%

		if isempty(Session("loginemail")) then
			Response.Redirect "http://www.genergyonline.com"	
		end if		
		
		nocache=rnd*1000000
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

    winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',status=yes,scrollbars='+scroll+',resizable=yes'

     // open new window and use the variables to position it
	popupwin=window.open("http://appserver1.genergy.com/g1_clients/g1nav.asp?nfc=<%=clng(nocache)%>","ClientProfiles",winprops)
	popupwin.focus('ClientProfiles')
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
