<%@Language="VBScript"%>
<%
		'if isempty(Session("name")) then
		'	Response.Redirect "index.asp"
		'end if		
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
      <script language="javascript" src="./sniffer.js"></script>
	  <script language="javascript1.2" src="./custom.js"></script>
	  <script language="javascript1.2" src="./style.js"></script>
	  <script>
var c=0
var flag=0
//var h=document.body.clientHeight
function openpopup(flag){
	c++
//configure "Open Logout Window
//alert(c)

	//if(c==2){
		window.open("logout.asp","","width=300,height=338")
		window.location="http://www.nyelectric.com"
        //window.close()
	//}
	if (self.closed) {
		alert("ok")
	}
}

//function check(){
   	
	//if (window.closed) {
	//	alert("ok")
	//}else{
	    //alert("shit")
	//}
	//alert(document.body.clientWidth)
//}
function nullfuntion()
{	1=1
}
	  </script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onunload="">
<p> 
  <script language="javascript1.2" src="./menu.js"></script>
</p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;</p>
<p align="left">&nbsp;<font face="Arial, Helvetica, sans-serif">Welcome 
  <% Response.write Session("name")%>
  , please remember to logout when completed working. <br>&middot;<A href="#" onclick="window.open('http://10.0.7.110/test/GComControl.asp?u=<%=trim(Session("login"))%>', 'GCommunicator', 'scroll=no, width=200, height=101, toolbar=no');">GCommunicator</A></font></p>
<p align="left">&nbsp;</p>
<IFRAME style="border: solid blue 0px;" name="app" width="100%" height="70%" src="main.asp" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 

</body>
</html>
