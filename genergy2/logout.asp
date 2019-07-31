<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<%
		Dim cnn, rs,sql
		set cnn = server.createobject("ADODB.Connection")
		set rs = server.createobject("ADODB.Recordset")
		cnn.open getConnect(0,0,"dbCore")
		
		sql = "delete from logintracking where username='"&trim(getKeyValue("user"))&"';insert into logintracking_LT (username,pageview) values ('"&trim(getKeyValue("user"))&"','logout')"
		cnn.execute sql		
%>

<head>
<title>Log Out</title>
<script language="JavaScript" type="text/javascript">
function popUp(){
  popper = window.open("logoutpopup.asp","LogOut","width=400,height=300,scrollbars=0,status=0,resizeable=0")
  //window.name = "opener";
  setTimeout("backToGenergyHome();", 2000);
}
function backToGenergyHome()
{
    parent.window.location="http://www.genergyonline.com"

}
</script>
<link rel="Stylesheet" href="styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" onload="popUp();">
&nbsp;
</body>
</html>