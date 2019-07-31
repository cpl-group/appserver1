<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim username, password, message, loggedin, width, height
username = request("userid")
password = request("paswd")
message = "You are currently not logged in."
loggedin = false
width=800
height=600
if trim(username)<>"" then
	dim rst1, cnn1
	set cnn1 = server.createobject("ADODB.connection")
	set rst1 = server.createobject("ADODB.recordset")
	cnn1.open application("cnnstr_lighting")
	rst1.open "SELECT * FROM Users u LEFT JOIN clients c ON c.id=u.clientid WHERE u.userid='"&username&"' AND u.paswd='"&password&"'", cnn1
	if rst1.EOF then
		message = "Login failed, either userid or password were incorrect."
	else
		message = "Sucessful login."
		loggedin = true
		session("clientid") = trim(rst1("clientid"))
		session("superuser") = trim(rst1("superuser"))
		session("userid") = trim(rst1("userid"))
		if not(isnull(rst1("window_width"))) then width=rst1("window_width")
		if not(isnull(rst1("window_height"))) then height=rst1("window_height")
	end if
	rst1.close
end if
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="../holiday/holiform.css">

<title>gEnergyOne Login</title>

</head>
<style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>
<script>
<%
if loggedin then
	response.write "window.open('lightingframe.asp', 'Lighting','toolbar=no, menubar=no, location=no, width="& width &", height="& height &"');"
	response.write "window.close()"
end if
%>
</script>

<body BGCOLOR="#FFFFFF" LINK="#0000CC" VLINK="#0000CC" TEXT="#000033" onload="document.forms['FrontPage_Form1'].userid.focus();">
<form method="POST" action="index.asp" name="FrontPage_Form1">
<table width="100%" border="0" height="100%" align="center">
<tr> 
      <td height="30"> 
        <div align="center"><img src="images/login_header.jpg" width="321" height="40"></div>
      </td>
    </tr>
    <tr> 
      <td height="137"> 
        <div align="center"> <font face="Arial, Helvetica, sans-serif" size="2"> 
          <%=message%>
          </font> 
          <p><font face="Arial, Helvetica, sans-serif">Username</font> 
            <input type="text" name="userid">
          </p>
          <p><font face="Arial, Helvetica, sans-serif">Password</font> 
            <input type="password" name="paswd">
          </p>
          <p> 
            <input type="submit" name="btnLogin" value="Login">
            <input type="reset" name="Reset" value="Reset">
          </p>
        </div>
      </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2"><span class="err"></span><span class="msg"></span>&nbsp;</font> 
          </b></div>
      </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><font color="#FF0000"><img src="images/login_client_footer.jpg" width="319" height="32"></font></b></font></div>
      </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><font color="#FF0000">NOTE: 
          </font><font face="Arial, Helvetica, sans-serif" size="1" color="#FF0000"><b>gEnergyOne 
          Rev. 1.1.2;</b></font><font color="#FF0000"> SOME ADVANCED FEATURES 
          REQUIRE INTERNET <a href="http://www.microsoft.com/windows/ie/downloads/ie6/default.asp" target="new">EXPLORER 
          +6.0</a></font></b></font> </div>
      </td>
    </tr>
    <tr>
      <td height="2">
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"></font></div>
      </td>
    </tr>
  </table>
  <div align="center"></div>
</form>
</body>
</html>





