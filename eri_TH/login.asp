<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<% 
message ="Please log in."

if trim(request("uname"))<> "" then 
	uname = Request("uname")
	pword = Request("pword")
	org = Request("org")
	logincomment = checkclientlogin(uname,pword, org)
	
	if logincomment = "1" then 
	%>
	<script>
	opener.location = "http://<%=request.servervariables("server_name")%><%=session("initialpath")%>"
	window.close()	
	 </script>
	<%
	else
	message="Login failed, system update ongoing, try again later."
	end if 
end if
%>
<html>
<head>

<title>gEnergyOne Login</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
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
background-repeat:no-repeat;
background-position:center;
}

.standard { font-family:Arial,Helvetica,sans-serif;font-size:8pt; }
.bottomline { border-bottom:1px solid #eeeeee; }
.floorlink { font-family:Arial,Helvetica,sans-serif;font-size:8pt; color:#0099ff; }
a.floorlink:hover { color:lightgreen; }
.shrunkenheader { font-family:Arial,Helvetica,sans-serif;font-size:7pt;font-weight:bold; }

-->
</style>
<script language="javascript">
//<!--
loaded = 0;
function preloadImg(){
  btnLoginOn = new Image(); btnLoginOn.src = "/images/login/login-1.gif";
  btnLoginOff = new Image(); btnLoginOff.src = "/images/login/login.gif";
  ResetOn = new Image(); ResetOn.src = "/images/login/reset-1.gif";
  ResetOff = new Image(); ResetOff.src = "/images/login/reset.gif";
  loaded = 1;
}

function msover(img){
  if (loaded) {
    img.src = eval(img.name + "On.src");
  }
}

function msout(img){
  if (loaded) {
    img.src = eval(img.name + "Off.src");
  }
}

//-->
</script>
<script>
mywidth = screen.availWidth - 8;
myheight = screen.availHeight - 28;
//alert (screen.height + " " + screen.availHeight + " " + myheight);
<%
if loggedin then
	response.write "opener = window.open('lightingframe.asp', 'Lighting','toolbar=no, menubar=no, location=no, width=' + mywidth + ', height=' + myheight);"
  response.write "opener.moveTo(0,0);"
	response.write "window.close()"
end if
%>
function sizeandcenter(){
  desiredwidth = 580;
  desiredheight = 430;
  window.moveTo(((mywidth/2) - (desiredwidth/2)),((myheight/2) - (desiredheight/2))); 
  window.resizeTo(desiredwidth,desiredheight);
}
</script>

<body BGCOLOR="#336699" background="images/myEnergyPlatformLogo.jpg" LINK="#0000CC" VLINK="#0000CC" TEXT="#000000" onload="document.forms['FrontPage_Form1'].uname.focus();preloadImg();">
<form method="POST" action="login.asp" name="FrontPage_Form1">
<table border=0 cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr>
  <td valign="middle" align="center">
  <!--[[img src="images/g_logo.jpg" width="209" height="72" border="0"]]--><br>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr valign="top" height="7">
    <td width="9" height="7"><img src="/images/login/cr_eee_blue1-nw.gif" alt="" width="10" height="8" border="0"></td>
    <td bgcolor="#eeeeee"><img src="/images/spacer.gif" alt="" width="210" height="7" border="0"></td>
    <td width="9" align="right"><img src="/images/login/cr_eee_blue2-ne.gif" alt="" width="10" height="8" border="0"></td>
  </tr>
  <tr> 
    <td align="center" colspan="3" bgcolor="#eeeeee">
    <table border=0 cellpadding="3" cellspacing="0" width="228">
    <tr>
      <td colspan="2" align="center"><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"><%=message%></span></td>
    </tr>
    <tr>
      <td align="right"><span class="standard">Username</span></td>
      <td><input type="text" name="uname" size="16" class="standard"></td>
    </tr>
    <tr>
      <td align="right"><span class="standard">Password</span></td>
      <td><input type="password" name="pword" size="16" class="standard">
                    <input name="org" type="hidden" id="group" value="genergy"> </td>
    </tr>
    <tr>
      <td colspan="2" align="center">
      <input type="image" src="/images/login/login.gif" name="btnLogin" value="Login" class="standard" onmouseover="msover(this);" onmouseout="msout(this);">
      <input type="image" src="/images/login/reset.gif" name="Reset" value="Reset" class="standard" onmouseover="msover(this);" onmouseout="msout(this);">
      </td>
    </tr>
    </table>
    </td>
  </tr>
  <tr valign="bottom" height="7">
    <td height="7"><img src="/images/login/cr_eee_blue1-sw.gif" alt="" width="10" height="8" border="0"></td>
    <td bgcolor="#eeeeee"><img src="/images/spacer.gif" alt="" width="186" height="7" border="0"></td>
    <td align="right"><img src="/images/login/cr_eee-blue2-se.gif" alt="" width="10" height="8" border="0"></td>
  </tr>
  </table>
        <p><font size="2" face="Arial, Helvetica, sans-serif"><br><span class="standard" style="font-family:Arial,Helvetica,sans-serif;font-size:8pt;color:#ffffff""> 
          <b>NOTE:</b> Some advanced features require <a href="http://www.microsoft.com/windows/ie/" target="new" style="color:#99ff99">Internet Explorer +6.0</a> </font></p>
        <p class="standard" style="font-family:Arial,Helvetica,sans-serif;font-size:8pt;color:#ffffff">Powered by myEnergyPlatform</p>
        <p class="standard" style="font-family:Arial,Helvetica,sans-serif;font-size:8pt;color:#ffffff">US Patent 7,069,161</p>
        <p class="standard" style="font-family:Arial,Helvetica,sans-serif;font-size:8pt;color:#ffffff">Copyright 2012 - All Rights Reserved</p></td>
</tr>
</table>
</form>
</body>
</html>




