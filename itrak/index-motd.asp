<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim username, password, message, loggedin, width, height
username = request("userid")
password = request("paswd")
message = "You are not currently logged in."
loggedin = false
width=800
height=600
if trim(username)<>"" then
	dim rst1, cnn1
	set cnn1 = server.createobject("ADODB.connection")
	set rst1 = server.createobject("ADODB.recordset")
	cnn1.open getconnect(0,0,"engineering")
	rst1.open "SELECT * FROM Users u LEFT JOIN clients c ON c.id=u.clientid WHERE u.userid='"&username&"' AND u.paswd='"&password&"'", cnn1
	if rst1.EOF then
		message = "Login failed, either userid or password were incorrect."
	else
		message = "Sucessful login."
		loggedin = true
		session("clientid") = trim(rst1("clientid"))
		session("superuser") = trim(rst1("superuser"))
		session("userid") = trim(rst1("userid"))
'		if not(isnull(rst1("window_width"))) then width=rst1("window_width")
'		if not(isnull(rst1("window_height"))) then height=rst1("window_height")
	end if
	rst1.close
end if
%>

<html>
<head>

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

.standard { font-family:Arial,Helvetica,sans-serif;font-size:8pt; }
.bottomline { border-bottom:1px solid #eeeeee; }
.floorlink { font-family:Arial,Helvetica,sans-serif;font-size:8pt; color:#0099ff; }
a.floorlink:hover { color:lightgreen; }
.shrunkenheader { font-family:Arial,Helvetica,sans-serif;font-size:7pt;font-weight:bold; }
a { color:#99ff99; }
a:visited { color:#99ccff; }

-->
</style>
<script language="javascript">
//<!--
loaded = 0;
function preloadImg(){
  btnLoginOn = new Image(); btnLoginOn.src = "images/login-1.gif";
  btnLoginOff = new Image(); btnLoginOff.src = "images/login.gif";
  ResetOn = new Image(); ResetOn.src = "images/reset-1.gif";
  ResetOff = new Image(); ResetOff.src = "images/reset.gif";
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
myheight = screen.availHeight;
//alert (screen.height + " " + screen.availHeight + " " + myheight);
<%
if loggedin then
	response.write "window.open('lightingframe.asp', 'Lighting','toolbar=no, menubar=no, location=no, width=' + mywidth + ', height=' + myheight);"
	response.write "window.close()"
end if
%>
function sizeandcenter(){
  desiredwidth = 600;
  desiredheight = 450;
  window.moveTo(((mywidth/2) - (desiredwidth/2)),((myheight/2) - (desiredheight/2))); 
  window.resizeTo(desiredwidth,desiredheight);
  window.focus();
}

</script>
<style type="text/css">
body { background-image:url(images/login-blue-2.gif); background-repeat:no-repeat; }
</style>
<body BGCOLOR="#336699" LINK="#0000CC" VLINK="#0000CC" TEXT="#000000" onload="sizeandcenter();document.forms['FrontPage_Form1'].userid.focus();preloadImg();">
<form method="POST" action="index.asp" name="FrontPage_Form1">
<table border=0 cellpadding="0" cellspacing="0" width="100%" height="100%">
<tr>
  <td valign="middle" align="center">
  <table border=0 cellpadding="4" cellspacing="0" width="468">
  <tr>
    <td width="240">
    <span class="standard" style="color:#ffffff;">
    <!-- messages go here -->
    <b>gEnergyOne Update</b><br><br>

    We've developed a new service. Now you can see how weather has been affecting your costs in heating and cooling. Click to learn more about the <a href="">Weather Analysis Report</a>.
    </span>
    </td>
    <td width="228">  
    <table border=0 cellpadding="0" cellspacing="0">
    <tr valign="top" height="7">
      <td width="9" height="7"><img src="images/cr_eee_blue1-nw.gif" alt="" width="10" height="8" border="0"></td>
      <td bgcolor="#eeeeee"><img src="images/treecontrols/spacer.gif" alt="" width="210" height="7" border="0"></td>
      <td width="9" align="right"><img src="images/cr_eee_blue2-ne.gif" alt="" width="10" height="8" border="0"></td>
    </tr>
    <tr> 
      <td align="center" colspan="3" bgcolor="#eeeeee">
      <!-- begin login fields table -->
      <table border=0 cellpadding="3" cellspacing="0" width="228">
      <tr>
        <td colspan="2" align="center"><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"><%=message%></span></td>
      </tr>
      <tr>
        <td align="right"><span class="standard">Username</span></td>
        <td><input type="text" name="userid" size="16" class="standard"></td>
      </tr>
      <tr>
        <td align="right"><span class="standard">Password</span></td>
        <td><input type="password" name="paswd" size="16" class="standard"></td>
      </tr>
      <tr>
        <td colspan="2" align="center">
        <input type="image" src="images/login.gif" name="btnLogin" value="Login" class="standard" onmouseover="msover(this);" onmouseout="msout(this);">
        <input type="image" src="images/reset.gif" name="Reset" value="Reset" class="standard" onmouseover="msover(this);" onmouseout="msout(this);">
        </td>
      </tr>
      </table>
      <!-- end login fields table -->
      </td>
    </tr>
    <tr valign="bottom" height="7">
      <td height="7"><img src="images/cr_eee_blue1-sw.gif" alt="" width="10" height="8" border="0"></td>
      <td bgcolor="#eeeeee"><img src="images/treecontrols/spacer.gif" alt="" width="186" height="7" border="0"></td>
      <td align="right"><img src="images/cr_eee-blue2-se.gif" alt="" width="10" height="8" border="0"></td>
    </tr>
    </table><br>
    </td>
  </tr>
  </table>
  <span class="standard" style="font-family:Arial,Helvetica,sans-serif;font-size:8pt;color:#ffffff""><b>NOTE:</b> gEnergyOne Rev. 1.1.2; Some advanced features require <a href="http://www.microsoft.com/windows/ie/downloads/ie6/default.asp" target="new" style="color:#99ccff">Internet Explorer +6.0</a>
  </td>
</tr>
</table>
</form>
</body>
</html>





