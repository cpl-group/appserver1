<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim b, profiletype, leaseid, user, m, luid
b = request.querystring("b")
m = request.querystring("m")
luid = request.querystring("luid")
leaseid= Request("leaseid")
profiletype=Request("profiletype")
user=session("loginemail")
dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_security")
%>
<html>
<head>
<title></title>
<script>
</script>
</head><style type="text/css">
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

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000" onload="parent.closeLoadBox('loadFrame2');">
&nbsp;<br>
<table cellpadding="0" cellspacing="0" border="0" width="100%">
  <tr>
    <td width="51%" valign="top"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
      <%
'addons available
rst1.open "SELECT Label, Link, Target, Active FROM tbladdonlinks JOIN tbladdons on tbladdons.SID=tbladdonlinks.SID WHERE CSID=4 and userid='"&user&"' ORDER BY listorder", cnn1
if rst1.eof then response.write "Please contact support@genergy.com for access to the following options:"
do while not(rst1.eof)
    response.write "<a href="""&rst1("Link")&"?m="& m &"&b="& b &"&luid="& luid &""" style="""" onMouseOver=""this.style.color='gray'"" onMouseOut=""this.style.color='Black'"" onclick=""parent.openLoadBox('loadFrame2')"">"&rst1("Label")&"</a><br>"
    rst1.movenext
loop
rst1.close

'addons NOT available
rst1.open "SELECT Label FROM tbladdons WHERE SID not in (SELECT SID FROM tbladdonlinks WHERE userid='" &user& "' and active=1) AND CSID=4 ORDER BY listorder", cnn1
do while not(rst1.eof)
    response.write "<li style=""color:cccccc"">" &rst1("Label")& "</li>"
    rst1.movenext
loop

%>
      </b></font>
    <td width="49%"><font face="Arial, Helvetica, sans-serif" size="2"><b></b></font><font face="Arial, Helvetica, sans-serif" size="2"><b> 
      </b></font> </tr>
</table>

</body>
</html>