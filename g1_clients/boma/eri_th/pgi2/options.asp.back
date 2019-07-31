<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, profiletype, user, meterid, utility
bldg = request.querystring("bldg")
meterid = request.querystring("meterid")
profiletype=Request("profiletype")
utility=request("utility")
user=session("loginemail")
dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")
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
Dim sql
sql = "SELECT Label, Link, Target, Active, tbladdonlinks.sid FROM tbladdonlinks JOIN tbladdons on tbladdons.SID=tbladdonlinks.SID WHERE CSID=4 and userid='"&user&"' ORDER BY listorder"
rst1.open sql, cnn1
if rst1.eof then response.write "Please contact support@genergy.com for access to the following options:"
do while not(rst1.eof)
    %><a href="javascript:document.location='<%=rst1("Link")%>?meterid=<%=meterid%>&bldg=<%=bldg%>&billingid='+parent.document.forms[0].billingid.value+'&utility=<%=utility%>&startdate='+parent.document.forms[0].startdate.value" style="" onMouseOver="this.style.color='gray'" onMouseOut="this.style.color='Black'" onclick="parent.openLoadBox('loadFrame2')"><%=rst1("Label")%></a><br><%
    rst1.movenext
loop
rst1.close

'addons NOT available
'To add links see Stored Porc in dbCore called usp_Edit_Client_Options. -- Anthony Corriero
rst1.open "SELECT Label FROM tbladdons WHERE SID not in (SELECT SID FROM tbladdonlinks WHERE userid='" &user& "' and active=1) AND CSID=4 ORDER BY listorder", cnn1
do while not(rst1.eof)
    response.write "<li style=""color:cccccc"">" &rst1("Label")& "</li>"
    rst1.movenext
loop

%>
      </b></font>
    <td width="49%"> 
      <p>&nbsp;</p>
      <font face="Arial, Helvetica, sans-serif" size="2"><b></b></font><font face="Arial, Helvetica, sans-serif" size="2"><b> 
      </b></font>
  </tr>
</table>

</body>
</html>