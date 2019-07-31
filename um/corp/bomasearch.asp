<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>BOMA Search</title>
<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")

sqlstr = " Select * from tblIP"

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<form name="form1" method="post" action=""><table width="100%" border="0">
   
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#228866">
  <td><span class="standardheader">BOMA Search Results</span></td>
</tr>
</table>
<table border="0" cellpadding="3" cellspacing="1" bgcolor="#cccccc" width="100%">
<tr bgcolor="#dddddd" style="font-weight:bold;"> 
  <td width="34%">Email Account</td>
  <td width="33%">Number Of Visits To BOMA Site</td>
  <td width="33%">IP Address</td>
</tr>
<% While not rst1.EOF %>
<tr bgcolor="#ffffff"> 
  <td><%=rst1("email")%></a></td>
  <td><%=rst1("count")%></td>
  <td><%=rst1("ip")%></td>
</tr>
<%
rst1.movenext
Wend
%>
</table>
</form>
<%
rst1.close
%>
</body>
</html>