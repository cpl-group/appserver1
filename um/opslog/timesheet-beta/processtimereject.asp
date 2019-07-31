<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
user1=Request.Form("user")
message=request.form("message")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")

strsql="sp_reject_email '" & user1 & "','" & message & "'"

cnn1.execute strsql

set cnn1=nothing


%>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<body bgcolor="#FFFFFF">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td bgcolor="#666699"><span class="standardheader">Time Sheet Approval</span></td>
</tr>
<tr>
  <td><br><blockquote>Time sheet was rejected successfully</blockquote></td>
</tr>
</table>

</body>
</html>