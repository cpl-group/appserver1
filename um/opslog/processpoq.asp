<html>
<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
poid=Request.form("poid")
user=Session("login")
message=request.form("message")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")

strsql="sp_po_question " & poid & ",'" & user & "','" & message & "'"

cnn1.execute(strsql)



set cnn1=nothing

%>
<head>
<title>Requisition Forms</title>
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td bgcolor="#666699"><span class="standardheader">Requisition Forms</span></td>
</tr>
</table>
<br>

<table border=0 cellpadding="3" cellspacing="0" width="90%" align="center">
<tr>
  <td>An email with your question has been sent out.</td>
</tr>
<tr>
  <td><br><input type="button" name="Button" value="Close Window" onclick="javascript:window.close()"></td>
</tr>
</table>

</body>
</html>