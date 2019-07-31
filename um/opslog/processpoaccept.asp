<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/gEnergy2/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" onload="opener.window.location.reload();">
<%
poid=Request("poid")
user=Session("login")
message=replace(request("message"),"'","")
acctponum = request("acctponum")

if request("action") = "accept" then
	Set cnnMain = Server.CreateObject("ADODB.Connection")
	cnnMain.Open getConnect(0,0,"intranet")
	
	strsql="sp_po_accept " & poid & ",'" & getxmlusername() & "','" & message & "'"
	
	cnnMain.execute(strsql)
	set cnnMain = nothing
	%>
	<table border=0 cellpadding="3" cellspacing="0" width="100%">
		<tr>
			<td bgcolor="#666699"><span class="standardheader">Requisition Form Acceptance</span></td>
		</tr>
	</table>
	<br>
	<table border=0 cellpadding="3" cellspacing="0" width="90%" align="center">
		<tr>
			<td>The requisition form has been accepted.</td>
		</tr>
		<tr>
			<td><br><input type="button" name="Button" value="Close Window" onclick="javascript:opener.window.location.reload();window.close()"></td>
		</tr>
	</table>
	<%
elseif request("action") = "approve" then
	Set cnnMain = Server.CreateObject("ADODB.Connection")
	cnnMain.Open getConnect(0,0,"intranet")
	
	strsql="sp_po_approve " & poid & ",'" & user & "','" & message & "'"
	
	cnnMain.execute(strsql)
	set cnnMain = nothing
	%>
	<table border=0 cellpadding="3" cellspacing="0" width="100%">
		<tr>
			<td bgcolor="#666699"><span class="standardheader">Requisition Form Approval</span></td>
		</tr>
	</table>
	<br>
	<table border=0 cellpadding="3" cellspacing="0" width="90%" align="center">
		<tr>
			<td>The requisition form has been approved.</td>
		</tr>
		<tr>
			<td><br><input type="button" name="Button" value="Close Window" onclick="javascript:opener.window.location.reload();window.close()"></td>
		</tr>
	</table>
	<%
end if		%>
</body>
</html>