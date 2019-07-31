<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Time Submitted</title>
<%
name=Request.Querystring("name")
start=Request.Querystring("start")	
end1=Request.Querystring("end1") 

		
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

	sqlstr = "exec sp_timesubmitted_email  '" & name & "','" &  start & "','" & end1 & "'"
		msg = "Time sheet for user "&name&" has been submitted. All parties are being notified via email."
	

cnn1.Execute sqlstr 
%>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF">

<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#6699cc"><span class="standardheader">Time Submission</span></td>
</tr>
<tr>
  <td><br>&nbsp;<%=msg%></td>
</tr>
</table>


</body>
</html>
