<%@LANGUAGE="VBSCRIPT" CODEPAGE="CP_ACP"%>
<html>
<head>
<title>Utility Manager Palm Sync</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body><center><h4>

<%
Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cnn.Open Application("cnnstr_palmserver")
cnn.commandTimeout = 600
Response.write("Executing Palm Sync....")
sql = "exec UMSYNC"
rs.Open sql, cnn
Response.write("<br><br>Palm sync complete.")

%>

</center></h4>
</body>
<script>window.close()</script>
</html>

