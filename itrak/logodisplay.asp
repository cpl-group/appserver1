<%
dim cid
cid = request("cid")

dim cnn1, rst1
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getconnect(0,0,"engineering")

rst1.open "SELECT distinct logo FROM clients WHERE id="&cid, cnn1
if not rst1.eof then logo = rst1("logo")
rst1.close

if trim(logo)="" then logo = "logos/genergy2.gif"

cnn1.close

%>
<html>
<head>
<title>logo</title>
</head>

<body>
<img src="<%=logo%>" width="225" height="40" border="0">
</body>
</html>
