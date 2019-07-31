<%
logoUrl = request("logo")
if trim(logoUrl)="" then logoUrl = "logos/genergy2.gif"
%>
<html>
<head>
<title>logo</title>
</head>

<body>
<img src="<%=logoUrl%>" width="225" height="40" border="0">
</body>
</html>
