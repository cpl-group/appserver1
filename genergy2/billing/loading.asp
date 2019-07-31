<%Option Explicit
response.flush
%>
<html>
<head>
<title>Loading...</title>
</head>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
<body>
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td width="100%" height="100%" align="center" valign="middle">
<h1>Loading...</h1>
<img src="/genergyonev2/images/general/spinner.gif" alt="" border="0">
</td></tr></table>
</body>
</html>
<script>
document.location = "<%=request("url")%>";
</script>