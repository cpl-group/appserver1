<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim link, sid, target
sid = request.querystring("sid")
link = request.querystring("link")
target = request.querystring("target")
%>
<html>
<head><title>Change Option Details</title>
<script>
function send(link, target)
{	opener.document.forms['options'].link<%=sid%>.value = link;
	opener.document.forms['options'].target<%=sid%>.value = target;
	window.close();
}

</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
<form>
<table width="100%" border="2" cellspacing="0" cellpadding="3" bordercolor="#CCCCCC" style="font-family:arial, helvetica, sans=serif; font-size:12px">
<tr><td>Link</td>
	<td>Target</td></tr>
<tr><td><input type="text" name="link" value="<%=link%>"></td>
	<td><input type="text" name="target" value="<%=target%>"></td></tr>
</table>
<input type="button" value="Change" onclick="send(link.value, target.value)">
</form>
</body>
</html>
