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
<script language="JavaScript" type="text/javascript">
function send(link, target)
{	opener.document.forms['options'].link<%=sid%>.value = link;
	opener.document.forms['options'].target<%=sid%>.value = target;
	window.close();
}

</script>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="#eeeeee">
<form>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td align="right">Link</td>
  <td><input type="text" name="link" value="<%=link%>"></td>
</tr>
<tr>
	<td align="right">Target</td>
	<td><input type="text" name="target" value="<%=target%>"></td>
</tr>
<tr>
  <td>&nbsp;</td>
  <td><input type="button" value="Change" onclick="send(link.value, target.value)" style="border:1px outset #ddffdd;background-color:ccf3cc;"></td>
</tr>
</table>

</form>
</body>
</html>
