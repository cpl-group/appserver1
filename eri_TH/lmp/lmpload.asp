<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<%
m=Request.QueryString("m")
d=Request.QueryString("d")
b=Request.QueryString("b")
s=Request.QueryString("s")
e=Request.QueryString("e")
i=Request.QueryString("i")
luid=Request.QueryString("luid")
lmp=request.querystring("lmp")
if isempty(i) then 
	i=100
end if


%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="102%" border="1" cellspacing="0" cellpadding="0" height="100%" align="center" bordercolor="#333333">
  <tr>
    <td>
      <div align="center"><img src=<%="makechartlmp.asp?m="& m & "&d=" & d & "&s=" & s & "&e="&e&"&b="&b&"&i="&i&"&luid="&luid&"&lmp="&lmp&"&nozoom="&request.querystring("nozoom")%>></div>
    </td>
  </tr>
</table>
<%
response.redirect "makechartlmp.asp?m="& m & "&d=" & d & "&s=" & s & "&e="&e&"&b="&b&"&i="&i&"&luid="&luid&"&lmp="&lmp&"&nozoom="&request.querystring("nozoom")
%>
</body>
</html>
