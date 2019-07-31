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
s=Request.QueryString("s")
e=Request.QueryString("e")
d=Request.QueryString("d")
portfolioid=request.querystring("portfolioid")
if isempty(i) then 
	i=100
end if


%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="102%" border="1" cellspacing="0" cellpadding="0" height="100%" align="center" bordercolor="#333333">
  <tr>
    <td>
      <div align="center"><img src=<%="PortfolioMakeChart.asp?d=" & d & "&s=" & s & "&e="&e&"&portfolioid="&portfolioid%>></div>
    </td>
  </tr>
</table>
<%
response.redirect "PortfolioMakeChart.asp?d=" & d & "&s=" & s & "&e="&e&"&portfolioid="&portfolioid
%>
</body>
</html>
