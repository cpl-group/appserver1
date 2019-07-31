<%option explicit%>
<%
dim  b, date1, date2, utype
b = request.querystring("b")
utype = request.querystring("utype")
date1 = request.querystring("date1")
date2 = request.querystring("date2")


%>
<html>
<head>
<title>Untitled Document</title>
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

<body bgcolor="#FFFFFF" text="#000000" onload="parent.closeLoadBox('loadFrame1');">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center" bordercolor="#333333">
  <tr>
    <td>
      <div align="center"><img src="<%="MakeChart.asp?b="& b &"&utype="& utype &"&date1="& date1 &"&date2="& date2%>"></div>
    </td>
  </tr>
</table>
<%
'response.redirect "MakeChart.asp?b="& b &"&utype="& utype &"&date1="& date1 &"&date2="& date2
%>
</body>
</html>
