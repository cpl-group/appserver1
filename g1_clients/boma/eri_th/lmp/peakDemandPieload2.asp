<%option explicit

dim b, explode, coor, byear, bperiod, luid
b=Request.QueryString("b")
luid = request.querystring("luid")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
coor=Request.QueryString("coor")
explode=Request.QueryString("explode")
%>
<html>
<head>
<title></title>
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

<body bgcolor="#FFFFFF" text="#000000" onload="parent.closeLoadBox('loadFrame1');parent.shownav('peakCnav')">
<table width="102%" border="1" cellspacing="0" cellpadding="0" height="100%" align="center" bordercolor="#333333">
  <tr>
    <td>
      <div align="center"><%if trim(luid)="" then%><a href="peakDemandPieload2.asp?b=<%=b%>&luid=<%=luid%>&explode=<%=explode%>&byear=<%=byear%>&bperiod=<%=bperiod%>&coor="><%end if%><img border="0" src="peakDemandPie.asp?b=<%=b%>&luid=<%=luid%>&explode=<%=explode%>&byear=<%=byear%>&bperiod=<%=bperiod%>&coor=<%=coor%>" Ismap><%if trim(luid)="" then%></a><%end if%></div>
    </td>
  </tr>
</table>
<%
'response.redirect "peakDemandPie.asp?b="& b &"&luid="& luid &"&explode="& explode &"&byear="& byear &"&bperiod="& bperiod &"&coor="& coor 
%>
</body>
</html>
