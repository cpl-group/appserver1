<%option explicit

dim bldg, explode, coor, byear, bperiod, billingid, utility
bldg=Request.QueryString("bldg")
billingid = request.querystring("billingid")
utility = request.querystring("utility")
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
      <div align="center"><%if trim(billingid)="" then%><a href="peakDemandPieload.asp?bldg=<%=bldg%>&billingid=<%=billingid%>&explode=<%=explode%>&byear=<%=byear%>&bperiod=<%=bperiod%>&utility=<%=utility%>&coor="><%end if%><img border="0" src="peakDemandPie.asp?bldg=<%=bldg%>&billingid=<%=billingid%>&explode=<%=explode%>&byear=<%=byear%>&bperiod=<%=bperiod%>&utility=<%=utility%>&coor=<%=coor%>" Ismap><%if trim(billingid)="" then%></a><%end if%></div>
    </td>
  </tr>
</table>
</body>
</html>
