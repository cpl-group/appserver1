<%option explicit

dim bldgid, lmpdate, qrytype, graphtype
bldgid = request("bldgid")
lmpdate = request("lmpdate")
qrytype = request("qrytype")
graphtype = request("graphtype")
%>
<html>
<head>
<title>LMP Chart</title>
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

<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" onload="parent.closeLoadBox('loadFrame<%if qrytype="dam" then%>2<%else%>1<%end if%>');"><!-- onload="propagateVarsUP();"> -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center">
  <tr>
    <td>
      <div align="center"><img src="<%="rtp/charts.asp?bldgid="&bldgid&"&lmpdate="&lmpdate&"&qrytype="&qrytype&"&graphtype="&graphtype%>"></div>
    </td>
  </tr>
</table>
</body>
</html>
