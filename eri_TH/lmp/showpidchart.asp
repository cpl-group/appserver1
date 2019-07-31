<%option explicit
dim pid, cdate, d
pid = request("pid")
d = request("d")
%>

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

<style type=3D"text/css"><!--A {text-decoration: none}--></style>
<body bgcolor="#FFFFFF" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#0099FF" onload="parent.closeLoadBox('loadFrame1');">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td>
      <div align="center"><img src="portfolioMakeChart.asp?d=<%=d%>&pid=<%=pid%>"></div>
    </td>
  </tr>
</table>
</body>
<script>
//openLoadBox('loadFrame1');
</script>

</html>
