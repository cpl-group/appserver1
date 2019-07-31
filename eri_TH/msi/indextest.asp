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

<script>
function meters(bldg,meterid){
var urlPop = "/cgi-bin/pgimeter.asp?b="+bldg+"&m=" + meterid
document.frames.info.location = urlPop
}

</script>
<body bgcolor="#FFFFFF" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF">
<table width="800" border="1" cellspacing="0" cellpadding="0" height="480" align="center">
  <tr>
    <td width="400"><object classid ="clsid:B2BE75F3-9197-11CF-ABF4-08000996E931" codebase = "whip.cab#version=-1,-1,-1,-1" height=600 width=400>
        <param name="Filename" value="<%=Request.QueryString("pgi")%>">
      </object></td>
    <td width="51%"><IFRAME name="info" src="null.htm" width="100%" height="100%" scrolling="no" marginwidth="8" marginheight="16"></IFRAME></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
