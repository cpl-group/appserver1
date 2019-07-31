<!-- #include file="secure.inc" -->
<html>
<head>
<title>Genergy ERI Management</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="Microsoft Theme" content="none, default">
</head>
<style type="text/css">
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
catnr = Request("qcatnr")
userid= request ("userid")
%>
<frameset rows="18%,*" border=0> 
  <frame name="title" src="title.asp?qcatnr=<%=catnr%>&amp;userid=<%=userid%>" marginwidth="0" marginheight="0" scrolling="auto">
  <frameset rows="*,53%"border=0> 
    <frame name="piclist" src="piclist.asp?qcatnr=<%=catnr%>" scrolling="auto" marginwidth="0" marginheight="0">
    <frame name="info" src="blank.htm" scrolling="auto" target="_self" marginwidth="0" marginheight="0">
  </frameset>
  <noframes> 
  <body>
  <p>This page uses frames, but your browser doesn't support them.</p>
  </body>
  </noframes> </frameset>
</html>
