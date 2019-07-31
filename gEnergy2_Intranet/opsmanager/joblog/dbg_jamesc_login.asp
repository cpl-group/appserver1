<html>
<head>
<%@Language="VBScript"%>
<%
Session("name")="James Canal"
Session("login")="davideg   "
Session("opslog")=3
Session("ts")= 4
Session("admin")= 5
Session("corp")= 5
response.redirect "frameset.html"
'response.Write("yo")
%>