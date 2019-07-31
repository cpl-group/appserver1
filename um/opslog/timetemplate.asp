<html>
<head>
<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%


'Request.QueryString("user")


		'if isempty(getKeyValue("name")) then
%>
<!--<script>
top.location="../index.asp"
window.close()
</script>-->
<%
			'			Response.Redirect "http://www.genergyonline.com"
		'else
			'if getKeyValue("ts") < 4 then 
				'getKeyValue("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				'Response.Redirect "../main.asp"
			'end if	
		'end if	



%>






<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
  <frame name="topFrame" src="printbutton.htm" scrolling="NO" style="border-bottom:1px solid #999999;">

<frame name="mainframe" scrolling="Yes" noresize src="timeformat.asp?user=<%=Request.QueryString("user")%>">
</frameset>
<noframes> 
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
