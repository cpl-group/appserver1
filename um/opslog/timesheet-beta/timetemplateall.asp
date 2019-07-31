<html>
<head>
<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(getKeyValue("name")) then
%>
<script>
top.location="../index.asp"
window.close()
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if	

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

 
			sqlstr = "select startweek as s,endweek as e from time_submission where username='payroll'"
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
			start=rst1("s")
			end1=rst1("e")
					
					end if
					rst1.close


%>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="40,*" frameborder="NO" border="0" framespacing="0"> 
  <frame name="topFrame" src="printbutton.htm" scrolling="NO" style="border-bottom:1px solid #999999;">
  <frame name="mainframe" scrolling="Yes" noresize src="<%="timeformatall.asp?revstart=" & start & "&revend=" & end1 %>" >
</frameset>
<noframes> 
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
