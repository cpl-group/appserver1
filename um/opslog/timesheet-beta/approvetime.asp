<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(getKeyValue("login")) then
			Response.Redirect "http://www.genergyonline.com"	
		end if
name1=Request.Querystring("userapp")

		
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")
	
	sqlstr = "exec sp_corpaccepttime_email  '" & name1 & "'"
	cnn1.Execute sqlstr 

	response.redirect "admintime.asp"	

%>
