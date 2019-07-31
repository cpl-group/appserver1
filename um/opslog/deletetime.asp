<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(Session("login")) then
			Response.Redirect "http://www.genergyonline.com"	
		end if
		
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")
yio=request.querystring("key")
n=request.querystring("u")
		sqlstr = "delete times where id=" & Request.Querystring("key") & ""
		
cnn1.Execute sqlstr 

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "timesheet.asp"& chr(34) & vbCrLf 
				 
				  
				  
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 				  
%>