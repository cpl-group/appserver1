<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")
yio=request.querystring("key")
n=request.querystring("u")
		sqlstr = "delete times where id=" & Request.Querystring("key") & ""
		
cnn1.Execute sqlstr 
if request.querystring("Window")="close" then 
tmpMoveFrame="window.close();opener.document.location = opener.document.location;"
else
tmpMoveFrame =  "document.location = ""timesheet.asp"";"
end if


				 
				  
				  
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 				  
%>