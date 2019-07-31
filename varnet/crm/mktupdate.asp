<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
cust=Request.Form("cust")
status=Request.Form("status")
id1=Request.Form("id1")
sit=Request.Form("sit")
manager = Request.Form("manager")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_main")


strsql = "Update mktlog Set contact='" & cust & "', status='" & status & "', situation='" & sit & "', salesmanager=" & manager &" where id='"& id1&"'"
cnn1.execute strsql


tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "mktview.asp?mkid="& id1 &"&cust="&cust& chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
set cnn1=nothing
%>


