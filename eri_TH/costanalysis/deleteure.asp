<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
entryid=Request.QueryString("eid")
pid = Request.QueryString("pid")
b=Request.QueryString("b")
date1=Request.QueryString("date1")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")
sql = "delete tblRPentries where id=" & entryid
cnn1.execute sql
set cnn1=nothing
urltemp = "unreported.asp?building=" & b & "&date1=" & date1 & "&pid=" & pid &"&action=new"
tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				urltemp & chr(34) & vbCrLf 
'response.write urltemp
'response.end

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>