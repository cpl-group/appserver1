<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
jobnum=Request.Form("jobnum")
podate=Request.Form("podate")
vendor=Request.Form("vendor")
jobadd=Request.Form("jobaddr")
shipadd=Request.Form("shipaddr")
req=Request.Form("req")
descr=Request.Form("description")
samt=Request.Form("ship_amt")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")

strsql = "insert po (jobnum,podate,vendor,jobaddr,shipaddr,requistioner,description)values ('" & jobnum& "', '" &podate & "', '" & vendor & "', '" & jobadd & "', '" & shipadd & "','" & req & "', '" & descr & "')"
cnn1.execute strsql

strsql = "select max (id) as id from po"
rst.Open strsql, cnn1, 0, 1, 1


if not rst.eof then
	poid = rst("id")
end if

set cnn1=nothing



tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "poview.asp?poid="& poid & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>
