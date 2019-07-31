<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
id=Request("id1")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
strsql = "Update po Set closed=1 where id='"& id &"'"

'Response.Write strsql
cnn1.execute strsql
'Response.write(job)

set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "acctpoview.asp" & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>