<%@Language="VBScript"%>
<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include file="adovbs.inc" -->
<%
dim id, cnn1, strsql, tmpMoveFrame
id=Request("id1")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")
strsql = "Update po Set closed=1, Acct_ponum = '"&request("acctponum")&"', closed_user='"&getXMLUserName()&"' where id='"& id &"'"

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