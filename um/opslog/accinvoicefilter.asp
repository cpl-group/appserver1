<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
flag=Request("flag")
job=Request("job")
d=Request("date")
item=Request("item")
user="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")
strsql = "Update invoice_submission Set closed=1 where (jobno='"& job &"' and invoice_date='"&d&"')"
' no more billdate strsql2="Update [job log] set billdate=getdate() where [entry id]='"& job &"'"

cnn1.execute strsql

set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "accinvoice.asp?item="& item &"&date="& flag & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>