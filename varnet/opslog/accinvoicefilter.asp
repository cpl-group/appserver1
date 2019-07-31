<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
flag=Request("flag")
job=Request("job")
d=Request("date")
item=Request("item")
user="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
strsql = "Update invoice_submission Set closed=1 where (jobno='"& job &"' and invoice_date='"&d&"')"
strsql2="Update [job log] set billdate=getdate() where [entry id]='"& job &"'"

cnn1.execute strsql
'Response.Write strsql2
'response.end
cnn1.execute strsql2
'Response.write(job)

set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "accinvoice.asp?item="& item &"&date="& flag & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>