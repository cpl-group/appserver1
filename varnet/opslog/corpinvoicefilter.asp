<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
flag=Request("flag")
job=Request("job")
d=Request("date")
item=Request("item")
count=Request("count")
user="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
strsql = "Update invoice_submission Set flag=1 where (jobno='"& job &"' and invoice_date='"&d&"')"
cnn1.execute strsql
sql="sp_corpinvoice_email '"& job &"'"
cnn1.Execute(sql)
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "corpinvoice.asp?item="& item &"&date="& flag & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>