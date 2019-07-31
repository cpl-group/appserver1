<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
flag=Request.Form("flag")
key=Request.Form("id")
job=Request.Form("job")
d=Request.Form("date")
hours=Request.Form("hours")
description=Request.Form("description")
overt=Request.Form("overt")
billh=Request.Form("billh")
v=Request.Form("v")
expense=Request.Form("expense")
user="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")
	
strsql = "Update invoice_submission Set date='" & d & "' , description='"& description &"', hours='"& hours &"', expense='"& expense &"', overt='"& overt &"', hours_bill='"& billh &"',value=" & v &" where ( id='"& key &"')"

'Response.Write strsql
cnn1.execute strsql
'Response.write(job)

set cnn1=nothing
tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "corpinvoicedetail.asp?day="& flag &"&job="& job & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>