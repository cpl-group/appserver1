<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
key=Request.Form("id")
flag=Request.Form("flag")
job=Request.Form("job")
d=Request.Form("date")
hours=Request.Form("hours")
description=Request.Form("description")
overt=Request.Form("overt")
billh=Trim(Request.Form("billh"))
v=Request.Form("v")
expense=Request.Form("expense")
c=Request.Form("typebox")
customer=Request.Form("customer")
contact=Request.Form("contact")
user="ghnet\"&Session("login")
des=Request.Form("des")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")
	
strsql = "Update invoice_submission Set date='" & d & "' , description='"& description &"', hours="& hours &", expense='"& expense &"', overt="& overt &", hours_bill="& billh &",value=" & v &", category=" & c &" where ( id='"& key &"')"


cnn1.execute strsql

set cnn1=nothing
tmpMoveFrame =  "parent.location = " & Chr(34) & _
				  "corpmain.asp?description="& des &"&day="& flag &"&job="& job & "&customer="& customer &"&contact="& contact &  chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>