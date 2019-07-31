
<%@Language="VBScript"%>

<%
u=Request.querystring("utility")
acct=Request.querystring("acctid")
id1=Request.querystring("meterid")
bldg=Request.querystring("bldg")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

strsql = "delete meters1 where acctid=ltrim('" &acct& "')and meterid=ltrim('" &id1& "')"

cnn1.execute strsql
set cnn1=nothing

tmpMoveFrame =  "parent.document.frames.meters.location = " & Chr(34) & _
				  "metersearch.asp?acctid="&acct&"&bldg="&bldg&"&utility="&u& chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>
