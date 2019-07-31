
<%@Language="VBScript"%>

<%

acct=Request.form("acct2")
id1=Request.form("id2")
m=Request.form("meter")
sd=Request.form("sd2")
loc=Request.form("loc2")
dof=Request.form("dof2")
online=Request.form("online2")
mc=Request.form("textarea")
u=Request.form("utility2")
b=Request.form("bldg2")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

strsql = "update meters1 set datestart=ltrim('" & sd & "'),dateoffline=ltrim('" &dof & "'),location=ltrim('" &loc& "'),metercomments=ltrim('" &mc & "'),utility=ltrim('" &u & "'),online='" & online & "' where meterid=" & id1 & ""

cnn1.execute strsql
set cnn1=nothing

tmpMoveFrame =  "parent.document.frames.meters.location = " & Chr(34) & _
				  "metersearch.asp?acctid="&acct&"&bldg="&b&"&utility="&u& chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>
