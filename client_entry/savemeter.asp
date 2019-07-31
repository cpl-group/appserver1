
<%@Language="VBScript"%>

<%
bldg=Request.form("bldg")
acct=Request.form("acct")
id1=Request.form("meterid")
m=Request.form("meter")

sd=Request.form("sd")
loc=Request.form("loc")
dof=Request.form("dof")
online=Request.form("online")
mc=Request.form("description")
u=Request.form("utility")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

strsql = "insert meters1 (bldgnum,acctid,meternum,datestart,dateoffline,location,metercomments,utility,online) values (ltrim('" &bldg& "'),ltrim('" &acct& "'),ltrim('" & m & "'),ltrim('" & sd & "'),ltrim('" &dof & "'),ltrim('" &loc& "'),ltrim('" &mc & "'),ltrim('" &u & "'),'" &online& "')"
'RESPONSE.WRITE STRSQL
'RESPONSE.END 
cnn1.execute strsql
set cnn1=nothing

tmpMoveFrame =  "parent.document.frames.meters.location = " & Chr(34) & _
				  "metersearch.asp?acctid="&acct&"&bldg="&bldg&"&utility="&u& chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>
