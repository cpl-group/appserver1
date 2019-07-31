<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
meterid=Request.form("meterid")
lmchannel=Request.form("lmchannel")
lmnum=request.Form("lmnum")
meternum=Request.form("meternum")
srvname=Request.form("srvname")
srvname1=Request.form("srvname1")
dbname=Request.form("dbname")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
'sql = "Update meters Set lmnum='"&lmnum&"', lmchannel='"&lmchannel&"' where meterid='"& meterid&"'"

sql ="update OPENQUERY(["& srvname1 &"], 'select lmnum,lmchannel from meters where meterid = "& meterid&"') Set lmnum='"&lmnum&"', lmchannel='"&lmchannel&"'"

'response.write sql
'response.end
cnn1.execute sql

set cnn1=nothing

tmpMoveFrame =  "parent.frames.meter.location = " & Chr(34) & _
				  "lminfo.asp?meternum="& meternum &"&srvname="&srvname1&"&dbname="&dbname& chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf  
%>