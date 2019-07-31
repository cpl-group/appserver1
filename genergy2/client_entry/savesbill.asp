
<%@Language="VBScript"%>

<%
acctid=Request.form("acctid")
utility=Request.form("utility")
yp=Request.form("ypid")
f=Request.Form("fuel")
tax=Request.Form("st")
p=Request.Form("prevbal")
t=Request.form("totalamt")
mlb=Request.Form("mlb")
bp=Request.Form("bp")
by=Request.form("by")
response.write acctid
response.write utility
response.write bp
response.write by
response.write yp
'response.end
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

strsql = "insert utilitybill_steam (acctid,ypid,fueladj,salestax,previous_balance,totalbillamt,mlbusage) values ('" & acctid& "'," & yp& "," & f & "," &tax & "," & p & "," &t & "," & mlb& ")"
'response.write strsql
'response.end

cnn1.execute strsql
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "acctdetailsteam.asp?acctid="&acctid&"&ypid="&yp&"&bp="&bp&"&by="&by& chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>
