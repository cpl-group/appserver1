
<%@Language="VBScript"%>

<%
acctid=Request.form("acctid")
utility=Request.form("utility")
yp=Request.form("ypid1")
f=Request.Form("fuel")
tax=Request.Form("st2")
p=Request.Form("prevbal")
t=Request.Form("totalamt")
therms=Request.Form("therms")
ccf=Request.form("ccf")
bp=Request.Form("bp")
by=Request.form("by")
response.write tamt
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

strsql = "insert utilitybill_gas (acctid,ypid,gasfueladj,salestax,previous_balance,totalbillamt,thermusage,ccfusage) values (" & acctid& "," & yp& "," & f & "," &tax & "," & p & "," &t & "," & therms& "," &ccf& ")"
'response.write strsql
'response.end

cnn1.execute strsql
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "acctdetailgas.asp?acctid="&acctid&"&ypid="&yp&"&bp="&bp&"&by="&by& chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>
