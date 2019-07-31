
<%@Language="VBScript"%>

<%
acctid=Request.form("acctid")
utility=Request.form("utility")
yp=Request.form("ypid1")
f=Request.Form("fuel")
rec=Request.Form("rec")
tax=Request.Form("st")
tamt=Request.form("tamt")
onp=Request.Form("onp")
offp=Request.Form("offp")
tkwh=Request.Form("tkwh")
ckwh=Request.form("costkwh")
uckwh=Request.Form("ucostkwh")
tkw=Request.Form("tkw")
ckw=Request.Form("costkw")
uckw=Request.form("ucostkw")
bp=Request.Form("bp")
by=Request.form("by")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

strsql = "insert utilitybill (acctid,ypid,fueladj,grossreceipt,salestax,totalbillamt,onpeakkwh,offpeakkwh,totalkwh,costkwh,unitcostkwh,totalkw,costkw,unitcostkw) values (" & acctid& "," & yp& "," & f & "," & rec & "," &tax & "," & tamt & "," & onp& "," &offp & "," & tkwh & "," & ckwh & "," &uckwh & "," &tkw & "," & ckw & "," & uckw & ")"
'response.write strsql
'response.end

cnn1.execute strsql
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "acctdetail.asp?acctid="&acctid&"&ypid="&yp&"&bp="&bp&"&by="&by& chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>
