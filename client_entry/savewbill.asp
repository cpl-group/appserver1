
<%@Language="VBScript"%>

<%
acct=Request.form("acct")
yp=Request.form("ypid2")
id1=Request.form("id1")
p=Request.form("prevbal")
s=Request.form("sewer")
t=Request.Form("totalamt")
h=Request.form("highhcf")
l=Request.Form("lowhcf")
bp=Request.form("bp")
by=Request.form("by")
avgdailyusage = Request("avgdailyusage")


Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

strsql = "insert utilitybill_water (acctid,ypid,previous_balance,sewercharge,totalbillamt,highhcfusage,lowhcfusage,avgdailyusage) values (" & acct& "," &yp& "," &p& "," & s& "," &t & "," & h& "," & l& "," & avgdailyusage & ")"

cnn1.execute strsql
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "acctwdetail.asp?acctid="&acct&"&ypid="&yp&"&bp="&bp&"&by="&by& chr(34) & vbCrLf

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 
%>
