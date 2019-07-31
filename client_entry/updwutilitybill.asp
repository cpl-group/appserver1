<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim acctid, ypid, id, watercharge, sewer, totalamt, totalhcf, avgcost, bp, by, action, avgdailyusage
action=Request.form("action")
acctid=Request.form("acctid")
ypid=Request.form("ypid")
id=Request.form("id")
watercharge=Request.form("watercharge")
sewer=Request.form("sewer")
totalamt=Request.Form("totalamt")
totalhcf=Request.form("totalhcf")
avgcost=Request.Form("avgcost")
bp=Request.form("bp")
by=Request.form("by")
avgdailyusage = Request("avgdailyusage")

dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

if action="UPDATE" then
	strsql = "update utilitybill_coldwater set watercharge='" & watercharge & "',totalamt='" & totalamt & "',sewercharge='" & sewer & "',totalhcf='" & totalhcf & "',avgcost='" & avgcost & "', avgdailyusage='"&avgdailyusage&"' where id='" & id & "'"
elseif action="SAVE" then
	strsql = "insert into utilitybill_coldwater (Ypid, acctid, watercharge, totalamt, sewercharge, totalhcf, avgcost, avgdailyusage) values ("&ypid&", '"&acctid&"', '"& watercharge &"', '"& totalamt &"', '"& sewer &"', '"& totalhcf &"', '"& avgcost &"', '"&avgdailyusage&"')"
end if

cnn1.execute strsql
set cnn1=nothing

dim tmpMoveFrame
tmpMoveFrame =  "document.location = ""acctwdetail.asp?acctid="&acctid&"&ypid="&ypid&"&bp="&bp&"&by="&by& """"
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>

