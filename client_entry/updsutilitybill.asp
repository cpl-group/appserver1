<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim acctid, ypid, id, f, tax, AvgCost, t, mlb, bp, by, action, HDD, CDD, Avg_DD
action=Request.form("action")
acctid=Request.form("acctid")
ypid=Request.form("ypid")
id=Request.form("id")
f=Request.Form("fuel")
tax=Request.Form("st")
AvgCost=Request.Form("AvgCost")
t=Request.form("totalamt")
mlb=Request.Form("mlb")
bp=Request.form("bp")
by=Request.form("by")
HDD = request("HDD")
CDD = request("CDD")
Avg_DD = request("Avg_DD")

dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")
if action="UPDATE" then
	strsql = "update utilitybill_steam set fueladj='" & f & "',salestax='" &tax & "',AvgCost='" & AvgCost & "',totalbillamt='" & t & "',mlbusage='" & mlb & "', HDD='" & HDD & "', CDD='" & CDD &"', Avg_DD='" & Avg_DD & "' where id=" & id
elseif action="SAVE" then
	strsql = "insert into utilitybill_steam (ypid, acctid, fueladj, salestax, AvgCost, totalbillamt, mlbusage, HDD, CDD, Avg_DD) values ("&ypid&", '"&acctid&"', '"&f&"', '"&tax&"', '"&AvgCost&"', '"&t&"', '"&mlb&"', '"& HDD &"', '"&CDD&"', '"&Avg_DD&"')"
end if
'response.write strsql
'response.end
cnn1.execute strsql
set cnn1=nothing

dim tmpMoveFrame
tmpMoveFrame =  "document.location = ""acctdetailsteam.asp?acctid="&acctid&"&ypid="&ypid&"&bp="&bp&"&by="&by&""""
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>

