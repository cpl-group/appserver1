<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim acctid, ypid, id, f, tax, p, t, therms, ccf, bp, by, action
action=Request.form("action")
acctid=Request.form("acctid")
ypid=Request.form("ypid")
id=Request.form("id")
f=Request.Form("fuel")
tax=Request.Form("st2")
p=Request.Form("AvgCostTherm")
t=Request.form("tamt2")
therms=Request.Form("therms")
ccf=Request.form("ccf")
bp=Request.form("bp")
by=Request.form("by")

dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

if action="UPDATE" then
	strsql = "update utilitybill_gas set fueladj='" & f & "',salestax='" &tax & "',AvgCostTherm='" &p & "',totalbillamt='" & t & "',thermusage='" & therms& "',ccfusage='" &ccf & "' where id=" & id & ""
elseif action="SAVE" then
	strsql = "insert into utilitybill_gas (ypid, acctid, fueladj, salestax, AvgCostTherm, totalbillamt, thermusage, ccfusage) values ("&ypid&", '"&acctid&"', '"&f&"', '"&tax&"', '"&p&"', '"&t&"', '"&therms&"', '"&ccf&"')"
end if

'response.write strsql
'response.end
cnn1.execute strsql
set cnn1=nothing
dim tmpMoveFrame
tmpMoveFrame =  "document.location = ""acctdetailgas.asp?acctid="&acctid&"&ypid="&ypid&"&bp="&bp&"&by="&by&"""" 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>

