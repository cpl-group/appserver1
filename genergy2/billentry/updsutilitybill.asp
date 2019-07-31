<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, acctid, ypid, id, f, tax, AvgCost, t, mlb, bp, by, action, HDD, CDD, Avg_DD,GRT,stp,MAC, mlbs_hr, note, taxincluded
bldg = request("bldg")

action=Request("action")
acctid=Request("acctid")
ypid=Request("ypid")
id=Request("id")
f=getNumber(Request("fuel"))
tax=getNumber(Request("st"))
AvgCost=getNumber(Request("AvgCost"))
t=getNumber(Request("totalamt"))
mlb=getNumber(Request("mlb"))
bp=getNumber(Request("bp"))
by=getNumber(Request("by"))
HDD=getNumber(request("HDD"))
CDD=getNumber(request("CDD"))
Avg_DD=getNumber(request("Avg_DD"))
GRT=getNumber(request("GRT"))
stp=getNumber(request("stp"))
MAC=getNumber(request("MAC"))
mlbs_hr=getNumber(request("mlbs_hr"))
note = replace(left(request("note"),500),"'","''")
taxincluded = trim(request("taxincluded"))

if trim(GRT) = "" then GRT = "0"
if trim(stp) = "" then stp = "0"
if trim(MAC) = "" then MAC = "0"
if taxincluded = "" then taxincluded = "0"

dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getLocalConnect(bldg)
if action="UPDATE" then
	strsql = "update utilitybill_steam set fueladj='" & f & "',salestax='" &tax & "',AvgCost='" & AvgCost & "',totalbillamt='" & t & "',mlbusage='" & mlb & "', HDD='" & HDD & "', CDD='" & CDD &"', Avg_DD='" & Avg_DD & "',salestaxpercent = "&stp&",grtpercent = "&GRT&",macdollar = "&MAC&", mlbs_hr="&mlbs_hr&", note='"&note&"', taxincluded='"&taxincluded&"' where id=" & id
elseif action="SAVE" then
	strsql = "insert into utilitybill_steam (ypid, acctid, fueladj, salestax, AvgCost, totalbillamt, mlbusage, HDD, CDD, Avg_DD, salestaxpercent, grtpercent, MACdollar, mlbs_hr, note, taxincluded) values ("&ypid&", '"&acctid&"', '"&f&"', '"&tax&"', '"&AvgCost&"', '"&t&"', '"&mlb&"', '"& HDD &"', '"&CDD&"', '"&Avg_DD&"', '"&stp&"', '"&grt&"', '"&mac&"', '"&mlbs_hr&"', '"&note&"', "&taxincluded&")"
end if
'response.write strsql
'response.end
cnn1.execute strsql
set cnn1=nothing

dim tmpMoveFrame
tmpMoveFrame =  "document.location = ""acctdetailsteam.asp?bldg="&bldg&"&acctid="&acctid&"&ypid="&ypid&"&bp="&bp&"&by="&by&""""
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

function getNumber(number)
	if isnumeric(number) and number<>"" then getNumber = number else getNumber = 0
end function
%>

