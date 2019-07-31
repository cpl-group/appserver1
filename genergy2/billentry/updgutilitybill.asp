<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, acctid, ypid, id, f, tax, p, t, therms, ccf, bp, by, action
dim sec_CostPerTherm,sec_TransCharge,sec_salestax,sec_ullp,sec_ull,sec_total, sec_salestaxp, note, GRTAmt, GRTrate, salestaxrate, fuelrate, conversion, taxincluded
bldg = request("bldg")

function isVnum(num)
	isVnum = num
	if trim(num) = "" then isVnum = 0
	if not(isnumeric(num)) then isVnum = 0
	isVnum = cdbl(isVnum)
end function

action=request("action")
acctid=request("acctid")
ypid=request("ypid")
id=request("id")
f=isVnum(request("fuel"))
tax=isVnum(request("st2"))
p=isVnum(request("AvgCostTherm"))
t=isVnum(request("tamt2"))
therms=isVnum(request("therms"))
ccf=isVnum(request("ccf"))
bp=request("bp")
by=request("by")
note = replace(left(request("note"),500),"'","''")
GRTAmt=isVnum(request("GRTAmt"))
GRTrate=isVnum(request("GRTrate"))
salestaxrate=isVnum(request("salestaxrate"))
fuelrate=isVnum(request("fuelrate"))
conversion=isVnum(request("conversion"))
taxincluded=isVnum(request("taxincluded"))



sec_CostPerTherm = request("sec_CostPerTherm")
'sec_Therm = request("sec_Therm")
sec_TransCharge = request("sec_TransCharge")
sec_salestaxp = request("sec_salestaxp")
sec_salestax = request("sec_salestax")
sec_ullp = request("sec_ullp")
sec_ull = request("sec_ull")
sec_total = request("sec_total")

dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getLocalConnect(bldg)

if action="UPDATE" then
	strsql = "update utilitybill_gas set fueladj='" & f & "',salestax='" &tax & "',AvgCostTherm='" &p & "',totalbillamt='" & t & "',thermusage='" & therms& "',ccfusage='" &cint(ccf)& "', sec_CostPerTherm='"&sec_CostPerTherm&"',sec_TransCharge='"&sec_TransCharge&"',sec_salestax='"&sec_salestax&"',sec_salestaxp='"&sec_salestaxp&"',sec_ullp='"&sec_ullp&"',sec_ull='"&sec_ull&"',sec_total='"&sec_total&"', note='"&note&"', GRTAmt='"&GRTAmt&"', GRTrate='"&GRTrate&"', salestaxrate='"&salestaxrate&"', fuelrate='"&fuelrate&"', conversion='"&conversion&"', taxincluded='"&taxincluded&"' where id=" & id & ""
elseif action="SAVE" then
	strsql = "insert into utilitybill_gas (ypid, acctid, fueladj, salestax, AvgCostTherm, totalbillamt, thermusage, ccfusage, sec_CostPerTherm,sec_TransCharge,sec_salestax,sec_ullp,sec_ull,sec_total,sec_salestaxp, note, GRTAmt, GRTrate, salestaxrate, fuelrate, conversion, taxincluded) values ("&ypid&", '"&acctid&"', '"&f&"', '"&tax&"', '"&p&"', '"&t&"', '"&therms&"', '"&cint(ccf)&"','"&sec_CostPerTherm&"','"&sec_TransCharge&"','"&sec_salestax&"','"&sec_ullp&"','"&sec_ull&"','"&sec_total&"','"&sec_salestaxp&"', '"&note&"', '"&GRTAmt&"', '"&GRTrate&"', '"&salestaxrate&"', '"&fuelrate&"', '"&conversion&"', '"&taxincluded&"')"
end if

'response.write strsql
'response.end
cnn1.execute strsql
set cnn1=nothing
dim tmpMoveFrame
tmpMoveFrame =  "document.location = ""acctdetailgas.asp?bldg="&bldg&"&acctid="&acctid&"&ypid="&ypid&"&bp="&bp&"&by="&by&"""" 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>

