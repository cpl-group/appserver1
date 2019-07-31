<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
function checkSuperval(byear, bperiod, building)
	dim billrt, billsql
	set billrt = server.createobject("ADODB.recordset")
	billsql = "SELECT (SELECT count(*) FROM Consumption c WHERE MeterID in (SELECT m.MeterID FROM Meters m INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid WHERE m.bldgnum='"&building&"' and m.lmp<>1 and Online<>'0' and leaseexpired=0 and lup.utility="&utilityid&") and c.BillYear="&byear&" and c.BillPeriod="&bperiod&") as total,(SELECT count(*) FROM Consumption c WHERE meterid in (SELECT m.MeterID FROM Meters m INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid WHERE m.bldgnum='"&building&"' and m.lmp<>1 and Online<>'0' and leaseexpired=0) and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and c.sValidate=1) as validated, (SELECT count(*) as total FROM utilitybill u WHERE ypid in (SELECT ypid FROM billyrperiod WHERE billyear="&byear&" and billperiod="&bperiod&" and bldgnum='"&building&"')) as utilitybills"
	billrt.open billsql, cnn1
	if not billrt.EOF then
		checkSuperval = (cint(billrt("total"))=cint(billrt("validated")) and cint(billrt("utilitybills"))>0)
	else
		checkSuperval = false
	end if
end function

dim vmeters, byear, bperiod, building, pid, utilityid, mscroll, yscroll, yscroll2, showscroll, showscroll2,viewtype
vmeters = split(request("meters"),",")
byear = request("byear")
bperiod = request("bperiod")
building = request("building")
pid = request("pid")
utilityid = request("utilityid")
mscroll = request("mscroll")
yscroll = request("yscroll")
yscroll2 = request("yscroll2")
showscroll = lcase(trim(request("showscroll")))
showscroll2 = lcase(trim(request("showscroll2")))
viewtype = request("t")
'response.write viewtype
'response.end

dim rst1, cnn1, sqlstr, index, cmd, updatenum
set cmd = server.createobject("ADODB.command")
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,building,"billing")
cmd.ActiveConnection = cnn1
for each index in vmeters
	sqlstr = "UPDATE consumption SET svalidate=1, sval='"&trim(getXMLusername())&"', validate=1,bval= case validate when 0 then '"&trim(getXMLusername())&"' end WHERE billyear="&byear&" and billperiod="&bperiod&" and meterid="&index
	cmd.commandtext = sqlstr
	cmd.execute
next
if checkSuperval(byear, bperiod, building) then
	dim prm
	cnn1.CursorLocation = adUseClient
	cmd.CommandText = "sp_superval"
	cmd.CommandType = adCmdStoredProc
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("super", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adVarChar, adParamInput, 4)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adVarChar, adParamInput, 2)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("post", adVarChar, adParamInput, 2)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("utility", adInteger, adParamInput, 2)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("lease", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("tidip", adVarChar, adParamInput, 5)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("note", adVarChar, adParamInput, 1000)
	cmd.Parameters.Append prm

	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("bldg")		= building
	cmd.Parameters("super")		= getXmlUserName()
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.Parameters("post")		= 0
	cmd.Parameters("utility")	= utilityid
	cmd.Parameters("lease")		= 0
	cmd.Parameters("tidip")		= split(getBuildingIP(building),"\")(1)
	cmd.Parameters("note")      = ""
    cmd.execute
	'response.write cmd.Parameters("bldg") &"," & cmd.Parameters("super")&"," & cmd.Parameters("by")&"," &cmd.Parameters("bp")	&"," & cmd.Parameters("post")&"," & cmd.Parameters("utility")&"," & cmd.Parameters("lease")&"," &cmd.Parameters("tidip")
	'response.end
	sendupdate building, byear, bperiod, getXmlUserName()
	dim processnote
	processnote = server.urlencode("<p style=""padding-top: 3px; padding-left: 3px; font-family: Arial, Helvetica, sans-serif;font-size:13"">All meters for "&request.form("buildingname")&", period "&bperiod&" of "&byear&" have been Accepted.</p>")
	response.redirect "re_index.asp?t="&viewtype&"&pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod&"&utilityid="&utilityid&"&yscroll="&yscroll&"&mscroll="&mscroll&"&yscroll2="&yscroll2&"&showscroll="&showscroll&"&showscroll2="&showscroll2&"&processnote="&processnote
else
	response.redirect "re_index.asp?t="&viewtype&"&pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod&"&utilityid="&utilityid&"&yscroll="&yscroll&"&mscroll="&mscroll&"&yscroll2="&yscroll2&"&showscroll="&showscroll&"&showscroll2="&showscroll2
end if

function sendupdate(bldg, by,bp,username)
	Dim emailarray, subject, masternote
	emailarray = "robertm@cplems.com"
	subject = "Building " & bldg & " Posted by Biller for Billperiod " &bp& "/" &by& ", ready for Robs Review."
	masternote = "Building has been posted by "&username
	sendmail emailarray,"filestore@cplems.com",subject, masternote
end function
%>
