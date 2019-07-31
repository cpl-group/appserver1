<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
function checkBillval(byear, bperiod, building)
	dim billrt, billsql
	set billrt = server.createobject("ADODB.recordset")
	billsql = "SELECT (SELECT count(*) FROM Consumption c WHERE MeterID in (SELECT m.MeterID FROM Meters m INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid WHERE m.bldgnum='"&building&"' and m.lmp<>1 and Online<>'0' and leaseexpired=0 and lup.utility="&utilityid&") and c.BillYear="&byear&" and c.BillPeriod="&bperiod&") as total,(SELECT count(*) FROM Consumption c WHERE meterid in (SELECT m.MeterID FROM Meters m INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid WHERE m.bldgnum='"&building&"' and m.lmp<>1 and Online<>'0' and leaseexpired=0 and lup.utility="&utilityid&") and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and c.Validate=1) as validated, (SELECT count(*) as total FROM utilitybill u WHERE ypid in (SELECT ypid FROM billyrperiod WHERE billyear="&byear&" and billperiod="&bperiod&" and bldgnum='"&building&"' and utility="&utilityid&")) as utilitybills"
	billrt.open billsql, cnn1
	'response.write billsql
	'response.end
	if not billrt.EOF then
		checkBillval = (cint(billrt("total"))=cint(billrt("validated")) and cint(billrt("utilitybills"))>0)
	else
		checkBillval = false
	end if
end function

dim vmeters, byear, bperiod, building, pid, utilityid, mscroll, yscroll, yscroll2, showscroll, showscroll2
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

dim rst1, cnn1, sqlstr, index, cmd
set cmd = server.createobject("ADODB.command")
set cnn1 = server.createobject("ADODB.connection")
cnn1.open getLocalConnect(building)
cmd.ActiveConnection = cnn1
for each index in vmeters
	sqlstr = "UPDATE consumption SET validate=1, bval='"&trim(session("login"))&"' WHERE billyear="&byear&" and billperiod="&bperiod&" and meterid="&index
	cmd.commandtext = sqlstr
	cmd.execute
	'response.write sqlstr&"<BR>"
next

if checkBillval(byear, bperiod, building) then
	dim prm
	cmd.CommandText = "sp_runinvoice_bldg_v2"
	cmd.CommandType = adCmdStoredProc
	'input params
	cmd.commandTimeout = 1800
	Set prm = cmd.CreateParameter("bldg", adChar, adParamInput, 10)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("p", adInteger, adParamInput)
	cmd.Parameters.Append prm
	cmd.parameters("bldg") = building
	cmd.parameters("by") = byear
	cmd.parameters("bp") = bperiod
	cmd.parameters("utility") = utilityid
	cmd.parameters("p") = 0
	'cmd.execute

	set cmd = nothing
	set cmd = server.createobject("ADODB.command")
	cmd.ActiveConnection = cnn1
	cnn1.CursorLocation = adUseClient
	'specify stored procedure to run
	cmd.CommandText = "sp_billerval"
	cmd.CommandType = adCmdStoredProc
	'input params
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adVarChar, adParamInput, 4)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adVarChar, adParamInput, 2)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm

	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("bldg")		= building
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.Parameters("user")		= getXmlUserName()
	cmd.execute
	
	sendupdate building, byear, bperiod, getXmlUserName()
	
	response.write "<p style=""padding-top: 3px; padding-left: 3px; font-family: Arial, Helvetica, sans-serif;font-size:13"">All meters for "&request.form("buildingname")&", period "&bperiod&" of "&byear&" have been Accepted. An Email has been sent to the bill processor.</p>"
else
	response.redirect "validation_select.asp?pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod&"&utilityid="&utilityid&"&yscroll="&yscroll&"&mscroll="&mscroll&"&yscroll2="&yscroll2&"&showscroll="&showscroll&"&showscroll2="&showscroll2
end if

function sendupdate(bldg, by,bp,username)
	
	emailarray = "robertm@cplems.com"
	subject = "Building " & bldg & " Posted by Biller for Billperiod " &bp& "/" &by& ", ready for Supervisor Review."
	masternote = "Building has been posted by "&username
	sendmail emailarray,"filestore@cplems.com",subject, masternote

end function

%>
