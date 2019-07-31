<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#include file="checksession.asp"-->
<%
function checkSuperval(byear, bperiod, building)
	dim billrt, billsql
	set billrt = server.createobject("ADODB.recordset")
	billsql = "SELECT (SELECT count(*) FROM Consumption c WHERE c.MeterID in (SELECT m.MeterID FROM Meters m  INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid WHERE m.bldgnum='"&building&"' and m.pp<>1 and Online<>'0' and leaseexpired=0) and c.BillYear="&byear&" and c.BillPeriod="&bperiod&") as total,(SELECT count(*) FROM Consumption c WHERE meterid in (SELECT m.MeterID FROM Meters m INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid WHERE m.bldgnum='"&building&"' and m.pp<>1 and Online<>'0' and leaseexpired=0) and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and c.sValidate=1) as validated"
	billrt.open billsql, cnn1
'	response.write billsql
'	response.end

	if not billrt.EOF then
		 checkSuperval = (cint(billrt("total"))=cint(billrt("validated")))
	else
		checkSuperval = false
	end if
end function

dim vmeters, byear, bperiod, building, pid
vmeters = split(request("meters"),",")
byear = request("byear")
bperiod = request("bperiod")
building = request("building")
pid = request("pid")

dim rst1, cnn1, sqlstr, index, cmd, updatenum
set cmd = server.createobject("ADODB.command")
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy1")
cmd.ActiveConnection = cnn1
for each index in vmeters
	sqlstr = "UPDATE consumption SET svalidate=1, sval='"&trim(session("login"))&"' WHERE billyear="&byear&" and billperiod="&bperiod&" and meterid="&index
	cmd.commandtext = sqlstr
'	response.write sqlstr&"<BR>"
	cmd.execute
next
'sqlstr = "SELECT count(*) as userupdates FROM validation WHERE biller='"&session("login")&"' and billyear="&byear&" and billperiod="&bperiod
'rst1.open sqlstr, cnn1
'response.write sqlstr
'response.end
'if not rst1.eof then updatenum = trim(rst1("userupdates"))
'rst1.close
if checkSuperval(byear, bperiod, building) then
	dim prm
	cnn1.CursorLocation = adUseClient
	'specify stored procedure to run
	cmd.CommandText = "sp_superval"
	cmd.CommandType = adCmdStoredProc
	'input params
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("super", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adVarChar, adParamInput, 4)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adVarChar, adParamInput, 2)
	cmd.Parameters.Append prm

	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("bldg")		= building
	cmd.Parameters("super")		= session("login")
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.execute
	response.write "<p style=""padding-top: 3px; padding-left: 3px; font-family: Arial, Helvetica, sans-serif;font-size:13"">All meters for "&request.form("buildingname")&", period "&bperiod&" of "&byear&" have been Accepted.</p>"
else
	response.redirect "bill_validation.asp?pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod
end if
%>
