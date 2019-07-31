<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

function checkSuperval(byear, bperiod, fieldValue, fieldCheck)
	dim billrt, billsql
	set billrt = server.createobject("ADODB.recordset")
	billsql = "SELECT (SELECT count(*) FROM Consumption c WHERE c.MeterID in (SELECT m.MeterID FROM Meters m  INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid WHERE m."&fieldCheck&"='"&fieldValue&"' and m.lmp<>1 and Online<>'0' and leaseexpired=0  and nobill=0 and lup.utility="&utilityid&") and c.BillYear="&byear&" and c.BillPeriod="&bperiod&") as total,(SELECT count(*) FROM Consumption c WHERE meterid in (SELECT m.MeterID FROM Meters m INNER JOIN tblleasesutilityprices lup on m.leaseutilityid=lup.leaseutilityid INNER JOIN tblleases l ON lup.billingid=l.billingid WHERE m."&fieldCheck&"='"&fieldValue&"' and m.lmp<>1 and Online<>'0' and leaseexpired=0 and lup.utility="&utilityid&") and c.BillYear="&byear&" and c.BillPeriod="&bperiod&" and c.sValidate=1) as validated"
	'response.write billsql
	billrt.open billsql, cnn1
	'response.end
  
	if not billrt.EOF then
		checkSuperval = (cint(billrt("total"))<=cint(billrt("validated")))
	else
		checkSuperval = false
	end if
	
end function

dim pid, building, byear, lid, bperiod, action, utilityid, filename, historic, note
pid = request("pid")
building = request("building")
note = request("note")
if instr(request("bperiod"),"/")>0 then
	byear = split(request("bperiod"),"/")(1)
	bperiod = split(request("bperiod"),"/")(0)
else
	byear = 0
	bperiod = 0
end if
utilityid = request("utilityid")
historic = request("historic")
lid = request("lid")
action = request("action")
if action="" then action = request("actions")

dim rst1, cnn1, sql, cmd, prm, sql2
set rst1 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
cnn1.open getConnect(pid,building,"billing")

'response.write checkSuperval(byear, bperiod, building, "bldgnum")
'response.End()

cnn1.commandTimeout = 0
if trim(action)="Delete Bill" or trim(action)="Delete All Bills" then
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "sp_unpostbill_v2"
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("lid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("delete", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("note", adVarChar, adParamInput, 1000)
	cmd.Parameters.Append prm
	cmd.parameters("bldg") = building
	if trim(action)="Delete Bill" then cmd.parameters("lid") = lid else cmd.parameters("lid") = 0
	cmd.parameters("by") = byear
	cmd.parameters("bp") = bperiod
	cmd.parameters("utility") = utilityid
	cmd.parameters("delete") = 1
	cmd.parameters("user") = getXmlUserName()
	cmd.parameters("note") = "" 'N.Ambo added note filed 9/29/2008
	'response.write "exec sp_unpostbill_v2 '"&cmd.parameters("bldg")&"',"&cmd.parameters("lid")&","&cmd.parameters("by")&","&cmd.parameters("bp")&","&cmd.parameters("utility")&","&cmd.parameters("delete")
	'response.end
	cmd.execute
elseif trim(action)="Produce Bills For Current Period" or trim(action)="Produce Partial Bills" then
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "sp_runinvoice_bldg_v2"
	cmd.commandTimeout = 0
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
	if trim(action)="Produce Bills For Current Period" then cmd.parameters("p") = 0 else cmd.parameters("p") = 1
	'response.write "exec sp_runinvoice_bldg_v2 '"&building&"','"&byear&"','"&bperiod&"','"&utilityid&"', "&cmd.parameters("p")&"<br>"
	'response.write cmd.activeconnection
	'response.end
	
	cmd.execute
	' Added by Tarun 07/14/2006 to generate invoice numbers for port authority
	' Modified by Tarun : Invoice numbers need to be generated after bills have been posted.

	'cmd.CommandText = "usp_GeneratePABillInvoiceId"
	'cmd.CommandType = adCmdStoredProc
	'cmd.Parameters.Delete "p" 
	'cmd.Execute 

elseif trim(action)="GenerateCorrections" then
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "sp_runCorrections_bldg"
	cmd.commandTimeout = 0
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
	if trim(action)="GenerateCorrections" then cmd.parameters("p") = 0 else cmd.parameters("p") = 1
	cmd.execute

	
elseif trim(action)="Post Bills" and checkSuperval(byear, bperiod, building, "bldgnum") then
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
	Set prm = cmd.CreateParameter("post", adVarChar, adParamInput, 2)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("utility", adInteger, adParamInput, 2)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("lease", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("tidip", adVarChar, adParamInput, 5)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("note", adVarchar, adParamInput, 1000)
	cmd.Parameters.Append prm
	
	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("bldg")		= building
	cmd.Parameters("super")		= getXmlUserName()
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.Parameters("post")		= 1
	cmd.Parameters("utility")	= utilityid
	cmd.Parameters("lease")		= 0
	cmd.Parameters("tidip")		= split(getBuildingIP(building),"\")(1)
	cmd.Parameters("note") = ""
	'response.write "exec sp_superval '"&building&"', '"&getXmlUserName()&"', "&byear&", "&bperiod&", 1, " &utilityid&",0,0," & split(getBuildingIP(building),"\")(1)
	'response.end
	cmd.execute

	' Modified by Tarun to 08/022006 to call Invoice Number generation process after the bills have been posted.

	cmd.Parameters.Delete "super"
	cmd.Parameters.Delete "post"
	cmd.Parameters.Delete "tidip"
	cmd.Parameters.Delete "note"

	cmd.parameters("bldg") = building
	cmd.parameters("by") = byear
	cmd.parameters("bp") = bperiod
	cmd.parameters("utility") = utilityid
	cmd.Parameters("lease")		= 0

	cmd.CommandText = "usp_GeneratePABillInvoiceId"
	cmd.CommandType = adCmdStoredProc

	cmd.Execute
	'#

	
elseif trim(action)="Post" and checkSuperval(byear, bperiod, lid, "leaseutilityid") then
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
	Set prm = cmd.CreateParameter("post", adVarChar, adParamInput, 2)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("lease", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("TIDIP", adVarChar, adParamInput, 5)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("note", adVarchar, adParamInput, 1000)
	cmd.Parameters.Append prm
	
	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("bldg")		= building
	cmd.Parameters("super")		= getXmlUserName()
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.Parameters("post")		= 1
	cmd.Parameters("utility")	= utilityid
	cmd.Parameters("lease")		= lid
	cmd.Parameters("tidip")		= split(getBuildingIP(building),"\")(1)
	cmd.Parameters("note") = ""
	'response.write "exec sp_superval '"&building&"', '"&getXmlUserName()&"', "&byear&", "&bperiod&", 1, " &utilityid & "," & lid & "," & split(getBuildingIP(building),"\")(1) & ", '" & note & "' ;" 
	'response.end
	cmd.execute

	' Modified by Tarun to 08/022006 to call Invoice Number generation process after the bills have been posted.
	cmd.Parameters.Delete "super"
	cmd.Parameters.Delete "post"
	cmd.Parameters.Delete "tidip"
	cmd.Parameters.Delete "note"
	
	cmd.parameters("bldg") = building
	cmd.parameters("by") = byear
	cmd.parameters("bp") = bperiod
	cmd.parameters("utility") = utilityid
	cmd.Parameters("lease")		= lid	

	cmd.CommandText = "usp_GeneratePABillInvoiceId"
	cmd.CommandType = adCmdStoredProc

	cmd.Execute
	'#
	
elseif trim(action)="Unpost Bills" or trim(action)="Unpost" then
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "sp_unpostbill_v2"
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("lid", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("delete", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 30)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("note", adVarChar, adParamInput, 1000)
	cmd.Parameters.Append prm
	cmd.parameters("bldg") = building
	if trim(action)="Unpost" then cmd.parameters("lid") = lid else cmd.parameters("lid") = 0
	cmd.parameters("by") = byear
	cmd.parameters("bp") = bperiod
	cmd.parameters("utility") = utilityid
	cmd.parameters("delete") = 0
	cmd.parameters("user") = getXmlUserName()
	cmd.parameters("note") = note
	cmd.execute
elseif trim(action)="IBS" then
	cnn1.CursorLocation = adUseClient
	'specify stored procedure to run
	cmd.CommandText = "sp_IBS"
	cmd.CommandType = adCmdStoredProc
	'input params
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("yr", adVarChar, adParamInput, 4)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("per", adVarChar, adParamInput, 2)
	cmd.Parameters.Append prm
'	Set prm = cmd.CreateParameter("filename", adVarChar, adParamInput, 25)
'	cmd.Parameters.Append prm

	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("bldg")		= building
	cmd.Parameters("yr")		= byear
	cmd.Parameters("per")		= bperiod
	cmd.execute
	'filename = cmd.Parameters("filename")
end if

if trim(action)<>"IBS" then 
	if trim(action)="GenerateCorrections" then
		response.redirect "processor_select.asp?pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod&"/"&byear&"&lid="&lid&"&utilityid="&utilityid&"&historic="&historic&"&corrections=true"
	else
		response.redirect "processor_select.asp?pid="&pid&"&building="&building&"&byear="&byear&"&bperiod="&bperiod&"/"&byear&"&lid="&lid&"&utilityid="&utilityid&"&historic="&historic&"&corrections=false"
	end if 
end if
%>
<html><head><title>Untitled</title></head>
<link rel="Stylesheet" href="../styles.css" type="text/css">
<body>
<div align="center">&nbsp;<br>&nbsp;<br><a href="/downloads/<%=filename%>">Download BMI for building <%=building%>, period <%=bperiod%> of <%=byear%></a><br>&nbsp;<br>&nbsp;<br>&nbsp;</div>
<div align="right"><a href="javascript:close()" style="background-color: Black; color:white; padding: 0px; margin-bottom: 6px; text-decoration: none;"><b>Close Window</b></a></div>
</body>
</html>

