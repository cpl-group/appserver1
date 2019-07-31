<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'5/19/2008 N.AMbo added this page to save and/or update utility bill data entered on the 'Historical Data Entry' screeen

if 	not(allowGroups("Genergy Users,clientOperations")) then
%> <!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim building, utilityid, pid, byear, bperiod
dim buttonval, totalkwh, totalkw, costkwh, costkw, totalbilled


dim cnn1, rst1, cmd4, prm
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set cmd4 = server.createobject("ADODB.Command")

pid = secureRequest("pid")
building = secureRequest("bldgNum")
utilityid = secureRequest("utilityid")
totalkwh = secureRequest("totalkwh")
totalkw = secureRequest("totalkw")
costkwh = secureRequest("totalkwhcost")
costkw = secureRequest("totalkwcost")
totalbilled = secureRequest("totalbilled")
if instr(secureRequest("bperiod"),"/")>0 then
	byear = split(secureRequest("bperiod"),"/")(1)
	bperiod = split(secureRequest("bperiod"),"/")(0)
else
	byear = secureRequest("byear")
	bperiod = secureRequest("bperiod")
end if

if not isnumeric(totalkwh) or not isnumeric(totalkw) or not isnumeric(costkwh) or not isnumeric(costkw) or not isnumeric(totalbilled) then
	%>
<script> alert("All values must be numeric!")</script>
<%
response.Write "Click the Back button in your browser."
else

	cnn1.open getLocalConnect(building)

	cmd4.CommandType = adCmdStoredProc
	cmd4.CommandText = "usp_insertBillingData"
	Set prm = cmd4.CreateParameter("bldgnum", adVarChar, adParamInput, 10, building)
	cmd4.Parameters.Append prm
	Set prm = cmd4.CreateParameter("billyear", adVarChar, adParamInput, 4, byear)
	cmd4.Parameters.Append prm
	Set prm = cmd4.CreateParameter("billperiod", adVarChar, adParamInput, 4, bperiod)
	cmd4.Parameters.Append prm
	Set prm = cmd4.CreateParameter("utility", adInteger, adParamInput, , utilityid)
	cmd4.Parameters.Append prm
	Set prm = cmd4.CreateParameter("totalkwh", adInteger, adParamInput, , totalkwh)
	cmd4.Parameters.Append prm
	Set prm = cmd4.CreateParameter("totalkw", adInteger, adParamInput, , totalkw)
	cmd4.Parameters.Append prm
	Set prm = cmd4.CreateParameter("costkwh", adInteger, adParamInput, , costkwh)
	cmd4.Parameters.Append prm
	Set prm = cmd4.CreateParameter("costkw", adInteger, adParamInput, , costkw)
	cmd4.Parameters.Append prm
	Set prm = cmd4.CreateParameter("totalbillamt", adInteger, adParamInput, , totalbilled)
	cmd4.Parameters.Append prm

	cmd4.ActiveConnection = cnn1

	On error resume next
	cmd4.Execute
	if err.number <> 0 then 
		msgbox "An error occured while saving/updating.Please try again."
		err.Clear
	end if

	cnn1.Close
	
	response.Redirect "historicaldataentry.asp?pid="&pid&"&bldgNum="&building&"&utilityid="&utilityid


end if
%>
