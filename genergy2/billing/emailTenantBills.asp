<%@ Language=VBScript %>
<%option explicit%>

<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"-->
<%end if

	function getNumber(number)
	'	response.write "|"&number&"|"
		if not(isNumeric(number)) then number = 0
		getNumber = number
	end function
	

	Dim  Billperiod, building, Billyear, PortFolioId, UtilityId, rpt, pdf, Genergy_Users, demo, sql
	
	' Set Parameters
	
	Billperiod = request("bperiod")
	building = request("building")
	Billyear = request("byear")
	PortFolioId = request("pid")
	UtilityId = trim(request("utilityid"))

	Dim rst1, rst2, rst3, cnn1
	Dim usage, demand, utilityname

	set rst1 = server.createobject("ADODB.Recordset")
	set cnn1 = server.createobject("ADODB.Connection")


	cnn1.open getLocalConnect(building)
	rst1.open "SELECT umeasure as usage, dmeasure as demand, utilitydisplay as utility " & _
			  "FROM tblutility WHERE UtilityId="&utilityid, getConnect(PortFolioId,building,"Billing")
	
	' Get Display names 
	If not rst1.eof then 
		usage = rst1("usage")
		demand = rst1("demand")
		utilityname = rst1("utility")
	End if
	rst1.close
	
	' Get Billyear, BillPeriod in case input parameters are blank.
	If trim(Billperiod)="" or trim(Billyear)="" then
		rst1.open "select top 1 BillYear, BillPeriod from tblmetersbyperiod "& _
					"WHERE bldgnum='"&building&"' and utility="&UtilityId&" ORDER BY billyear desc, billperiod desc", cnn1
		If rst1.eof then
			response.write "No information for this building"
			response.end
		Else
			Billyear = cint(rst1("billyear"))
			Billperiod = cint(rst1("billperiod"))
		End if
		rst1.close
	End if	

	dim DBlocalIP
	if trim(building)<>"" then DBlocalIP = ""	
	

	Dim cmd, prm
	set cmd = server.createobject("ADODB.Command") 
					
	cmd.ActiveConnection = cnn1
	cmd.CommandType = adCmdStoredProc
	cmd.CommandText = "usp_emailTenantBills"

	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("file", adVarChar, adParamInput,2)
	cmd.Parameters.Append prm


	cmd.parameters("bldg") = building
	cmd.parameters("by") = Billyear 
	cmd.parameters("bp") = Billperiod
	cmd.parameters("utility") = utilityid
	cmd.parameters("file") = "1"
							
	cmd.execute	
	
	Response.Write "<HTML><HEAD></HEAD><BODY><P> Tenant Bill email process has been started.<BR>"
	Response.Write "</P></BODY></HTML>"
     
	set rst1 = Nothing
	set cmd = Nothing
	set cnn1 = Nothing
	
	
%>	
	
