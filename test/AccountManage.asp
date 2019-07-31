<%
dim usrmg, cmdPort, prmPort
set usrmg = server.createobject("UsrMgr.AcctMgr")

set cmdPort = server.createobject("adodb.command")
cmdPort.activeConnection = application("cnnstr_supermod")
cmdPort.CommandText = "sp_NewBldg_Pid"
cmdPort.CommandType = adCmdStoredProc
Set prmPort = cmdPort.CreateParameter("name", adVarChar, adParamInput, 15)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("pid", adInteger, adParamInput)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("sub", adBoolean, adParamInput)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("15min", adSmallInt, adParamInput)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("min", adSmallInt, adParamInput)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("bldgname", adVarChar, adParamInput, 50)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("strt", adVarChar, adParamInput, 100)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("city", adVarChar, adParamInput, 35)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("state", adChar, adParamInput, 2)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("zip", adVarChar, adParamInput, 9)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("sqft", adInteger, adParamInput)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("region", adInteger, adParamInput)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("facilitytype", adInteger, adParamInput)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("btbldgname", adVarChar, adParamInput, 50)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("btSTRT", adVarChar, adParamInput, 100)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("btcity", adVarChar, adParamInput, 35)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("btstate", adChar, adParamInput, 2)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("btzip", adVarChar, adParamInput, 9)
cmdPort.Parameters.Append prmPort
Set prmPort = cmdPort.CreateParameter("OIP", adVarChar, adParamOutput, 25)
cmdPort.Parameters.Append prmPort

cmdPort.Parameters("pid") = 0
cmdPort.Parameters("sub") = 0
cmdPort.Parameters("15min") = 1
cmdPort.Parameters("min") = 1
cmdPort.Parameters("STRT") = 0
cmdPort.Parameters("city") = 0
cmdPort.Parameters("state") = 0
cmdPort.Parameters("zip") = 0
cmdPort.Parameters("sqft") = 0
cmdPort.Parameters("region") = 0
cmdPort.Parameters("facilitytype") = 0
cmdPort.Parameters("btbldgname") = 0
cmdPort.Parameters("btSTRT") = 0
cmdPort.Parameters("btcity") = 0
cmdPort.Parameters("btstate") = 0
cmdPort.Parameters("btzip") = 0

'response.write "exec sp_NewBldg_Pid '"&portfolio&"', 0, 0, '"&buildings15&"', '"&buildings1&"', '"&portfolioname&"', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0" 
'response.end

function createNewAccount(orderid)
	dim rstNA, cmdNA, sql
	set rstNA = server.createobject("ADODB.recordset")
	set cmdNA = server.createobject("ADODB.command")
	cmdNA.ActiveConnection = application("cnnstr_security")
	cmdNA.CommandType = adCmdText
	
	dim username, password, fullname, companyname, pid, accountid, userblocks
	sql = "SELECT * FROM G1V2orders WHERE id="&orderid
	rstNA.open sql, application("cnnstr_security")
	if not rstNA.eof then
		username = rstNA("userid")
		password = rstNA("password")
		fullname = rstNA("firstname")&" "&rstNA("lastname")
		companyname = rstNA("company")
	else
		createNewAccount = err.Description
	end if
	rstNA.close
	usrmg.AddUser username, password, fullname, "GenergyOne Clients","clientFinancials|clientOperations"
	response.end
	cmdPort.Parameters("name") = left(companyname,3) 'this is a deprecated portfolio code field
	cmdPort.Parameters("bldgname") = companyname 'this is actually the Portfolio Name
	cmdPort.execute()
	rstNA.Open "SELECT max(id) as id FROM portfolio", application("cnnstr_supermod")
	if not rstNA.eof then pid = rstNA("id") else pid = 0
	rstNA.close
	
	rstNA.Open "SELECT isnull(sum(blocksize),0) FROM G1V2orderitems oi INNER JOIN G2_Block_Definitions bd ON oi.typeid=bd.id WHERE type=2 and orderid="&orderid, application("cnnstr_security")
	if not rstNA.eof then userblocks = rstNA(0) else userblocks = 0
	rstNA.close
	
	sql = "SET NOCOUNT ON;INSERT INTO G2_accounts (appid, usertype, name, email, userid, subs, pid) VALUES (1, 1, '"&replace(fullname,"'", "''")&"', 'email@email.com', '"&username&"',"&userblocks&", "&pid&");SELECT max(id) as newid FROM G2_accounts;"
	cmdNA.commandText = sql
	set rstNA = cmdNA.execute()
	if not rstNA.eof then accountid = rstNA("newid") else createNewAccount = err.Description
	rstNA.close
	
	sql = "SELECT * FROM (SELECT type, name, typeid FROM G1V2orderItems WHERE orderid="&orderid&" UNION ALL SELECT 1 as type, s.servicelabel as name, s.id as typeid FROM services s LEFT JOIN service_rates sr ON sr.serviceid=s.id WHERE isnull(price,0)=0 and servicetype=2) dd WHERE type=1"
	rstNA.open sql, application("cnnstr_security")
	do until rstNA.eof
		cmdNA.commandText = "INSERT INTO G2_userservices (userid, serviceid, servicelevel, serviceindex) VALUES ('"&accountid&"', '"&rstNA("typeid")&"', 'p', '"&pid&"')"
		cmdNA.execute()
		rstNA.movenext
	loop
	'createNewAccount = ""
	exit function
NewAccount_error:
	createNewAccount = errors
end function

function deleteUser(userid)
'	DelUser(fullName As Variant, OrgUnit As Variant)
end function
%>