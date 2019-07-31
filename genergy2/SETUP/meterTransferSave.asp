<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim bldg, tid, pid, lid, transfermeter, bperiod, oldtid, meterid
lid = secureRequest("lid")
tid = secureRequest("tid")
pid = secureRequest("pid")
bldg = secureRequest("bldg")
transfermeter = secureRequest("transfermeter")
oldtid = secureRequest("oldtid")
if instr(secureRequest("bybp"),"|")=0 then
	bperiod = split("0|0","|")
else
	bperiod = split(secureRequest("bybp"),"|")
end if
'dim DBlocalIP
'DBlocalIP = "["&getBuildingIP(bldg)&"].genergy2.dbo."

dim cnn1, rst1, cmd, prm, lastmeterid
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

cmd.activeconnection = cnn1
cmd.CommandText = "sp_reassign_meters"
cmd.CommandType = adCmdStoredProc
Set prm = cmd.CreateParameter("meterid", adInteger, adParamInput, 250)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("lid", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("cip", adVarChar, adParamInput,24)
cmd.Parameters.Append prm

cmd.Parameters("bp")		= bperiod(1)
cmd.Parameters("by")		= bperiod(0)
cmd.Parameters("lid")		= lid
cmd.Parameters("cip")		= getIP(pid,bldg)

dim eachmeter
for each eachmeter in split(transfermeter,",")
  cmd.Parameters("meterid")	= eachmeter
'  response.write "exec sp_reassign_meters "&eachmeter&", "&bperiod(1)&", "&bperiod(0)&", "&lid&"<br>"
  cmd.Execute
next
'response.end

rst1.open "SELECT max(meterid) as id FROM meters", getConnect(pid,bldg,"billing")
meterid = rst1("id")
rst1.close
if instr(request.servervariables("HTTP_REFERER"),"TenantMeterTransfer.asp")>0 then
	Response.Redirect "TenantMeterTransfer.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&oldtid="&oldtid
else
	Response.Redirect "tenantedit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&meterid="&lastmeterid
end if
%>
