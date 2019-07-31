<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, utype, bldg, ypid, action, billyear, billperiod, datestart, dateend, utility
pid = secureRequest("pid")
utype = secureRequest("utype")
action = secureRequest("action")
ypid = secureRequest("ypid")
bldg = secureRequest("bldg")
billyear = secureRequest("billyear")
billperiod = secureRequest("billperiod")
datestart = secureRequest("datestart")
dateend = secureRequest("dateend")
utility = secureRequest("utility")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

if trim(action)="Save" then
	strsql = "INSERT INTO billyrperiod (billyear, billperiod, datestart, dateend, utility, bldgnum) values ('"&billyear&"', '"&billperiod&"', '"&datestart&"', '"&dateend&"', '"&utility&"', '"&bldg&"')"
elseif trim(action)="Delete" then
	strsql = "DELETE FROM billyrperiod WHERE ypid="&ypid
else
	strsql = "UPDATE billyrperiod set billyear='"&billyear&"', billperiod='"&billperiod&"', datestart='"&datestart&"', dateend='"&dateend&"', utility='"&utility&"' WHERE ypid="&ypid
end if
'response.write strsql
'response.end
'Logging Update
logger(strsql)
'end Log

cnn1.Execute strsql

'if trim(action)="Save" then 'need to find the bid for the building just added
'	rst1.Open "SELECT max(ypid) as ypid FROM billyrperiod", cnn1
'	if not rst1.eof then ypid = rst1("ypid")
'end if
utype = utility
Response.redirect "billPeriodView.asp?pid="&pid&"&bldg="&bldg&"&utype="&utype
%>