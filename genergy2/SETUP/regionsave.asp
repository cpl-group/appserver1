<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql, cmd, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, action, city, citycode, copyfrom
rid = secureRequest("rid")
city = secureRequest("city")
citycode = secureRequest("citycode")
copyfrom = secureRequest("copyfrom")
action = secureRequest("action")

if trim(action)="Copy" then
	cmd.CommandText = "sp_copy_region"
	cmd.CommandType = adCmdStoredProc
	'input params
	Set prm = cmd.CreateParameter("id", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("region", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("rcode", adVarChar, adParamInput, 5)
	cmd.Parameters.Append prm
	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("id")		= copyfrom
	cmd.Parameters("region")	= city
	cmd.Parameters("rcode")		= citycode
	cmd.execute
elseif trim(action)="Save" then
	strsql = "INSERT INTO regions (city, city_code) values ('"&city&"', '"&citycode&"')"
else
	strsql = "UPDATE regions set city='"&city&"', city_code='"&citycode&"' WHERE id="&rid
end if
'response.Write strsql
'response.End
cnn1.Execute strsql
if trim(action)="Save" then 'need to find the bid for the building just added
	rst1.Open "SELECT max(id) as id FROM regions", cnn1
	if not rst1.eof then rid = rst1("id")
end if

Response.redirect "regionedit.asp?rid="&rid
%>