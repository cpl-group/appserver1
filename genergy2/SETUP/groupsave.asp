<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, action, meters, viewname, grouplabel, groupid, gtype, edit, tenants, groups, input, output, isportfolio
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
action = secureRequest("action")
meters = secureRequest("meters")
tenants = secureRequest("tenants")
viewname = secureRequest("name")
grouplabel = secureRequest("label")
groupid = secureRequest("groupid")
gtype = secureRequest("type")
edit = secureRequest("edit")
input = secureRequest("inputGroups")
output = secureRequest("outputGroups")
groups = secureRequest("groups")
isportfolio = secureRequest("isportfolio")
groups = "0"
if trim(bldg)="" then bldg=0
if trim(tid) ="" then tid=0
if trim(edit)="" then
	edit = 0
else
	edit = 1
end if
if trim(isportfolio)<>"" then if trim(bldg)<>"0" then isportfolio = 0 else isportfolio = 1

dim cnn1, rst1, cmd, strsql, prm
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set cmd = server.createobject("ADODB.command")
if trim(isportfolio)="1" then
  cnn1.open getMainConnect(pid)
else
  cnn1.open getLocalConnect(bldg)
end if
'@string varchar(500),@type int,@group varchar(200) ,@name varchar(100),@user varchar(20),@edit int,@cid int, @bldg varchar(20),@tn varchar(20) AS
cmd.ActiveConnection = cnn1
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "Groups_v2"
Set prm = cmd.CreateParameter("string", adVarChar, adParamInput, 500)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("type", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("group", adVarChar, adParamInput, 200)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("name", adVarChar, adParamInput, 200)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("user", adVarchar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("edit", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("cid", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("tn", adVarChar, adParamInput,20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("portfolio", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("label", adVarChar, adParamInput, 40)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("OIP", adVarChar, adParamOutput, 25)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("Oname", adVarChar, adParamOutput, 25)
cmd.Parameters.Append prm
'response.write gtype
'response.end
if gtype=2 then
	cmd.parameters("string") = bldg
elseif gtype=5 then
  cmd.parameters("string") = input
  groups = output
else
	cmd.parameters("string") = meters
end if
if trim(groups)="" then groups = " "
if cmd.parameters("string") = "" then cmd.parameters("string")=" "
cmd.parameters("type") = gtype
cmd.parameters("group") = groups
cmd.parameters("name") = groupid
cmd.parameters("user") = getXMLUserName()
cmd.parameters("edit") = edit
cmd.parameters("cid") = cint(pid)
cmd.parameters("bldg") = bldg
cmd.parameters("tn") = tid
cmd.parameters("portfolio") = isportfolio
cmd.parameters("label") = grouplabel
'response.write "exec groups '"&cmd.parameters("string")&"','"&gtype&"','"&groups&"','"&groupid&"','"&getXMLUserName()&"','"&edit&"','"&cint(pid)&"','"&bldg&"','"&tid&"','"&isportfolio&"','"&grouplabel&"', 0"
'response.end

cmd.execute
if trim(trim(cmd.Parameters("OIP")))<>"" and trim(trim(cmd.Parameters("Oname")))<>"" then setView trim(trim(cmd.Parameters("Oname"))), trim(cmd.Parameters("OIP")), isportfolio

Response.redirect "groupview.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid
%>
