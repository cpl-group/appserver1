<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, uid, bldg, tripcode,action

pid = secureRequest("pid")
uid = secureRequest("uid")
bldg = secureRequest("bldg")
tripcode = secureRequest("tripcode")
action = secureRequest("action")

'dim DBmainmodIP
'DBmainmodIP = "["&application("superip")&"].mainmodule.dbo."

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

if trim(lcase(action))="save" then
  strsql = "INSERT INTO super_tripcodes (pid, bldgnum, uid, tripcode) values ('"&pid&"', '"&bldg&"','"&uid&"','"&tripcode&"')"
else
  strsql = "UPDATE super_tripcodes set tripcode='"&tripcode&"' WHERE bldgnum='"&bldg&"' and uid = " & uid
end if
logger(strsql)
cnn1.Execute strsql
set cnn1 = nothing
%>
<script>
window.close()
</script>