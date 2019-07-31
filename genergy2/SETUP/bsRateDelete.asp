<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim micid
micid = request("micid")

dim cnn, rst, strsql
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getLocalConnect(buildingNum)

strsql = "delete from misc_inv_credit where id = '" & micid & "'"

rst.open strsql, cnn
%>
<script>
alert("Information deleted.");
window.close();
</script>