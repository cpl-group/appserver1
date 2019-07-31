<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
'buildingNum=1450&bldgCPFrom=WGBH&utility=6&billyear=2003&billperiod=12
dim buildingNum, bldgCPFrom, utility, utilityto, billyear, billperiod
buildingNum = request("buildingNum")
bldgCPFrom = request("bldgCPFrom")
utility = request("utility")
utilityto = request("utilityto")
billyear = request("billyear")
billperiod = request("billperiod")

dim cmd, prm

Set cmd = server.createobject("ADODB.Command")
cmd.CommandText = "sp_copy_billperiods_v2"
cmd.CommandType = adCmdStoredProc
cmd.ActiveConnection = getLocalConnect(buildingNum)

Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("obldg", adVarChar, adParamInput,20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("utility", adSmallInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("nutility", adSmallInt, adParamInput)
cmd.Parameters.Append prm

cmd.Parameters("bldg") = buildingNum
cmd.Parameters("obldg") = bldgCPFrom
cmd.Parameters("by") = billyear
cmd.Parameters("bp") = billperiod
cmd.parameters("utility") = utility
cmd.parameters("nutility") = utilityto
'@bldg varchar(20),@obldg varchar(20),@by int,@bp int,@utility smallint ...... @bldg is the bldg you want to copy data to, @obldg is the bldg you want to copy data from. If you send 0 for by and bp, it will copy the entire history for the given utility. If you send 0 for the utility, it will copy all utilities for the given periods.  
cmd.Execute
'response.write "exec sp_copy_billperiods '"&buildingNum&"', '"&bldgCPFrom&"', "&billyear&", "&billperiod&", "&utility&"<br>"&getLocalConnect(buildingNum)
'response.end
%>
<script>
alert("Information copied.");
window.close();
</script>
