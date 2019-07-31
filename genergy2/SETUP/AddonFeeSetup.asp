<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim bldg, pid, bldgName, adfid, adfamount, adfsubmit

pid = request("pid")
bldg = request("bldg")
adfid = request("adfid")
adfamount = trim(request("adfamount"))
adfsubmit = request("adfsubmit")
if adfamount="" or not(isnumeric(adfamount)) then adfamount = 0
dim cnn1, cnnMainModule, strsql, rst1, cmd
set cnnMainModule = server.createobject("ADODB.connection")
cnnMainModule.open getConnect(pid,bldg,"billing")
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

''submit section
if adfsubmit="Add New" then
	cmd.CommandText = "INSERT INTO Building_AddonFee (bldgnum,AddonFee) Values ('"&bldg&"',"&adfamount&")"
elseif adfsubmit="Change" then
	cmd.CommandText = "UPDATE Building_AddonFee SET AddonFee="&adfamount&" WHERE id="&adfid
elseif adfsubmit="Delete" then
	cmd.CommandText = "DELETE Building_AddonFee WHERE id="&adfid
end if
if cmd.CommandText<>"" then
	cmd.activeconnection = cnn1
	'response.write cmd.CommandText
	'response.end
	cmd.execute()
	adfamount = 0
	adfid=0
end if

strsql = "select bldgname from buildings where bldgnum = '"&bldg&"'"
rst1.open strsql, cnn1
if not rst1.eof then
	bldgName = rst1("bldgname")
end if
rst1.close
%>	
<link rel="Stylesheet" href="setup.css" type="text/css">
<title>Addon Fee Setup for <%=bldgname%></title>
<script>
function closeWinda(){
window.close()
}
</script>
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form action="addonfeesetup.asp" method="post">
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#3399cc">
		<td>
			<font color='white'>Addon Fee Setup for <%=bldgname%></font>
		</td>
		<td align = "right">
			<input name="close"  style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" type="button" value="Close Window" onclick="javascript:closeWinda();">
		</td>
	</tr>
</table>
<table width="100%" cellspacing="0" cellpadding="2">
<tr bgcolor="#CCCCCC"><td>Admin Fee Amount</td></tr>
</table>
<div style="width:100%;height:130;overflow:auto;background-color:white">
<table width="100%" cellspacing="0" cellpadding="2">
<%
rst1.open "SELECT * FROM building_addonfee WHERE bldgnum='"&bldg&"'", cnn1
do until rst1.eof
	if trim(adfid)=trim(rst1("id")) then adfamount = rst1("addonfee")%>
	<tr style="cursor:hand;" bgcolor="<%if trim(adfid)=trim(rst1("id")) then%>lightgreen<%else%>#FFFFFF<%end if%>" onMouseOver="this.bgColor = 'lightgreen';" onMouseOut="this.bgColor = '<%if trim(adfid)=trim(rst1("id")) then%>lightgreen<%else%>#FFFFFF<%end if%>';" onclick="document.location.href='AddonFeeSetup.asp?bldg=<%=bldg%>&adfid=<%=cint(rst1("id"))%>'"><td><%=formatcurrency(rst1("addonfee"))%></td></tr><%
	rst1.movenext
loop
rst1.close
%>
</table>
</div>
<table>
<tr><td><%if trim(adfid)<>"" then%>Edit Fee<%else%>Add New Fee<%end if%></td><td><input type="Text" name="adfamount" value="<%=adfamount%>"></td></tr>
<tr><td colspan="2"><%if trim(adfid)<>"" then%>&nbsp;<input name="adfsubmit" type="submit" value="Change">&nbsp;<input name="adfsubmit" type="submit" value="Delete"><%end if%>&nbsp;<input name="adfsubmit" type="submit" value="Add New"></td></tr>
</table>
<center><a href="AddonFeeAssoc.asp?bldg=<%=bldg%>&adfid=<%=adfid%>&">Go To Associations</a></center>
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="adfid" value="<%=adfid%>">
</form>
</body>
