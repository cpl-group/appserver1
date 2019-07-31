<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim bldg, pid, bldgName, adfid, adfamount, adfsubmit

bldg = request("bldg")
adfamount = trim(request("adfamount"))
adfsubmit = request("adfsubmit")
if adfamount="" or not(isnumeric(adfamount)) then adfamount = 0
dim cnn1, cnnMainModule, strsql, rst1, cmd
set cnnMainModule = server.createobject("ADODB.connection")
cnnMainModule.open getConnect(0,bldg,"billing")
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

''submit section
if adfsubmit="Add New" then
	cmd.CommandText = "INSERT INTO Building_AddonFee (bldgnum,AddonFee) Values ('"&bldg&"',"&adfamount&")"
end if
if cmd.CommandText<>"" then
	cmd.activeconnection = cnn1
	'response.write cmd.CommandText
	'response.end
	cmd.execute()
	rst1.open "SELECT max(id) FROM Building_AddonFee", cnn1
	if not rst1.eof then adfid = rst1(0)
	%>
	<script>
	try{
	var aID = opener.document.forms[0].meteraddonID;
	var oOption = opener.document.createElement("OPTION");
	var lindex = aID.options.length;
	aID.add(oOption, lindex);
	oOption.text = "<%=adfamount%>";
	oOption.value = "<%=adfid%>";
	aID.selectedIndex = lindex;
	window.close();
	}catch(exception){alert(exception.description);}
	</script>
	<%
	rst1.close
	response.end
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
<form action="addonfeequick.asp" method="post">
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#3399cc">
		<td>
			<font color='white'>New Addon Fee for <strong><%=bldgname%></strong></font>
		</td>
	</tr>
</table>
<table width="100%" cellspacing="0" cellpadding="2">
<tr bgcolor="#CCCCCC">
	<td>Amount:</td><td><input type="Text" name="adfamount" size="10" value="0"></td>
	<td colspan="2"><input name="adfsubmit" type="submit" value="Add New">&nbsp;<input name="" type="button" value="Cancel" onclick="window.close()"></td></tr>
</table>
<input type="hidden" name="bldg" value="<%=bldg%>">
</form>
</body>
