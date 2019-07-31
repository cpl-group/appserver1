<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if
dim bldg, pid, bldgName, adfid, adfamount, adfsubmit, billingid, action, outmeters, inmeters
pid = request("pid")
bldg = request("bldg")
adfid = request("adfid")
adfamount = trim(request("adfamount"))
adfsubmit = request("adfsubmit")
billingid = trim(request("billingid"))
action = trim(request("action"))
outmeters = trim(request("outmeters"))
inmeters = trim(request("inmeters"))
if adfid="" or not(isnumeric(adfid)) then adfid = 0
if adfamount="" or not(isnumeric(adfamount)) then adfamount = 0
if billingid="" or not(isnumeric(billingid)) then billingid = 0
dim cnn1, cnnMainModule, strsql, rst1, cmd
set cnnMainModule = server.createobject("ADODB.connection")
cnnMainModule.open getConnect(pid,bldg,"billing")
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

''submit section
if action="<" and inmeters<>"" then 'subtract from fee
	cmd.CommandText = "UPDATE MeterPrices SET AddonFee=0 WHERE meterid in ("&inmeters&")"
elseif action=">" and outmeters<>"" then 'add to fee
	cmd.CommandText = "UPDATE MeterPrices SET AddonFee="&adfid&" WHERE meterid in ("&outmeters&")"
end if
if cmd.CommandText<>"" then
	cmd.activeconnection = cnn1
	'response.write cmd.CommandText
	'response.end
	cmd.execute()
end if

strsql = "select bldgname from buildings where bldgnum = '"&bldg&"'"
rst1.open strsql, cnn1
if not rst1.eof then
	bldgName = rst1("bldgname")
end if
rst1.close
dim leasefilter
if billingid<>"0" and billingid<>"" then leasefilter = " leaseutilityid in (SELECT leaseutilityid FROM tblleasesutilityprices WHERE billingid="&billingid&") AND "
%>	
<link rel="Stylesheet" href="setup.css" type="text/css">
<title>Addon Fee Associations for <%=bldgname%></title>
<script>
function closeWinda(){
window.close()
}

function selectAll(sel){
	for (var i=sel.length-1;i>=0;i--){
		sel[i].selected = true;
	}
}
</script>
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form action="addonfeeAssoc.asp" method="get">
<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
	<tr bgcolor="#3399cc">
		<td>
			<font color='white'>Addon Fee Associations for <%=bldgname%></font>
		</td>
		<td align = "right">
			<input name="close"  style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" type="button" value="Close Window" onclick="javascript:closeWinda();">
		</td>
	</tr>
</table>
<table align="center" width="100%">
<tr><td align="center">Lease Filter:<br>
	<select name="billingid" onchange="forms[0].submit()">
		<option value="0">All Meters</option><%
		rst1.open "SELECT * FROM tblleases WHERE bldgnum='"&bldg&"'", cnn1
		'if billingid=0 and not rst1.eof then billingid = trim(rst1("billingid"))
		do until rst1.eof%>
			<option value="<%=rst1("billingid")%>" <%if cint(billingid)=cint(rst1("billingid")) then response.write "SELECTED"%>><%=rst1("billingname")%></option><%
			rst1.movenext
		loop
		rst1.close%>
	</select>
</td>
<td align="center">Admin Fees:<br>
	<select name="adfid" onchange="forms[0].submit()">
		<option value="0">No Admin Fee</option><%
		rst1.open "SELECT * FROM building_addonfee WHERE bldgnum='"&bldg&"'", cnn1
		do until rst1.eof%>
			<option value="<%=rst1("id")%>" <%if trim(adfid)=trim(rst1("id")) then response.write "SELECTED"%>><%=formatcurrency(rst1("addonfee"))%></option><%
			rst1.movenext
		loop
		rst1.close%>
	</select>
</td></tr>
</table>
<table align="center">
<tr><td align="center"><a href="#" onclick="selectAll(document.forms[0].outmeters)">select all</a><br>
	<select name="outmeters" size="10" multiple><optgroup label="Other Meters                               "></optgroup><%
	rst1.open "SELECT *, isnull(ba.addonfee,0) as adfamount FROM meters m LEFT JOIN MeterPrices mp ON mp.meterid=m.meterid LEFT JOIN building_addonfee ba ON ba.id=mp.addonfee WHERE m.bldgnum='"&bldg&"' AND "&leasefilter&" isnull(mp.Addonfee,0)<>"&adfid&" and ("&adfid&"<>0 or mp.Addonfee in (SELECT id FROM Building_AddonFee)) ORDER BY meternum", cnn1
	do until rst1.eof
		if trim(adfid)=trim(rst1("addonfee")) then adfamount = rst1("addonfee")%>
		<option value="<%=rst1("meterid")%>"><%=rst1("meternum")%> (<%=formatcurrency(rst1("adfamount"))%>)</option><%
		rst1.movenext
	loop
	rst1.close%>
	</select>
</td>
<td align="center">
	<input type="submit" name="action" value=" &gt; " style="border: 1px solid; font-size: 13px; font-weight: bold;"><br>
	<input type="submit" name="action" value=" &lt; " style="border: 1px solid; font-size: 13px; font-weight: bold;">
</td>
<td align="center"><a href="#" onclick="selectAll(document.forms[0].inmeters)">select all</a><br>
	<select name="inmeters" size="10" multiple><optgroup label="Meters in Fee                              "></optgroup><%
	rst1.open "	SELECT *, isnull(ba.addonfee,0) as adfamount FROM meters m LEFT JOIN MeterPrices mp ON mp.meterid=m.meterid LEFT JOIN building_addonfee ba ON ba.id=mp.addonfee WHERE m.bldgnum='"&bldg&"' AND "&leasefilter&" (isnull(mp.Addonfee,0)="&adfid&" or ("&adfid&"=0 and mp.Addonfee not in (SELECT id FROM Building_AddonFee))) ORDER BY meternum", cnn1
	do until rst1.eof
		if trim(adfid)=trim(rst1("addonfee")) then adfamount = rst1("addonfee")%>
		<option value="<%=rst1("meterid")%>"><%=rst1("meternum")%> (<%=formatcurrency(rst1("adfamount"))%>)</option><%
		rst1.movenext
	loop
	rst1.close%>
	</select>
	</td>
</tr>
</table>
<center><a href="AddonFeeSetup.asp?bldg=<%=bldg%>&adfid=<%=adfid%>&">Go To Add/Remove</a></center>
<input type="hidden" name="bldg" value="<%=bldg%>">
</form>
</body>
