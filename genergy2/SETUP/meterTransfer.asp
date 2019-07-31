<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not( _
	checkgroup("Genergy Users")<>0 _
	or checkgroup("clientOperations")<>0 _
	) then%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")

dim pid, bldg, tid, lid, transfermeter
pid = request("pid")
bldg = request("bldg")
tid = request("tid")
lid = request("lid")
transfermeter = request("transfermeter")

dim bldgfilter
bldgfilter = request("bldgfilter")

dim mwhere
if trim(bldgfilter)<>"" then
	mwhere = "and b.bldgnum='"&bldgfilter&"'"
end if
%>
<html>
<head>
<title>Meter Transfers</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<script>
function reloadFilter(bldg,transfermeter)
{	document.location = 'meterTransfer.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&bldgfilter='+bldg+'&transfermeter='+transfermeter
}
</script>
<body>
<form name="form2" method="post" action="meterTransferSave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td><span class="standardheader">Transfer Meter</span></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#eeeeee">
	<td align="right" width="35%"><span class="standard">Select meter:</span></td>
	<td><select name="transfermeter" onchange="reloadFilter('<%=bldgfilter%>',this.value)">
			<%
			rst1.open "SELECT * FROM meters m INNER JOIN buildings b on b.bldgnum=m.bldgnum WHERE leaseutilityid in (SELECT lup.leaseutilityid FROM tblleasesutilityprices lup INNER JOIN tblLeases l ON lup.billingid=l.billingid WHERE bldgnum='"&bldg&"' and leaseutilityid<>"&lid&") "&mwhere&" order by meternum", cnn1
			do until rst1.eof
				%><option value="<%=rst1("meterid")%>"<%if trim(rst1("meterid"))=trim(transfermeter) then response.write " SELECTED"%>><%=rst1("meternum")%></option><%
				rst1.movenext
			loop
			rst1.close
			%>
		</select>
	</td>
</tr>
<%if trim(transfermeter)<>"" then%>
<tr bgcolor="#eeeeee">
	<td align="right" width="35%"><span class="standard">Include data back to bill period:</span></td>
	<td><select name="bybp">
			<%
			rst1.open "SELECT billyear, billperiod FROM billyrperiod byp INNER JOIN meters m ON byp.bldgnum=m.bldgnum WHERE m.meterid="&transfermeter&" and byp.datestart<getdate() ORDER BY billyear desc, billperiod desc", cnn1
			do until rst1.eof
				%><option value="<%=rst1("billyear")%>|<%=rst1("billperiod")%>"><%=rst1("billyear")%>|<%=rst1("billperiod")%></option><%
				rst1.movenext
			loop
			rst1.close
			%>
		</select>
	</td>
</tr>
<%end if%>
<tr bgcolor="#cccccc"> 
	<td><span class="standard">&nbsp;</span></td>
	
	<td>
		<%if "l"<>"" then%>
			<input type="submit" name="action" value="Transfer" class="standard">
		<%end if%>
	</td>
</tr>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
</form>
</body>
</html>
