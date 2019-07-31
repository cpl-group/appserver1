<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'N.Ambo 5/19/2008 added this asp page so that user can add new bill periods from the 'Historical Data Entry' screen in G1Console
'This page was copied from the original page used for bill period entry in the utility manager and has been slightly modified

if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim bldg, ypid, pid, utype, billyear, billperiod

pid = request("pid")
bldg = request("bldgNUm")
utype = request("utilityid")
if instr(request("bperiod"),"/")>0 then
	billyear = split(request("bperiod"),"/")(1)
	billperiod = split(request("bperiod"),"/")(0)
else
	billyear = request("byear")
	billperiod = request("bperiod")
end if


dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim datestart, dateend, utility
'if trim(ypid)<>"" then
if request("bperiod") <> "" then
	'strsql =  "SELECT * FROM billyrperiod WHERE ypid=" & ypid
	rst1.open "SELECT * FROM billyrperiod WHERE bldgnum='" &bldg& "' and billyear='"&billyear& "' and billperiod='"&billperiod& "' and utility=" &utype, cnn1
	response.Write strsql
	if not rst1.EOF then
		ypid = rst1("ypid")
		billyear = rst1("billyear")
		billperiod = rst1("billperiod")
		datestart = rst1("datestart")
		dateend = rst1("dateend")
		utility = rst1("utility")
	end if
	rst1.close
end if

dim bldgname, portfolioname
if trim(bldg)<>"" then
	rst1.Open "SELECT bldgname, name FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Building View</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body>
<form name="form2" method="post" action="billPeriodSave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">

<tr bgcolor="#3399cc">
	<td><span class="standardheader">
		<%if trim(ypid)<>"" then%>
      Update Bill Period  
		<%else%>
			Add New Bill Period  
		<%end if%>
	</span></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Year</span></td>
	<td><input type="text" name="billyear" value="<%=billyear%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Period</span></td>
	<td><input type="text" name="billperiod" value="<%=billperiod%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Start Date</span></td>
	<td><input type="text" name="datestart" value="<%=datestart%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">End Date</span></td>
	<td><input type="text" name="dateend" value="<%=dateend%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Utility</span></td>
	<td>
		<select name="utility">
			<%
			rst1.open "SELECT * FROM tblutility ORDER BY utilitydisplay", cnn1
			do until rst1.eof
				%><option value="<%=rst1("utilityid")%>"<%if trim(utility)=trim(rst1("utilityid")) then response.write " SELECTED"%>><%=rst1("utilitydisplay")%></option><%
				rst1.movenext
			loop
			rst1.close
			%>
		</select>
	</td>
</tr>
<tr bgcolor="#eeeeee"> 
	<td style="border-bottom:1px solid #cccccc;"><span class="standard">&nbsp;</span></td>
	
	<td style="border-bottom:1px solid #cccccc;">
	<%if not(isBuildingOff(bldg)) then%>
		<%if trim(ypid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
			<input type="submit" name="action" value="Delete" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
		<%end if%>
	<%end if%>
		<input type="button" name="cancel" value="Cancel" onclick="location='historicaldataentry.asp?pid=<%=pid%>&bldgNum=<%=bldg%>'" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
	</td>
</tr>
</table>
	
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="utype" value="<%=utype%>">
<input type="hidden" name="ypid" value="<%=ypid%>">

</form>
</body>
</html>






