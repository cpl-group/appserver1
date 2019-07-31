<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim bldg, ypid, pid, utype
pid = secureRequest("pid")
utype = secureRequest("utype")
bldg = secureRequest("bldg")
ypid = secureRequest("ypid")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim billyear, billperiod, datestart, dateend, utility
if trim(ypid)<>"" then
	rst1.Open "SELECT * FROM billyrperiod WHERE ypid=" & ypid, cnn1
	if not rst1.EOF then
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
<form name="form2" method="post" action="billPeriodSaveG1.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#000000">
<%
dim showWeirdBlackBar
showWeirdBlackBar = false
if allowGroups("Genergy Users") AND showWeirdBlackBar then
%>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td><span class="standardheader"><a href="index.asp" target="main" class="breadcrumb" style="text-decoration:none;"><img src="images/aro-left-000.gif" align="left" width="13" height="13" border="0">Utility Manager Setup</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="frameset.asp" target="main" class="breadcrumb" style="text-decoration:none;">Update Meters</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="portfolioviewG1.asp" target="main" class="breadcrumb" style="text-decoration:none;">Set Up Portfolios</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="regionView.asp" target="main" class="breadcrumb" style="text-decoration:none;">Set Up Rates</a></span></td>
  </tr>
  </table>
<%end if%>
  </td>
</tr>
<tr bgcolor="#3399cc">
	<td><span class="standardheader">
		<%if trim(ypid)<>"" then%>
      Update Bill Period | <span style="font-weight:normal;"><a href="portfolioeditG1.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingeditG1.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a></span>
		<%else%>
			Add New Bill Period | <span style="font-weight:normal;"><a href="portfolioeditG1.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingeditG1.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a></span>
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
		<input type="button" name="cancel" value="Cancel" onclick="location='billPeriodViewG1.asp?pid=<%=pid%>&bldg=<%=bldg%>'" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
	</td>
</tr>
</table>
	<table cellpadding="3" cellspacing="0" border="0"><tr><td>
<%if ypid <> "" then%>
	<%strsql = "SELECT bp.datestart, bp.dateend, billingname, tenantnum, bp.id FROM Billyrperiod b, Billyrperiod_partial bp, tblleasesutilityprices lup, tblleases l WHERE b.ypid=bp.ypid and lup.billingid=l.billingid and bp.lid=lup.leaseutilityid and b.ypid="&ypid&" ORDER BY l.billingname"
	rst1.open strsql, cnn1
	if not rst1.eof then%>
		Partial Bills for this period:</b><%if not(isBuildingOff(bldg)) then%> <a href="#" onclick="window.open('partialBill.asp?pid=<%=pid%>&bldg=<%=bldg%>&utype=<%=utype%>&ypid=<%=ypid%>&pypid=','paritalbills','width=300,height=130,scrollbars=no');">Add</a><%end if%>
		<table border="0" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#dddddd">
			<td><b>Tenant</b></td>
			<td align="center"><b>Date&nbsp;Start</b></td>
			<td align="center"><b>Date&nbsp;End</b></td>
		</tr>
		<%do until rst1.eof	%>
		<tr bgcolor="#ffffff" style="cursor:hand" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" <%if not(isBuildingOff(bldg)) then%>onclick="window.open('partialBill.asp?pid=<%=pid%>&bldg=<%=bldg%>&utype=<%=utype%>&ypid=<%=ypid%>&pypid=<%=rst1("id")%>','paritalbills','width=300,height=130,scrollbars=no');"<%end if%>>
			<td nowrap><%=rst1("billingname")%>&nbsp;[<%=rst1("tenantnum")%>]</td>
			<td align="center"><%=rst1("DateStart")%></td>
			<td align="center"><%=rst1("dateEnd")%></td>
		</tr>
		<%rst1.movenext
		loop
		rst1.close%>
		</table>
	<%else%>
	No partial bills set for this bill period. <%if not(isBuildingOff(bldg)) then%><a href="#" onclick="window.open('partialBill.asp?pid=<%=pid%>&bldg=<%=bldg%>&utype=<%=utype%>&ypid=<%=ypid%>&pypid=','paritalbills','width=300,height=130,scrollbars=no');">Add Parital bill</a><%end if%>
	<%end if%>
<%end if%>
</td></tr></table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="utype" value="<%=utype%>">
<input type="hidden" name="ypid" value="<%=ypid%>">

</form>
</body>
</html>






