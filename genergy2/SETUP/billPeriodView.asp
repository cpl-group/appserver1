<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, utype, action, sqlstr, historic
pid = secureRequest("pid")
bldg = secureRequest("bldg")
utype = secureRequest("utype")
if lcase(request("historic"))="true" then historic=true else historic=false
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim bldgname, portfolioname
if trim(bldg)<>"" then
	rst1.Open "SELECT bldgname, name, strt FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("strt")
		portfolioname = rst1("name")
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Bill Period View</title>
<script>
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}
function billPeriodEdit(ypid)
{	document.location = 'billPeriodEdit.asp?pid=<%=pid%>&bldg=<%=bldg%>&utype=<%=utype%>&ypid='+ypid
}
function copyBillPeriod(){
	open('copyBPFromBuilding.asp?bldg=<%=bldg%>','' ,'width=430, height=187, scrollbars=no');
}
function updateTripcode(value, action){
	open('updateTripcode.asp?bldg=<%=bldg%>&pid=<%=pid%>&uid=<%=utype%>&tripcode='+value+'&action='+action,'UpdateTrip','width=200, height=100, scrollbars=no');
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<FORM>

<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
<tr>
  <td bgcolor="#000000">
<%if allowGroups("Genergy Users") AND false = true then%>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td><span class="standardheader"><a href="index.asp" target="main" class="breadcrumb" style="text-decoration:none;"><img src="images/aro-left-000.gif" align="left" width="13" height="13" border="0">Utility Manager Setup</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="frameset.asp" target="main" class="breadcrumb" style="text-decoration:none;">Update Meters</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="portfolioview.asp" target="main" class="breadcrumb" style="text-decoration:none;">Set Up Portfolios</a></span></td>
    <td width="12"><span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
    <td><span class="standardheader"><a href="regionView.asp" target="main" class="breadcrumb" style="text-decoration:none;">Set Up Rates</a></span></td>
  </tr>
  </table>
<%end if%>
  </td>
</tr>
<tr>
  <td bgcolor="#3399CC">
  <table border=0 cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td><span class="standardheader">Manage Bill Periods | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a></span></span></td>
    <td align="right">
	<label style="border:1px solid #6699cc; color:white; font-weight: bold; border-bottom-style: solid;cursor:hand" onClick="
	document.location='billperiodView.asp?pid=<%=pid%>&bldg=<%=bldg%>&utype=<%=utype%>&historic=<%if historic then%>false<%else%>true<%end if%>'
	" onMouseOver="this.style.borderColor='white';" onMouseOut="this.style.borderColor='#6699cc';" type="" src="" value="New Job">&nbsp;<%if historic then%>Hide<%else%>Show<%end if%>&nbsp;Historical&nbsp;Periods&nbsp;</label>
	<button id="qmark2" onClick="openCustomWin('help.asp?page=billperiodview','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) Quick Help</button></td>
  </tr>
  </table>
  </td>
</tr>
<tr>
  <td bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;">
  <table border=0 cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td>
    <span class="standard">
    Show bill periods for:&nbsp;
    <select name="utility" onChange="document.location='billPeriodView.asp?pid=<%=pid%>&bldg=<%=bldg%>&utype='+this.value+'&historic=<%=historic%>'">
    <option value="">Select Utility Type</option>
    <%
    rst1.open "SELECT * FROM tblutility ORDER BY utilitydisplay", cnn1
    do until rst1.eof
      %><option value="<%=rst1("utilityid")%>"<%if trim(utype)=trim(rst1("utilityid")) then response.write " SELECTED"%>><%=rst1("utilitydisplay")%></option><%
      rst1.movenext
    loop
    rst1.close
    %>
    </select>
    </span>
    </td>
    <td align="right">
	<%if not(isBuildingOff(bldg)) then%>
		<input type="button" value="Copy Bill Periods From Building" onClick="copyBillPeriod();" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><input type="button" value="Add Bill Period" onClick="billPeriodEdit('');" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
	<%end if%>
	</td>
  </tr>
  </table>
  </td>
</tr>
</table>
<%
'response.write cnn1 &"<BR>" & bldg
'response.end 
sqlstr = "SELECT *  FROM billyrperiod byp left join super_tripcodes st on st.bldgnum = byp.bldgnum and st.uid=byp.utility WHERE utility='"&utype&"' and byp.bldgnum='"&bldg&"' "
if not historic then sqlstr = sqlstr & " and (billyear = "& year(date()) + 1 & " or billyear = "& year(date())& " or billyear = " & year(date()) - 1 &") "
sqlstr = sqlstr & " ORDER BY billyear desc, billperiod desc"
rst1.Open sqlstr, cnn1


if not rst1.EOF then

if isnull(rst1("tripcode"))then 
	action = "save" 
else
	action = "update"
end if 
%>
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
	<tr bgcolor="#FFFFFF">
      <td colspan=5><span class="standard"><b>Trip code:</b></span>&nbsp;<input name="tripcode" type="text" size="5" maxlength="4" value="<%=rst1("tripcode")%>"> 
        <%if not(isBuildingOff(bldg)) then%><input name="UpdateTrip" type="button" value="Update Tripcode" onClick="updateTripcode(tripcode.value,'<%=action%>')"><%end if%>
	</td>
	</tr>
	<tr bgcolor="#dddddd">
		<td><span class="standard"><b>Year</b></span></td>
		<td><span class="standard"><b>Period</b></span></td>
		<td><span class="standard"><b>Date Start</b></span></td>
		<td><span class="standard"><b>Date End</b></span></td>
		<td><span class="standard"><b>Date Trans</b></span></td>
	</tr>

	<%do until rst1.EOF%>
	<tr bgcolor="#ffffff" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" onClick="billPeriodEdit('<%=rst1("ypid")%>');">
		<td><span class="standard"><%=rst1("billyear")%></span></td>
		<td><span class="standard"><%=rst1("billperiod")%></span></td>
		<td><span class="standard"><%=rst1("DateStart")%></span></td>
		<td><span class="standard"><%=rst1("dateEnd")%></span></td>
		<td><span class="standard"><%=rst1("dateTrans")%></span></td>
	</tr>
	
	<%rst1.movenext
	loop%>
</table>
<%
else
	if utype<>"" then%>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td>There are no bill periods set up for this utility in <%=bldgname%></td>
  </tr>
  <tr><td><input type="button" value="Add Bill Period" onClick="billPeriodEdit('');" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td></tr>
  </table>
	<%end if
end if
rst1.close
%>
</FORM>
</body>
</html>
