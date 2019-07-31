<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, lid, oldtid, transfermeter
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
lid = secureRequest("lid")
oldtid = secureRequest("oldtid")
transfermeter = secureRequest("transfermeter")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, rst2, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim bldgname, portfolioname
if trim(bldg)<>"" then
  rst1.open "SELECT bldgname, name FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
	end if
	rst1.close
end if

dim billingname
if trim(bldg)<>"" then
  rst1.open "SELECT billingname FROM tblleases WHERE billingid='"&tid&"'", cnn1
	if not rst1.EOF then
		billingname = rst1("billingname")
	end if
	rst1.close
end if

%>
<html>
<head>
<title>Meter Transfers</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<script language="JavaScript" type="text/javascript">
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}
function reloadFilter(transfermeter)
{	document.location = 'TenantMeterTransfer.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&oldtid=<%=oldtid%>&transfermeter='+transfermeter
}
function meterEdit(meterid,lid)
{	
  document.location.href = "contentfrm.asp?action=meteredit&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid+"&meterid="+meterid
}
</script>
<body>
<form name="form2" method="post" action="meterTransferSave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td><span class="standardheader">
  <a href="index.asp" target="main"><img src="images/aro-left-39c.gif" align="left" width="13" height="13" border="0"></a>
	Transfer Meters | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a> &gt; <a href="tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>" style="color:#ffffff;"><%=billingname%></a></span>
	</span></td>
  <td align="right"><button id="qmark2" onclick="openCustomWin('help.asp?page=tenanttransfer3','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) Quick Help</button></td>
</tr>
<tr>
  <td align="center" colspan="2" bgcolor="#cccccc">
  <table border=0 cellpadding="0" cellspacing="3">
  <tr>
    <td><img src="images/num_one.gif" alt="Step 1" width="13" height="13" border="0"></td>
    <td><span class="standard">Enter Account Details</span></td>
    <td width="30"><span class="standard">&nbsp;</span></td>
    <td><img src="images/num_two.gif" alt="Step 2" width="13" height="13" border="0"></td>
    <td><span class="standard">Assign New Lease Utilities</span></td>
    <td width="30"><span class="standard">&nbsp;</span></td>
    <td><img src="images/num_three.gif" alt="Step 3" width="13" height="13" border="0"></td>
    <td><span class="standard"><b>Transfer Meters</b></span></td>
  </tr>
  </table>
  </td>
</tr>
</table>
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr valign="top" bgcolor="#eeeeee">
  <td width="25%" bgcolor="#eeeeee" style="padding:3px;">
  <!--lid = [[%=lid%]][[br]]-->
  <%
	if trim(oldtid)<>"" then
    rst1.open "select * from tblutility tu join tblleasesutilityprices tlup on tu.utilityid=tlup.utility where billingid="&oldtid, cnn1
    if not rst1.EOF then%>
      <span class="standard">
      <b>Previous Account's Meters</b>
      <br><br>
      <% do until rst1.EOF %>
      <%=rst1("utilitydisplay")%><br>
      <%
      rst2.Open "SELECT * FROM meters WHERE leaseutilityid='"&rst1("leaseutilityid")&"'", cnn1
      if not rst2.EOF then
        do until rst2.EOF
        %>
        <li><%=rst2("meternum")%><br>
        <%
        rst2.movenext
        loop
      end if
      rst2.close

      rst1.movenext
      loop
    end if
    rst1.close
  end if
  %>
  </td>
	<td valign="top">
	    <span class="standard">
	    <b>Transfer to New Account</b><br><br>
	    <select name="transfermeter" size="4" multiple><!-- onchange="reloadFilter(this.value)"> -->
			<%
			rst1.open "SELECT * FROM meters WHERE leaseutilityid in (SELECT lup.leaseutilityid FROM tblleasesutilityprices lup INNER JOIN tblLeases l ON lup.billingid=l.billingid WHERE bldgnum='"&bldg&"' and leaseutilityid<>"&lid&") and online=1", cnn1
			do until rst1.eof
				%><option value="<%=rst1("meterid")%>"<%if trim(rst1("meterid"))=trim(transfermeter) then response.write " SELECTED"%>><%=rst1("meternum")%></option><%
				rst1.movenext
			loop
			rst1.close
			%>
		</select>&nbsp;
      Include data back to bill period:
      <select name="bybp">
      <option value="0|0">Entire History</option>
			<%
			rst1.open "SELECT distinct billyear, billperiod FROM billyrperiod byp INNER JOIN meters m ON byp.bldgnum=m.bldgnum WHERE byp.bldgnum='"&bldg&"' and byp.datestart<getdate() ORDER BY billyear desc, billperiod desc", cnn1
			do until rst1.eof
				%><option value="<%=rst1("billyear")%>|<%=rst1("billperiod")%>"><%=rst1("billyear")%> | Period <%=rst1("billperiod")%></option><%
				rst1.movenext
			loop
			rst1.close
			%>
		</select>
		</span>
	</td>
</tr>
<tr bgcolor="#cccccc"> 
	<td><span class="standard">&nbsp;</span></td>
	
	<td>
		<%if "l"<>"" then%>
			<input type="submit" name="action" value="Transfer" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
		<%end if%>
	</td>
</tr>
</table>
<% if lid<>"" then
rst2.Open "SELECT * FROM meters WHERE leaseutilityid='"&lid&"'", cnn1
  if not rst2.EOF then%>
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#dddddd">
    <td><span class="standard"><b>Meter</b></span></td>
    <td><span class="standard"><b>Start Date</b></span></td>
    <td><span class="standard"><b>Date Off</b></span></td>
    <td><span class="standard"><b>Last Read</b></span></td>
    <td><span class="standard"><b>Location</b></span></td>
    <td><span class="standard"><b>Floor</b></span></td>
    <td><span class="standard"><b>Riser</b></span></td>
  </tr>

  <%do until rst2.EOF%>
  <tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="meterEdit(<%=rst2("meterid")%>,<%=lid%>);">
    <td><span class="standard"><%=rst2("meternum")%></span></td>
    <td><span class="standard"><%=rst2("datestart")%></span></td>
    <td><span class="standard"><%=rst2("dateoffline")%></span></td>
    <td><span class="standard"><%=rst2("datelastread")%></span></td>
    <td><span class="standard"><%=rst2("location")%></span></td>
    <td><span class="standard"><%=rst2("floor")%></span></td>
    <td><span class="standard"><%=rst2("riser")%></span></td>
  </tr>
  
  <%rst2.movenext
  loop%>
  </table>
  <% 
  end if
rst2.close
%>
<% end if%>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
<input type="hidden" name="oldtid" value="<%=oldtid%>">
</form>
</body>
</html>
