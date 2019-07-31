<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, lid, oldtid
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
lid = secureRequest("lid")
oldtid = secureRequest("oldtid")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, rst2, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim AdminFee, TenantRate, AddonFee, ModifyRate, Coincident, Profile, prtgraph, FullOnPeak, utility, procname
if trim(lid)<>"" then
	rst1.Open "SELECT * FROM tblleasesutilityprices WHERE leaseutilityid='"&lid&"'", cnn1
	if not rst1.EOF then
		AdminFee = rst1("AdminFee")
		TenantRate = rst1("rateTenant")
		AddonFee = rst1("AddonFee")
		ModifyRate = rst1("rateModify")
		Coincident = rst1("Coincident")
		Profile = rst1("loadProfile")
		prtgraph = rst1("prtgraph")
		FullOnPeak = rst1("FullOnPeak")
		utility = rst1("utility")
		procname = rst1("procname")
	end if
	rst1.close
end if

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
<title>Building View</title>
<script>
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}
function meterEdit(meterid)
{	
  document.location = "contentfrm.asp?action='meteredit'&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid="+meterid;
//  document.location.href = 'meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid='+meterid
}
function meterTransfer(lid){	
<!--   if (parent.name != "") { -->
<!--     tgt = parent.contentfrm -->
<!--   } else { -->
<!--     tgt = document; -->
<!--   } -->
  document.location.href = 'TenantMeterTransfer.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=' + lid + '&oldtid=<%=oldtid%>'
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body>
<form name="form2" method="post" action="leaseutilitysave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td><span class="standardheader">
  <a href="index.asp" target="main"><img src="images/aro-left-39c.gif" align="left" width="13" height="13" border="0"></a>
		<%if trim(lid)<>"" then%>
			Update Lease Utility | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a> &gt; <a href="tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>" style="color:#ffffff;"><%=billingname%></a></span>
		<%else%>
			Add New Lease Utility  | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a> &gt; <a href="tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>" style="color:#ffffff;"><%=billingname%></a></span>
		<%end if%>
	</span></td>
   <td align="right"><button id="qmark2" onclick="openCustomWin('help.asp?page=tenanttransfer2','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) Quick Help</button></td>
</tr>
<tr>
  <td align="center" colspan="2" bgcolor="#cccccc">
  <table border=0 cellpadding="0" cellspacing="3">
  <tr>
    <td><img src="images/num_one.gif" alt="Step 1" width="13" height="13" border="0"></td>
    <td><span class="standard">Enter Account Details</span></td>
    <td width="30"><span class="standard">&nbsp;</span></td>
    <td><img src="images/num_two.gif" alt="Step 2" width="13" height="13" border="0"></td>
    <td><span class="standard"><b>Assign New Lease Utilities</b></span></td>
    <td width="30"><span class="standard">&nbsp;</span></td>
    <td><img src="images/num_three.gif" alt="Step 3" width="13" height="13" border="0"></td>
    <td><span class="standard">Transfer Meters</span></td>
  </tr>
  </table>
  </td>
</tr>
</table>
<table border=0 cellpadding="0" cellspacing="0" width="100%">
<tr valign="top">
  <td width="25%" bgcolor="#eeeeee" style="padding:3px;">
  <%
	if trim(oldtid)<>"" then
    rst1.open "select * from tblutility tu join tblleasesutilityprices tlup on tu.utilityid=tlup.utility where billingid="&oldtid, cnn1
    if not rst1.EOF then%>
    <table border=0 cellpadding="2" cellspacing="0" width="100%">
    <tr><td colspan="2"><span class="standard"><b>Previous Account's Lease Utilities</b></span></td></tr>
    <%
    do until rst1.EOF%>
    <tr><td bgcolor="#dddddd" colspan="2"><span class="standard"><%=rst1("utilitydisplay")%></span></td></tr>
    <tr><td><span class="standard">Admin Fee</span></td><td><span class="standard"><%=rst1("adminfee")%></span></td></tr>
    <tr>
      <td><span class="standard">Account Rate</span></td>
      <td>
      <span class="standard">
  		<%  rst2.open "SELECT type FROM ratetypes WHERE id='" & rst1("ratetenant") & "'", getConnect(pid,bldg,"billing")%>
      <%if not rst2.eof then %><%=rst2("type")%><% end if %>
      <% rst2.close %>
      </span>
      </td>
    </tr>
    <tr><td><span class="standard">Add-on Fee</span></td><td><span class="standard"><%=rst1("addonfee")%></span></td></tr>
    <tr><td><span class="standard">Modify Rate</span></td><td><span class="standard"><%=rst1("ratemodify")%></span></td></tr>
    <tr><td><span class="standard">Coincident</span></td><td><span class="standard"><%=rst1("coincident")%></span></td></tr>
    <tr><td><span class="standard">Full On peak</span></td><td><span class="standard"><%=rst1("fullonpeak")%></span></td></tr>
      <%
      rst1.movenext
      loop
    end if
    %>
    </table>
    
    <%
    rst1.close
  end if
  %>
  </td>
  <td bgcolor="#eeeeee">
  <table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#ffffff">
  <tr bgcolor="#eeeeee"><td colspan="3"><span class="standard"><b>New Account's Lease Utilities</b></span></td></tr>
  <tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Utility Type</span></td>
	<td colspan="2"><select name="utility">
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
      <td align="right"><span class="standard">Admin Fee</span></td> 
      <td><input type="text" name="AdminFee" value="<%=AdminFee%>"></td>
      <td><span class="standard"><input type="checkbox" value="1" name="Coincident" <%if Coincident="True" then Response.Write "CHECKED"%>> Lease Expired</span></td>
    </tr>
    <tr bgcolor="#eeeeee">
      <td align="right"><span class="standard">Account Rate</span></td>
      <td><select name="TenantRate">
          <%
          rst1.open "SELECT * FROM ratetypes WHERE regionid in (SELECT region FROM buildings WHERE bldgnum='"& bldg &"')", cnn1
          do until rst1.eof
            %><option value="<%=rst1("id")%>"<%if trim(tenantrate)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("type")%></option><%
            rst1.movenext
          loop
          rst1.close
          %>
        </select>
      </td>
      <td><span class="standard"><input type="checkbox" value="1" name="Profile"> Profile</span></td>
    </tr>
    <tr bgcolor="#eeeeee">
      <td align="right"><span class="standard">Rate Function</span></td>
      <td><select name="procname">
          <%
          rst1.open "SELECT * FROM functiontypes", cnn1
          do until rst1.eof
            %><option value="<%=rst1("id")%>"<%if trim(procname)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("description")%></option><%
            rst1.movenext
          loop
          rst1.close
          %>
        </select>
      </td>
      <td><span class="standard"><input type="checkbox" value="1" name="prtgraph"> prtgraph</span></td>
    </tr>
    <tr bgcolor="#eeeeee">
      <td align="right"><span class="standard">Add on Fee</span></td>
      <td><input type="text" name="AddonFee" value="<%=AddonFee%>"></td>
      <td><span class="standard"><input type="checkbox" value="1" name="FullOnPeak"> Full On Peak</span></td>
    </tr>
    <tr bgcolor="#eeeeee">
      <td align="right"><span class="standard">Modify Rate</span></td>
      <td><input type="text" name="ModifyRate" value="<%=ModifyRate%>"></td>
      <td></td>
    </tr>
    
    <tr bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;"><span class="standard">&nbsp;</span></td>
      
      <td style="border-bottom:1px solid #cccccc;" colspan="2">
        <%if trim(lid)<>"" then%>
          <input type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
        <%else%>
          <input type="submit" name="action" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
        <%end if%>
      </td>
    </tr>
  </table>
  </td>
</tr>
<tr bgcolor="#eeeeee">
  <td>&nbsp;</td>
  <td>
  <%
	if trim(tid)<>"" then
    rst1.open "select * from tblutility tu join tblleasesutilityprices tlup on tu.utilityid=tlup.utility where billingid="&tid, cnn1
    if not rst1.EOF then%>
        <table border=0 cellpadding="2" cellspacing="0" bgcolor="#ffffff" width="100%">
        <tr bgcolor="#dddddd">
        <td><span class="standard"><b>Utility Type</b></span></td>
        <td><span class="standard"><b>Admin Fee</b></span></td>
        <td><span class="standard"><b>Account Rate</b></span></td>
        <td><span class="standard"><b>Add-on Fee</b></span></td>
        <td><span class="standard"><b>Modify Rate</b></span></td>
        <td><span class="standard"><b>Lease Expired</b></span></td>
        <td><span class="standard"><b>Full On peak</b></span></td>
        <td><span class="standard">&nbsp;</span></td>
        </tr>
        <% do until rst1.EOF %>
        <tr bgcolor="#eeeeee">
        <td><span class="standard"><%=rst1("utilitydisplay")%></span></td>
        <td><span class="standard"><%=rst1("adminfee")%></span></td>
        <td><span class="standard">
        <%
        rst2.open "select type from ratetypes where id='" & rst1("ratetenant") & "'", getConnect(pid,bldg,"billing")
        if not rst2.eof then response.write rst2("type")
        rst2.close
        %>
        </span></td>
        <td><span class="standard"><%=rst1("addonfee")%></span></td>
        <td><span class="standard"><%=rst1("ratemodify")%></span></td>
        <td><span class="standard"><%=rst1("coincident")%></span></td>
        <td><span class="standard"><%=rst1("fullonpeak")%></span></td>
        <td><span class="standard"><input type="button" value="Transfer Meter" onclick="meterTransfer(<%=rst1("leaseutilityid")%>);" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></span></td>
        </tr>
        <% 
        rst1.movenext
        loop%>
        </table>
    <%end if
    rst1.close
  end if
  %>
  </td>
</tr>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
<input type="hidden" name="oldtid" value="<%=oldtid%>">
<input type="hidden" name="transfer" value="1">


<!--
[[%
	rst1.Open "SELECT * FROM meters WHERE leaseutilityid='"&lid&"'", cnn1
	if not rst1.EOF then%]]
		[[BR]][[table width="100%" border="0" cellpadding="3" cellspacing="1"]]
		[[tr bgcolor="#cccccc"]]
			[[td]][[span class="standard"]][[b]]Meter[[/b]][[/span]][[/td]]
			[[td]][[span class="standard"]][[b]]Multiplier[[/b]][[/span]][[/td]]
			[[td]][[span class="standard"]][[b]]Location[[/b]][[/span]][[/td]]
		[[/tr]]

		[[%do until rst1.EOF%]]
		[[tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="meterEdit([[%=rst1("meterid")%]]);"]]
			[[td]][[span class="standard"]][[%=rst1("meternum")%]][[/span]][[/td]]
			[[td]][[span class="standard"]][[%=rst1("multiplier")%]][[/span]][[/td]]
			[[td]][[span class="standard"]][[%=rst1("location")%]][[/span]][[/td]]
		[[/tr]]
		
		[[%rst1.movenext
		loop%]]
	[[/table]]
	[[%
	else
		Response.Write "[[BR]]There are no meters set up for this Lease Utility.[[br]]"
	end if
	rst1.close
%]]
[[input type="button" value="Add Meter" onclick="meterEdit('');" id=1 name=1]]
-->
</form>
</body>
</html>






