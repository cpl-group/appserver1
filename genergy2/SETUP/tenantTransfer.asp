<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, newtid
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
newtid = secureRequest("newtid")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql, rst2
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim bldgname, portfolioname
if trim(bldg)<>"" then
  rst1.open "SELECT bldgname, name FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
'	rst1.Open "SELECT bldgname FROM buildings WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
	end if
	rst1.close
end if
if trim(tid)<>"" then
	rst1.Open "SELECT * FROM tblleases WHERE billingid='"&tid&"'", cnn1

	dim tenantnum, flr, sqft, taxexempt, billingname, leaseexpired, interm, intermcharges, startdate
	if not rst1.EOF then
		tenantnum = rst1("tenantnum")
		flr = rst1("flr")
		sqft = rst1("sqft")
		taxexempt = rst1("taxexempt")
		billingname = rst1("billingname")
		leaseexpired = rst1("leaseexpired")
		interm = rst1("interm")
		intermcharges = rst1("intermcharges")
		startdate = rst1("startdate")
		bldg = rst1("bldgnum")
	end if
	rst1.close
end if

%>
<html>
<head>
<title>Building View</title>
<script language="JavaScript">
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}
function leaseUtilityEdit(lid)
{	document.location.href = 'leaseutilityedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid='+lid
}
function meterEdit(meterid,lid)
{	
  if (parent.frames.length > 2) {
    document.location.href = "contentfrm.asp?action='meteredit'&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid+"&meterid="+meterid
  } else {
    document.location.href = "frameset.asp?action='meteredit'&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid+"&meterid="+meterid
  }
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#ffffff">
<form name="form2" method="post" action="tenantxfersave.asp">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr bgcolor="#3399cc">
	<td colspan="2" height="26">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td>
    <span class="standardheader">
    <%if trim(tid)<>"" then%>
      Transfer Account | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a> &gt; <%=billingname%></span>
    <%else%>
      Add New Account to <%=bldgname%>
    <%end if%>
    </span>
    </td>
   <td align="right"><button id="qmark2" onclick="openCustomWin('help.asp?page=tenanttransfer1','Help','width=400,height=500,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) Quick Help</button></td>
  </tr>
  <tr>
    <td colspan="2" align="center" bgcolor="#cccccc">
    <table border=0 cellpadding="0" cellspacing="3">
    <tr>
      <td><img src="images/num_one.gif" alt="Step 1" width="13" height="13" border="0"></td>
      <td><span class="standard"><b>Enter Account Details</b></span></td>
      <td width="30"><span class="standard">&nbsp;</span></td>
      <td><img src="images/num_two.gif" alt="Step 2" width="13" height="13" border="0"></td>
      <td><span class="standard">Assign New Lease Utilities</span></td>
      <td width="30"><span class="standard">&nbsp;</span></td>
      <td><img src="images/num_three.gif" alt="Step 3" width="13" height="13" border="0"></td>
      <td><span class="standard">Transfer Meters</span></td>
    </tr>
    </table>
    </td>
  </tr>
  </table>
	</td>
</tr>
<tr bgcolor="#fffff" valign="top">
  <td bgcolor="#eeeeee" width="35%" style="border-bottom:1px solid #cccccc;">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#eeeeee">
    <td colspan="2"><span class="standard"><b>Previous Account</b></span></td> 
  </tr>
  <tr bgcolor="#eeeeee">
    <td><span class="standard">Account Number</span></td> 
    <td><%=tenantnum%></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td><span class="standard">Billing Name</span></td>
    <td><%=billingname%></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td><span class="standard">Start Date</span></td>
    <td><%=startdate%></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td><span class="standard">Floor</span></td>
    <td><%=flr%></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td><span class="standard">SQFT</span></td>
    <td><%=sqft%></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td><span class="standard">Tax Exempt</span></td>
    <td><%if taxexempt="True" then Response.Write "Yes" else Response.Write "No" end if%></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td><span class="standard">Interm Charges</span></td>
    <td><%if interm="True" then Response.Write "Yes:" else Response.Write "None"%>&nbsp;(<%=intermcharges%>)</td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td><span class="standard">Lease Offline</span></td>
    <td><input type="checkbox" value="1" name="oldleaseexpired" <% if leaseexpired="True" then Response.Write "CHECKED"%>></td>
  </tr>
  </table>  
  
  </td>
  <td bgcolor="#ffffff" style="border-left:1px outset #ffffff;border-bottom:1px solid #cccccc;">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#eeeeee">
    <td colspan="2" width="65%"><span class="standard">&nbsp;&nbsp;<b>New Account</b></span></td> 
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Account Number</span></td> 
    <td><input type="text" name="tenantnum" maxlength="12" value=""></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Billing Name</span></td>
    <td><input type="text" name="billingname" value=""></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Start Date</span></td>
    <td><input type="text" name="startdate" value=""></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Floor</span></td>
    <td><input type="text" name="flr" value="<%=flr%>"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">SQFT</span></td>
    <td><input type="text" name="sqft" value="<%=sqft%>"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Tax Exempt</span></td>
    <td><input type="checkbox" value="1" name="taxexempt"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Interim Charges</span></td>
    <td><input type="checkbox" value="1" name="interm">&nbsp;<input type="text" name="intermcharges" value="<%=intermcharges%>"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Lease Offline</span></td>
    <td><input type="checkbox" value="0" name="leaseexpired"></td>
  </tr>
  <tr bgcolor="#eeeeee"> 
    <td><span class="standard">&nbsp;</span></td>
    <td><input type="submit" name="action" value="Continue" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"></td>
  </tr>
  </table>  
  </td>
</tr>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">

</form>
</body>
</html>