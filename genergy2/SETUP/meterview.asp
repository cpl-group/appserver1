<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, lid, util
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
lid = secureRequest("lid")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
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
<title>Building View</title>
<script>
function leaseUtilityEdit(lid)
{	parent.editfrm.location.href = 'leaseutilityedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid='+lid
}
function meterEdit(meterid)
{	
  if (meterid == "") {
    parent.location.href = 'meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid='+meterid
  } else {
    parent.editfrm.location.href = 'meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid='+meterid
  }
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body topmargin=0 leftmargin=0>
<form name="form2" method="post">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td colspan="2">
<%
if trim(tid)<>"" then
	rst1.Open "SELECT * FROM meters WHERE leaseutilityid='"&lid&"' order by meternum", cnn1 %>
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td bgcolor="#3399cc">
    <table border=0 cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td><span class="standardheader">Meters  | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;" target="_parent"><%=portfolioname%></a> &gt; <a href="buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;" target="_parent"><%=bldgname%></a> &gt; <a href="tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>" style="color:#ffffff;" target="_parent"><%=billingname%></a></span></span></td>
      <td align="right"><%if not(isBuildingOff(bldg)) then%><a href="javascript:meterEdit('');"><img src="images/add-meter.gif" width="64" height="15" border="0"></a><%end if%></td>
    </tr>
    </table>
    </td>
  </tr>
	<% if not rst1.EOF then%>
		<%do until rst1.EOF%>
    <tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="meterEdit(<%=rst1("meterid")%>);">
      <td><%=rst1("meternum")%></td>
    </tr>
		<%rst1.movenext
		loop%>
	<tr><td><span class="standard" style="border-bottom:1px solid #ffffff;"><%if not(isBuildingOff(bldg)) then%><input type="button" value="Add Meter" onclick="meterEdit('');" id=1 name=1 class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><%end if%></span></td></tr>
	</table>
	<%
	else %>
		<tr><td>There are no meters set up for this account.</td></tr>
		</table>
	<%end if
	rst1.close
else
	Response.Write "<BR>There are no meters set up for this building."
end if
%>
</form>
</body>
</html>