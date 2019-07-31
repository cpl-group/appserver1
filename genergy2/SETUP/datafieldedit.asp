<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, lid, meterid
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
lid = secureRequest("lid")
meterid = secureRequest("meterid")

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim field(7), meternum
	rst1.Open "SELECT * FROM datasource d, meters m WHERE d.meterid=m.meterid and m.meterid='"&meterid&"'", cnn1
	if not rst1.EOF then
		field(0) = rst1("fieldname1")
		field(1) = rst1("fieldname2")
		field(2) = rst1("fieldname3")
		field(3) = rst1("fieldname4")
		field(4) = rst1("fieldname5")
		field(5) = rst1("fieldname6")
		field(6) = rst1("fieldname7")
		meternum = rst1("meternum")
	end if
	rst1.close


dim bldgname, portfolioname
if trim(bldg)<>"" then
  rst1.open "SELECT bldgname, name FROM buildings b JOIN portfolio p ON b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
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

dim utilitydisplay
if trim(lid)<>"" then
	rst1.open "select * from tblutility tu join tblleasesutilityprices tlup on tu.utilityid=tlup.utility where leaseutilityid="&lid, cnn1
	utilitydisplay = rst1("utilitydisplay")
	rst1.close
end if
%>
<html>
<head>
<title>Building View</title>
<script>
function openCustomWin(clink, cname, cspec)
{	open(clink, cname, cspec)
}

function meterEdit(meterid)
{	document.location.href = 'meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid='+meterid
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#ffffff">
<form name="form2" method="post" action="datafieldsave.asp">

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#3399cc">
	<td colspan="5">
	<span class="standardheader">
			Update <%=utilitydisplay%> Meter Fields| <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;" target="_parent"><%=portfolioname%></a> &gt; <a href="buildingedit.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;" target="_parent"><%=bldgname%></a> &gt; <a href="tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>" style="color:#ffffff;" target="_parent"><%=billingname%></a></span>
	</span></td>
</tr>
<tr bgcolor="#eeeeee" class="standard">
  <td colspan="5" bgcolor="#dddddd" style="border-bottom:1px solid #ffffff;">
  <table border=0 cellpadding="3" cellspacing="1" class="standard">
  <tr>
    <td align="right">Meter</td>
    <td><%=Meternum%></td>
  </tr>
  </table>
  </td>
</tr>
<tr valign="top">
  <td bgcolor="#eeeeee" width="30%" style="border-bottom:1px solid #cccccc;"> 
  <table border="0" cellpadding="3" cellspacing="1" class="standard">
  <tr bgcolor="#eeeeee">
    <td align="right">Field Name 1</td>
    <td><input type="text" name="fieldname1" value="<%=field(0)%>" size="14"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right">Field Name 2</td>
    <td><input type="text" name="fieldname2" value="<%=field(1)%>" size="14"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right">Field Name 3</td>
    <td><input type="text" name="fieldname3" value="<%=field(2)%>" size="14"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right">Field Name 4</td>
    <td><input type="text" name="fieldname4" value="<%=field(3)%>" size="14"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right">Field Name 5</td>
    <td><input type="text" name="fieldname5" value="<%=field(4)%>" size="14"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right">Field Name 6</td>
    <td><input type="text" name="fieldname6" value="<%=field(5)%>" size="14"></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right">Field Name 7</td>
    <td><input type="text" name="fieldname7" value="<%=field(6)%>" size="14"></td>
  </tr>
  </table>
  </td>
</tr>
<tr bgcolor="#cccccc">
  <td colspan="5" bgcolor="#cccccc" style="border-top:1px solid #999999;">
		<input type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
		  <input type="reset" name="action" value="Cancel" class="standard" onclick="<%
		  if allowGroups("clientOperations") then
		  	response.write "parent.document.location='meteredit.asp?pid="&pid&"&bldg="&bldg&"&meterid=&"&meterid
		  else
			response.write "history.back()"
		  end if
		  %>;" style="border:1px outset #ddffdd;background-color:ccf3cc;">
  </td>
</tr>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
<input type="hidden" name="meterid" value="<%=meterid%>">
</form>
</body>
</html>






