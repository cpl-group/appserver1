<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, ctid
pid = secureRequest("pid")
bldg = secureRequest("bldg")
ctid = secureRequest("ctid")
'dim DBmainmodIP

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
if trim(bldg)<>"" then
  cnn1.open getLocalConnect(bldg) 
 ' DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."
else 
  cnn1.open getMainConnect(pid)
end if

dim name, address, city, state, zip, phone, fax, email, administrative, m_report, submeter_bills
if trim(ctid)<>"" then
	rst1.Open "SELECT * FROM contacts WHERE id='"&ctid&"'", cnn1
	if not rst1.EOF then
		name = rst1("name")
		address = rst1("address")
		city = rst1("city")
		state = rst1("state")
		zip = rst1("zip")
		phone = rst1("phone")
		fax = rst1("fax")
		email = rst1("email")
		administrative = rst1("administrative")
		m_report = rst1("m_report")
		submeter_bills = rst1("submeter_bills")
	end if
	rst1.close
end if

dim bldgname, portfolioname, breadcrumbtrail
if trim(bldg)<>"" then 
	rst1.Open "SELECT * FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	bldgname = rst1("bldgname")
	portfolioname = rst1("name")
	rst1.close
	breadcrumbtrail = "<a href='portfolioedit.asp?pid=" & pid & "' style='color:#ffffff;'>" & portfolioname & "</a> &gt; <a href='buildingedit.asp?pid=" & pid & "&bldg=" & bldg & "' style='color:#ffffff'>" & bldgname & "</a>"
else
	rst1.open "select name from portfolio where id='" & pid & "'", cnn1
	portfolioname = rst1("name")
	breadcrumbtrail = "<a href=""portfolioedit.asp?pid=" & pid & """ style=""color:#ffffff;"">" & portfolioname & "</a>"
	rst1.close
end if
%>
<html>
<head>
<title>Contact View</title>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body>
<form name="form2" method="post" action="contactsave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td colspan="2"><span class="standardheader">
		<%if trim(ctid)<>"" then%>
			Update Contact | <span style="font-weight:normal;"><%=breadcrumbtrail%></span>
		<%else%>
			Add New Contact | <span style="font-weight:normal;"><%=breadcrumbtrail%></span>
		<%end if%>
	</span></td>
</tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee">
	  <td width="30%" align="right" nowrap><span class="standard">Name</span></td> 
	<td width="70%"><input type="text" name="name" value="<%=name%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">Address</span></td>
	<td><input type="text" name="address" value="<%=address%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">City</span></td>
	<td><input type="text" name="city" value="<%=city%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">State</span></td>
	<td><input type="text" name="state" value="<%=state%>" max="2"></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">Zip Code</span></td>
	<td><input type="text" name="zip" value="<%=zip%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">Phone</span></td>
	<td><input type="text" name="phone" value="<%=phone%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">Fax</span></td>
	<td><input type="text" name="fax" value="<%=fax%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">Email</span></td>
	<td><input type="text" name="email" value="<%=email%>"></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">Include this Contact in 
        <br>Administrative Automation </span></td>
	<td><input type="checkbox" name="administrative" value="1"<%if administrative="True" then response.write " CHECKED"%>></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">Include this Contact in 
        <br>Maintenance Report Automation</span></td>
	<td><input type="checkbox" name="m_report" value="1"<%if m_report="True" then response.write " CHECKED"%>></td>
</tr>
<tr bgcolor="#eeeeee">
	  <td align="right" nowrap><span class="standard">Include this Contact In 
        <br>Billing Notices and Tenant Automation emails</span></td>
	<td><input type="checkbox" name="submeter_bills" value="1"<%if submeter_bills="True" then response.write " CHECKED"%>></td>
</tr>
<tr bgcolor="#eeeeee"> 
	  <td nowrap style="border-bottom:1px solid #cccccc;"><span class="standard">&nbsp;</span></td>
	
	<td style="border-bottom:1px solid #cccccc;">
	<%if not(isBuildingOff(bldg)) then%>
		<%if trim(ctid)<>"" then%>
			<input type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
			<input type="submit" name="action" value="Delete" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
		<%else%>
			<input type="submit" name="action" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
		<%end if%>
	<%end if%>
		<input type="button" name="action" value="Cancel" onclick="document.location='contactView.asp?pid=<%=pid%>&bldg=<%=bldg%>';" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
	</td>
</tr>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="ctid" value="<%=ctid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
</form>
</body>
</html>
