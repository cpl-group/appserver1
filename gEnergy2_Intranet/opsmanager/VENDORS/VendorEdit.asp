<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<%
dim cnn, rst, strsql
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

dim company, vendor, name, address_1, address_2, city, state, zip, telephone, fax_number, companyname, contact_1_name
company = trim(request("company"))
vendor = trim(request("vendor"))

rst.Open "select * from companycodes where code = '"&company&"'", cnn
if not rst.eof then
	companyname = rst("name")
end if
rst.close

if company<>"" and vendor<>"" then
	rst.open "SELECT * FROM "&company&"_MASTER_APM_VENDOR WHERE vendor='"&vendor&"'", cnn
	if not rst.eof then
		name			= rst("name")
		address_1		= rst("address_1")
		address_2		= rst("address_2")
		city	 		= rst("city")
		state			= rst("state")
		zip				= rst("zip")
		telephone		= rst("telephone")
		fax_number		= rst("fax_number")
		contact_1_name	= rst("contact_1_name")
	end if
	rst.close
end if

if vendor="" then vendor="0"
dim ticket
set ticket = New tickets
ticket.Label="Vendor"
ticket.Note="Vendor Master Ticket "
ticket.requester = "JOBLOGADMIN"
ticket.department = "OPERATIONS"
ticket.userid = "JOBLOGADMIN"
if vendor<>"0" then ticket.findtickets "vendorid", vendor
%>
<html>
<head>
<title>Genergy Vendors</title>
<script>
function closewin(){
	window.close()
}
function copymain(){
	document.form1.bAddress_street.value = document.form1.cAddress_street.value
	document.form1.bAddress_floor.value = document.form1.cAddress_floor.value 
	document.form1.bCity.value = document.form1.cCity.value 
	document.form1.bstate.value = document.form1.cstate.value  
	document.form1.bzip.value = document.form1.czip.value 
}
function screencompany(company){
	document.location.href="VendorEdit.asp?company="+company	
}

function checkform(frm){
	var err = "";
	if(frm.name.value=='') err+="Select company name\n";
	if(frm.address_1.value=='') err+="No company address entered\n";
	if(frm.city.value=='') err+="No company city entered\n";
	if(frm.state.value=='') err+="No company state entered\n";
	if(frm.zip.value=='') err+="No company zip code entered\n";
	if(err=="") 
		return true;
	else alert(err);
		return false;
}
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">		
</head>
<body bgcolor="#eeeeee" marginwidth=0 marginheight=0 topmargin=0 leftmargin=0>
<form name="form1" method="post" action="VendorSave.asp" onsubmit="return(checkform(this))">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr valign="top" bgcolor="#6699cc"><td><span class="standardheader">Vendor for <%=companyname%></span></td><td align="right" width="1%"><%if vendor<>"0" then ticket.MakeButton%></td></tr>
<tr valign="middle" bgcolor="#eeeeee">
	<td style="border-bottom:1px solid #cccccc;" colspan="2">
		<table border=0 cellpadding="3" cellspacing="0">
		<%if company="" then
			company="GY"%>
			<tr valign="middle" bgcolor="#eeeeee">
				<td>Company</td>
				<td>
				<select name="company" onchange="screencompany(this.value)"><%
			        rst.Open "select * from companycodes where active = 1 order by name", cnn
			        if not rst.eof then
			        do until rst.eof
				        %><option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><%=rst("name")%></option><%
				        rst.movenext
			        loop
			        end if
			        rst.close%>
				</select>
				</td>
			</tr>
		<%else%>
			<input type="hidden" name="company" value="<%=company%>">
		<%end if%>
		<tr>
			<td>Vendor Name</td>
			<td><input name="name" type="text" size="30" maxlength="255" value="<%=name%>"></td>
		</tr>
		<tr><td></td>
			<td><%if vendor<>"0" then ticket.Display 0,true,true,false %></td>
		</tr>
		</table>
	</td>
</tr>
<tr valign="middle" bgcolor="#eeeeee">
	<td style="border-top:1px solid #ffffff;" colspan="2">
	<table border=0 cellpadding="3" cellspacing="0">
	<tr>
		<td></td>
		<td><b>Main Address</b></td>
	</tr>
	<tr>
		<td align="right">Street</td>
		<td><input name="address_1" type="text" size="30" maxlength="30" value="<%=address_1%>"></td>
	</tr>
	<tr>
		<td></td>
		<td><input name="address_2"  type="text" size="30" maxlength="30" value="<%=address_2%>"></td>
	</tr>
	<tr>
		<td align="right">City</td>
		<td><input name="city" type="text" size="15" maxlength="29" value="<%=city%>"></td>
	</tr>
	<tr>
		<td align="right">State</td>
		<td><input name="state" type="text" size="4" maxlength="4" value="<%=state%>"> &nbsp;&nbsp;Zip: <input name="zip" type="text" size="10" maxlength="10" value="<%=zip%>"></td>
	</tr>
	<tr>
		<td align="right">Telephone</td>
		<td><input name="telephone" type="text" size="15" maxlength="15" value="<%=telephone%>"></td>
	</tr>
	<tr>
		<td align="right">FAX number</td>
		<td><input name="fax_number" type="text" size="15" maxlength="15" value="<%=fax_number%>"></td>
	</tr>
	<tr>
		<td align="right">Contact Name</td>
		<td><input name="contact_1_name" type="text" size="15" maxlength="30" value="<%=contact_1_name%>"></td>
	</tr>
	<tr>
		<td></td>
	</tr>
	<tr bgcolor="#eeeeee">
		<td></td>
		<td>
		<input type="hidden" name="vendor" value="<%=vendor%>">
		<input type="submit" name="action" value="<%if vendor="0" then response.write "Save" else response.write "Update"%>"> &nbsp;<input type="button" value="Cancel" onclick="closewin();">
		</td>
	</tr>
	</table>
	</td>
</tr>
</table>		
</form>
</body>
</html>
