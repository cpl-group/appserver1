 <%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim cnn, sqlstr, company, vendor, name, address_1, address_2, city, state, zip, telephone, fax_number, companyname, contact_1_name
company			= secureRequest("company")
vendor			= secureRequest("vendor")
name			= secureRequest("name")
address_1		= secureRequest("address_1")
address_2		= secureRequest("address_2")
city			= secureRequest("city")
state			= secureRequest("state")
zip				= secureRequest("zip")
telephone		= secureRequest("telephone")
fax_number		= secureRequest("fax_number")
companyname		= secureRequest("companyname")
contact_1_name	= secureRequest("contact_1_name")
set cnn = server.createobject("ADODB.connection")
cnn.open getConnect(0,0,"intranet")
if company<>"" then
	select case request("action")
	case "Save"
		sqlstr = "INSERT INTO "&company&"_MASTER_APM_VENDOR (vendor, name, address_1, address_2, city, state, zip, telephone, fax_number, contact_1_name) VALUES ('"&vendor&"', '"&name&"', '"&address_1&"', '"&address_2&"', '"&city&"', '"&state&"', '"&zip&"', '"&telephone&"', '"&fax_number&"', '"&contact_1_name&"')"
	case "Update"
		sqlstr = "UPDATE "&company&"_MASTER_APM_VENDOR SET name='"&name&"', address_1='"&address_1&"', address_2='"&address_2&"', city='"&city&"', state='"&state&"', zip='"&zip&"', telephone='"&telephone&"', fax_number='"&fax_number&"', contact_1_name='"&contact_1_name&"' WHERE vendor='"&vendor&"'"
	end select
end if
'response.write sqlstr
'response.end
if sqlstr<>"" then
	logger(sqlstr)
	cnn.execute sqlstr
end if
%>
<script>
opener.location.reload();
window.close();
</script>