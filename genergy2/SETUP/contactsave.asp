<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, ctid, action
pid = secureRequest("pid")
bldg = secureRequest("bldg")
ctid = secureRequest("ctid")
action = secureRequest("action")
dim name, address, city, state, zip, phone, fax, email, administrative, m_report, submeter_bills
name = secureRequest("name")
address = Left(secureRequest("address"),50)
city = secureRequest("city")
state = Left(secureRequest("state"),2)
zip = secureRequest("zip")
phone = secureRequest("phone")
fax = secureRequest("fax")
email = secureRequest("email")
administrative = trim(secureRequest("administrative"))
m_report = trim(secureRequest("m_report"))
submeter_bills = trim(secureRequest("submeter_bills"))
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

if trim(action)="Save" then
	strsql = "INSERT INTO contacts (name, address, city, state, zip, phone, fax, email, administrative, cid, bldgnum, m_report, submeter_bills) values ('"&name&"', '"&address&"', '"&city&"', '"&state&"', '"&zip&"', '"&phone&"', '"&fax&"', '"&email&"', '"&administrative&"', '"&pid&"', '"&bldg&"', '"&m_report&"', '"&submeter_bills&"')"
elseif trim(action)="Update" then
	strsql = "UPDATE contacts set name='"&name&"', address='"&address&"', city='"&city&"', state='"&state&"', zip='"&zip&"', phone='"&phone&"', fax='"&fax&"', email='"&email&"', administrative='"&administrative&"', m_report='"&m_report&"', submeter_bills='"&submeter_bills&"' WHERE id='"&ctid&"'"
elseif trim(action)="Delete" then
	strsql = "DELETE contacts WHERE id='"&ctid&"'"
end if
'response.Write strsql
'response.End

'Logging Update
logger(strsql)
'end Log

cnn1.Execute strsql

Response.redirect "contactview.asp?pid="&pid&"&bldg="&bldg
%>