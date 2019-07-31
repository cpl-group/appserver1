<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim tenantnum, flr, sqft, taxexempt, billingname, leaseexpired, oldleaseexpired, interm, intermcharges, startdate, bldg, action, tid, pid, oldtid
tenantnum = secureRequest("tenantnum")
oldtid = secureRequest("tid")
pid = secureRequest("pid")
flr = secureRequest("flr")
sqft = secureRequest("sqft")
taxexempt = secureRequest("taxexempt")
billingname = secureRequest("billingname")
leaseexpired = secureRequest("leaseexpired")
oldleaseexpired = secureRequest("oldleaseexpired")
interm = secureRequest("interm")
intermcharges = secureRequest("intermcharges")
startdate = secureRequest("startdate")
bldg = secureRequest("bldg")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

strsql = "INSERT INTO tblleases (tenantnum, flr, sqft, taxexempt, billingname, leaseexpired, interm, intermcharges, startdate, bldgnum) values ('"&tenantnum&"', '"&flr&"', '"&sqft&"', '"&taxexempt&"', '"&billingname&"', '"&leaseexpired&"', '"&interm&"', '"&intermcharges&"', '"&startdate&"', '"&bldg&"')"

'Logging Update
logger(strsql)
'end Log

cnn1.Execute strsql
strsql = "UPDATE tblleases set leaseexpired='"&oldleaseexpired&"' WHERE billingid="&oldtid

'Logging Update
logger(strsql)
'end Log

cnn1.Execute strsql

rst1.Open "SELECT max(billingid) as id FROM tblleases", cnn1
if not rst1.eof then tid = rst1("id")

dim returnpage
returnpage = "newleaseutilityedit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&oldtid="&oldtid

Response.Redirect returnpage

%>