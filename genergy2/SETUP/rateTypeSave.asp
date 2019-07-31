<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, cmd, strsql, prm
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, action, rtid, rtype, rtypecopyto, rtypecopyfrom
action = secureRequest("action")
rid = secureRequest("rid")
rtid = secureRequest("rtid")
rtype = secureRequest("rtype")
rtypecopyfrom = secureRequest("rtypecopyfrom")
rtypecopyto = secureRequest("rtypecopyto")

if trim(action)="Save" then
	strsql = "INSERT INTO ratetypes (type, regionid) values ('"&rtype&"', '"&rid&"')"
elseif trim(action)="Copy" then
	strsql = "INSERT INTO ratetypes (type,regionid) values ('"&rtypecopyto&"','"&rid&"');insert into rate (rate, type, peak, utility, rateFrom, RateTo, ItemType, linecharge, monthstart, monthend, startdate, enddate) select rate, a.type, peak, utility, rateFrom, RateTo, ItemType, linecharge, monthstart, monthend, startdate, enddate from (select [id] as type from ratetypes where type = '"&rtypecopyto&"') a, rate r where r.type = "&rtypecopyfrom 
else
	strsql = "UPDATE ratetypes set type='"&rtype&"' WHERE id="&rtid
end if
'response.write strsql
'response.end
cnn1.Execute strsql
if trim(action)="Save" then 'need to find the bid for the building just added
	rst1.Open "SELECT max(id) as id FROM ratetypes", cnn1
	if not rst1.eof then rtid = rst1("id")
end if

Response.redirect "rateTypeView.asp?rid="&rid&"&rtid="&rtid
%>