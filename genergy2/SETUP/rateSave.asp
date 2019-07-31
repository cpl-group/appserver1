<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, action, rtid, rateid, rate, peak, utility, ratefrom, rateto, Itemtype, linecharge, monthstart, monthend, startdate, enddate
action = secureRequest("action")
rid = secureRequest("rid")
rtid = secureRequest("rtid")
rateid = secureRequest("rateid")
rate = secureRequest("rate")
peak = secureRequest("peak")
utility = secureRequest("utility")
ratefrom = secureRequest("ratefrom")
rateto = secureRequest("rateto")
Itemtype = secureRequest("Itemtype")
linecharge = secureRequest("linecharge")
monthstart = secureRequest("monthstart")
monthend = secureRequest("monthend")
startdate = secureRequest("startdate")
enddate = secureRequest("enddate")

if trim(action)="Save" then
	strsql = "INSERT INTO rate (rate, type, peak, utility, ratefrom, rateto, Itemtype, linecharge, monthstart, monthend, startdate, enddate) values ('"&rate&"', '"&rtid&"', '"&peak&"', '"&utility&"', '"&ratefrom&"', '"&rateto&"', '"&Itemtype&"', '"&linecharge&"', '"&monthstart&"', '"&monthend&"', '"&startdate&"', '"&enddate&"')"
elseif trim(action)="Delete" then
	strsql = "DELETE FROM rate WHERE id="&rateid
else
	strsql = "UPDATE rate set rate='"&rate&"', type='"&rtid&"', peak='"&peak&"', utility='"&utility&"', ratefrom='"&ratefrom&"', rateto='"&rateto&"', Itemtype='"&Itemtype&"', linecharge='"&linecharge&"', monthstart='"&monthstart&"', monthend='"&monthend&"', startdate='"&startdate&"', enddate='"&enddate&"' WHERE id="&rateid
end if
'response.write strsql
'response.end
on error resume next
cnn1.Execute strsql
if err.number<>0 then
	if err.number=-2147217900 then response.redirect "rateEdit.asp?ratedup=yes&rid="&rid&"&rtid="&rtid&"&rateid="&rateid&"&rate="&rate&"&peak="&peak&"&utility="&utility&"&ratefrom="&ratefrom&"&rateto="&rateto&"&Itemtype="&Itemtype&"&linecharge="&linecharge&"&monthstart="&monthstart&"&monthend="&monthend&"&startdate="&startdate&"&enddate="&enddate
end if

on error goto 0


if trim(action)="Save" then 'need to find the bid for the building just added
	rst1.Open "SELECT max(id) as id FROM rate", cnn1
	if not rst1.eof then rateid = rst1("id")
end if

'if trim(action)="Delete" then Response.redirect "rateTypeEdit.asp?rid="&rid&"&rtid="&rtid
'Response.redirect "rateTypeView.asp?rid="&rid&"&rtid="&rtid&"&rateid="&rateid
%>
<script>
opener.document.location="/genergy2/setup/rateTypeView.asp?rid=<%=rid%>&rtid=<%=rtid%>"
window.close()
</script>