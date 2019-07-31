<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim rid, action, season, smonth, sday, emonth, eday, seasonid,effective_date
rid = request("rid")
seasonid = request("seasonid")
season = request("season")
smonth = request("smonth")
sday = request("sday")
emonth = request("emonth")
eday = request("eday")
action = request("action")
effective_date = request("effective_date")

if trim(action)="Save" then
	strsql = "INSERT INTO rateseasons (effective_date, season, smonth, sday, emonth, eday, regionid) values ('" & effective_date & "', '"&season&"', '"&smonth&"', '"&sday&"', '"&emonth&"', '"&eday&"', '"&rid&"')"
else
	strsql = "UPDATE rateseasons set season='"&season&"', smonth='"&smonth&"', effective_date='"&effective_date&"', sday='"&sday&"', emonth='"&emonth&"', eday='"&eday&"' WHERE id="&seasonid
end if
'response.Write strsql
'response.End
cnn1.Execute strsql
if trim(action)="Save" then 'need to find the bid for the building just added
	rst1.Open "SELECT max(id) as id FROM rateseasons", cnn1
	if not rst1.eof then seasonid = rst1("id")
end if

Response.redirect "seasonView.asp?rid="&rid&"&seasonid="&seasonid
%>