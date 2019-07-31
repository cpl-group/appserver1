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

dim rid, action, sweekday, stime, eweekday, etime, seasonid, rPid, peakname, label
action = secureRequest("action")
rid = secureRequest("rid")
rPid = secureRequest("rPid")
sweekday = secureRequest("sweekday")
stime = secureRequest("stime")
eweekday = secureRequest("eweekday")
etime = secureRequest("etime")
seasonid = secureRequest("seasonid")
label = secureRequest("label")
peakname = secureRequest("peakname")

if trim(action)="Save" then
	strsql = "INSERT INTO ratePeak (sweekday, stime, eweekday, etime, seasonid, peakname, label) values ('"&sweekday&"', '"&stime&"', '"&eweekday&"', '"&etime&"', '"&seasonid&"', '"&peakname&"', '"&label&"')"
else
	strsql = "UPDATE ratePeak set sweekday='"&sweekday&"', stime='"&stime&"', eweekday='"&eweekday&"', etime='"&etime&"', seasonid='"&seasonid&"', peakname='"&peakname&"', label='"&label&"' WHERE id="&rPid
end if
cnn1.Execute strsql
if trim(action)="Save" then 'need to find the bid for the building just added
	rst1.Open "SELECT max(id) as id FROM ratepeak", cnn1
	if not rst1.eof then rPid = rst1("id")
end if

Response.redirect "seasonView.asp?rid="&rid
%>