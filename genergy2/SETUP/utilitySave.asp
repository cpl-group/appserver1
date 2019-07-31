<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim uid, action, utility, utilitydisplay, utilitySuffix
uid = secureRequest("uid")
utility = secureRequest("utility")
utilitydisplay = secureRequest("utilitydisplay")
utilitySuffix = secureRequest("utilitySuffix")
action = secureRequest("action")


if trim(action)="Save" then
	strsql = "INSERT INTO tblutility (utility, utilitydisplay, utilitySuffix) values ('"&utility&"', '"&utilitydisplay&"', '"&utilitySuffix&"')"
else
	strsql = "UPDATE tblutility set utility='"&utility&"', utilitydisplay='"&utilitydisplay&"', utilitySuffix='"&utilitySuffix&"' WHERE utilityid="&uid
end if
'response.Write strsql
'response.End
cnn1.Execute strsql
if trim(action)="Save" then 'need to find the bid for the building just added
	rst1.Open "SELECT max(id) as id FROM regions", cnn1
	if not rst1.eof then uid = rst1("id")
end if

Response.redirect "utilityView.asp?uid="&uid
%>