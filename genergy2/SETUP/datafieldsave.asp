<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim bldg, action, tid, pid, lid, meterid
meterid = secureRequest("meterid")
lid = secureRequest("lid")
tid = secureRequest("tid")
pid = secureRequest("pid")
bldg = secureRequest("bldg")
action = secureRequest("action")

dim cnn1, rst1, strsql, strsqlDS
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim field(7)
field(0) = secureRequest("fieldname1")
field(1) = secureRequest("fieldname2")
field(2) = secureRequest("fieldname3")
field(3) = secureRequest("fieldname4")
field(4) = secureRequest("fieldname5")
field(5) = secureRequest("fieldname6")
field(6) = secureRequest("fieldname7")


	strsql = "UPDATE datasource set fieldname1='"&field(0)&"', fieldname2='"&field(1)&"', fieldname3='"&field(2)&"', fieldname4='"&field(3)&"', fieldname5='"&field(4)&"', fieldname6='"&field(5)&"', fieldname7='"&field(6)&"' WHERE meterid="&meterid
response.Write strsql

'Logging Update
logger(strsql)
'end Log

cnn1.Execute strsql

Response.Redirect "meteredit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&meterid="&meterid
%>