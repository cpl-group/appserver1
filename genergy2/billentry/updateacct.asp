<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim id, acctid, action, vendor, name1, addr2, utility, bldg, accounttype, escoRef, locked
id = Request("id")
acctid = Request("acctid")
action = Request("action")
vendor = Request("vendor")
name1 = Request("name1")
addr2 = Request("addr2")
utility = cint(Request("utility"))
bldg = Request("bldg")
accounttype = Request("accounttype")
escoRef = Request("escoRef")
locked = Request("locked")
if trim(locked)="" then locked = 0
dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getLocalConnect(bldg)
if lcase(trim(action))="save" then
  strsql = "insert into tblacctsetup (acctid,vendorname,serviceaddr,vendor,utility,bldgnum, Escoref, Esco, locked) values ('"&acctid&"','" & name1 & "','" &addr2 & "', '" &vendor& "','" &utility& "','" &bldg& "', '"&escoRef&"', "&accounttype&", "&locked&")"
elseif lcase(trim(action))="update" then
  strsql = "update tblacctsetup set vendor='"&vendor&"', vendorname='"&name1&"', serviceaddr='"&addr2&"', EscoRef='"&escoRef&"', Esco='"&accounttype&"', locked='"&locked&"' where acctid='"&acctid&"'"
end if
'response.write strsql
'response.end
cnn1.execute strsql
set cnn1=nothing%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" bgcolor="#3399CC">
  <tr>
    <td>
<%if lcase(trim(action))="save" then%>
      <div align="center"><i><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Account Created</font></b></i></div>
<%elseif lcase(trim(action))="update" then%>
      <div align="center"><i><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Account Updated</font></b></i></div>
<%else%>
      <div align="center"><i><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">????</font></b></i></div>
<%end if%>
    </td>
  </tr>
</table>
<div align="center"><i><b></b></i></div>


