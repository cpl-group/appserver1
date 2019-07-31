
<%@Language="VBScript"%>

<%

acct=Request.Querystring("acctid")
v=Request.Querystring("vendor")
name1=Request.Querystring("name1")
addr=Request.Querystring("addr2")
bldg=Request.Querystring("bldg")


Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

strsql = "update tblacctsetup set vendor='" & Request.Querystring("vendor") & "',vendorname='" & name1 & "',serviceaddr='" &addr & "', EscoRef='"& request("esco") &"', Esco='"&Request.Querystring("accounttype")&"', locked='"&Request.Querystring("locked")&"' where acctid='" & acct & "'"
'response.write strsql
'response.end
cnn1.execute strsql
set cnn1=nothing%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" bgcolor="#3399CC">
  <tr>
    <td>
      <div align="center"><i><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Account Updated</font></b></i></div>
    </td>
  </tr>
</table>
<div align="center"><i><b></b></i></div>


