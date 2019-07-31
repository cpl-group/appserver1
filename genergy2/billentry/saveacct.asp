<%option explicit%>

<%
dim id1, v, name, addr, u, b, accounttype, escoRef, locked
id1=Request.querystring("acctnum")
v=Request.querystring("vendor")
name=Request.querystring("vname")
addr=Request.querystring("addr")
u=Request.querystring("utility")
b=Request.querystring("bldg")
accounttype=Request("accounttype")
escoRef=Request("escoRef")
locked=Request("locked")


dim cnn1, strsql, rst3, str3
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy2")

strsql = "insert into tblacctsetup (acctid,vendorname,serviceaddr,vendor,utility,bldgnum, Escoref, Esco, locked) values('" &id1& "','" & name & "','" &addr & "', '" &v& "','" &u& "','" &b& "', '"&escoRef&"', "&accounttype&", "&locked&")"

cnn1.execute strsql

Set rst3 = Server.CreateObject("ADODB.recordset")
	 str3="select ypid from billyrperiod where bldgnum='" &b&"' and utility='"&u&"'"
	rst3.Open str3, cnn1, 0, 1, 1
	if rst3.eof then%>
<html>
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Account 
        has been saved</i></font></font></div>
    </td>
  </tr>
<tr> 
    <td width="1106"><b><font face="Arial, Helvetica, sans-serif" size="2"><i>No 
      bill periods have been defined for this building. Please contact <a href ="mailto:george_nemeth@genergy.com">George 
      Nemeth</a> to add billperiods.</i></font></b></td>
  </tr>
</table>
  <%else%>
  <table width="100%"  >
   <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><font size="4" color="#FFFFFF"><i>Account 
        has been saved</i></font></font></div>
    </td>
  </tr>
 <tr> 
    <td width="1106"><b><i><font face="Arial, Helvetica, sans-serif" size="2">Bill 
      periods for this building have already been defined. To update or add a 
      billperiod, please contact <a href="mailto:george_nemeth@genergy.com">George 
      Nemeth</a>.</font></i></b></td>
  </tr>
  
</table>
</html>
<%end if
set cnn1=nothing%>
