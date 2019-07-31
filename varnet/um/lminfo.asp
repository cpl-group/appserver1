
<html>
<head>
<%@Language="VBScript"%>
<script>
function modify(meterid, meternum, srvname, dbname){
    parent.frames.detail.location="lmdetail.asp?meterid="+meterid+"&meternum="+meternum+"&srvname="+srvname+"&dbname="+dbname
}
</script>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<%
meternum=request("meternum")
bldgnum=request("bldgnum")
srvname=request("srvname")
dbname=request("dbname")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sql = "select meterid,meternum,bldgnum,lmnum,lmchannel,g1onlinedate from ["&srvname&"]."&dbname&".dbo.meters where meternum like '%" & meternum & "%'"
	'response.write sql
	'response.end
rst1.Open sql, cnn1, 0, 1, 1

if rst1.eof then
%>
<table width="100%" border="0">
  <tr>
    <td>
      <div align="center"> 
        <p><font face="Arial, Helvetica, sans-serif"><i>
		No Meter Available</i></font></p>
      </div>
    </td>
  </tr>
</table>
<%
else
count=0
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#3399CC">
      <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>METER 
        LIST</b></font></div>
    </td>
  </tr>
</table>
<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
  	<td width="11%" height="2"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp</font><font size="2"></font></td>
    <td width="12%" height="2"><font face="Arial, Helvetica, sans-serif" size="2">Meterid</font><font size="2"></font></td>
    <td width="12%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Meternum</font></td>
    <td width="14%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Bldgnum</font></td>
	<td width="15%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Lmnum</font></td>
	<td width="16%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Lmchannel</font></td>
    <td width="20%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">G1onlinedate</font></td>
  </tr>
  <% 
	do until rst1.EOF 
    count=count+1
  %>
  <form name="form1" method="post" action="">
    <tr> 
	  <td width=11%> <font size="2"> 
	  
        <input type="hidden" name="meterid" value="<%=trim(rst1("meterid"))%>"> 
        <input type="button" name="submit" value="edit" onclick='modify(meterid.value, "<%=meternum%>", "<%=srvname%>", "<%=dbname%>")'> 
        </font></td>	
      <td width=12%> <font size="2"> <%=rst1("meterid")%> </font></td>
      <td width="12%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("meternum")%> </font></td>
	  <td width="14%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("bldgnum")%></font></td>
	  <td width="15%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("lmnum")%></font></td>	
	  <td width="16%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("lmchannel")%></font></td>
	  <td width="20%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("g1onlinedate")%></font></td>	
    </tr>
  </form>
  <%
	rst1.movenext
	loop
%>
</table>
<table width="100%" border="0" bgcolor="#3399CC">
  <tr> 
    <td> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=count%> Meter(s) found</font></b></div>
    </td>
  </tr>
</table>

<%
end if
rst1.close
set cnn1=nothing
%>
</body>
</html>
