
<html>
<head>
<%@Language="VBScript"%>
<script>
function modify(meterid){
    parent.frames.meter.location="lminfo.asp"
}
</script>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<%

meterid=request("meterid")
meternum=request("meternum")
srvname=request("srvname")
dbname=request("dbname")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

    sql = "select meternum,bldgnum,lmnum,lmchannel,g1onlinedate from "&srvname&"."&dbname&".dbo.meters where meterid ='" & meterid & "' "
	
rst1.Open sql, cnn1, 0, 1, 1
if not rst1.eof then
%>

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
  <form name="form1" method="post" action="lmupdate.asp">
  <%Set rst2 = Server.CreateObject("ADODB.recordset")
  sql = "select ip from master.dbo.rm where srvname='"&srvname&"'"
  rst2.open sql,cnn1,0,1,1 %>
  
  <input type="hidden" name="srvname" value="<%=rst2("ip")%>">
  <input type="hidden" name="srvname1" value="<%=srvname%>">
  <input type="hidden" name="dbname" value="<%=dbname%>">  
    <tr> 
	  <td width=11%> <font size="2"> 
        <input type="hidden" name="meterid" value="<%=meterid%>"> 
        <input type="submit" name="submit" value="Save"> 
        </font></td>	
      <td width=12%> <font size="2"> <%=meterid%> </font></td>
      <td width="12%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2">
	    <input type="hidden" name="meternum" value="<%=meternum%>">  
        <%=rst1("meternum")%> </font></td>
	  <td width="14%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("bldgnum")%></font></td>
	  <td width="15%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="lmnum" value="<%=rst1("lmnum")%>"></font></td>	
	  <td width="16%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="lmchannel" value="<%=rst1("lmchannel")%>"></font></td>
	  <td width="20%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("g1onlinedate")%></font></td>	
    </tr>
  </form>
</table>
<%
end if
rst1.close
set cnn1=nothing
%>
</body>
</html>
