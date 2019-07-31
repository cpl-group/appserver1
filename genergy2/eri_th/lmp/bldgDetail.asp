<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim ypid, bldg, meterid, luid, sd, utility, billingid
ypid = request("Y")
luid = request("luid")
bldg = request("bldg")
meterid = request("meterid")
billingid = request("billingid")
utility = request("utility")
dim billingtype
if trim(luid)<>"" then
    billingtype = "Tenant "
elseif trim(bldg)<>"" then
    billingtype = "Building "
elseif trim(meterid)<>"" then
    billingtype = "Meter "
else
    billingtype = " "
end if


'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")

' open connection
'cnn.open getConnect(0,bldg,"intervaldata")'getLocalConnect(bldg)'"Provider=SQLOLEDB;Data Source=10.0.7.16;User Id=genergy1;Password=g1appg1;Initial Catalog=dbIntervaldata;"
 cnn.open getConnect(0,bldg,"billing")
'response.write getLocalConnect(bldg)
'response.write getConnect(0,bldg,"billing")
'dim PIDVAL
'PIDVAL=getPID(bldg)	
'response.write PIDVAL
'response.write getLocalConnect(bldg)

'response.write  Application("IntranetIP")
'response.write Application("dbDefault")
'response.write getConnect(0,bldg,"Intervaldata")
'Response.Write "ypid :" & ypid
'Response.Write "bldg :" & bldg
'response.end
cnn.CursorLocation = adUseClient

' specify stored procedure to run
cmd.CommandText = "sp_abc"
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("ypid", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bldgnum", adChar, adParamInput, 20)
cmd.Parameters.Append prm

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

'return set to recordset rs
'response.write "exec sp_abc "&ypid&", '"&bldg&"'"
'response.end
cnn.test ypid, bldg, rs


%>
<html>
<head>
<title></title>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">
<font face="arial" size="2">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#000000" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Billing History</b></font></td>
    <td bgcolor="#000000" width="46%" height="2"><div align="right"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><a href="javascript:document.location='opt_accthist.asp?bldg=<%=bldg%>&meterid=<%=meterid%>&billingid=<%=billingid%>&utility=<%=utility%>'" style="text-decoration:none;color:white" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Billing History</a></b></font></div></td>
  </tr>
  <tr>
    <td width="46%">&nbsp;</td>
    <td width="46%">&nbsp;</td>
  </tr>
</table>

<span style="font-size:10" rowspan="2" align="center"><%if not rs.eof then%>From <%=rs("datefrom")%> To <%=rs("dateto")%><%end if%></span>
<%
dim Bsqft, Osqft, wsqft
if not(rs.EOF) then
	Bsqft = rs("Bsqft")
	Osqft = rs("Osqft")
	wsqft = rs("wsqft")
else
  response.write "No available data."
end if
%>
<table border="1" width="100%">
<tr style="font-family:arial; font-size:12; color:black" bgcolor="#0099FF">
    <td rowspan="2" align="center"><b>Account ID</b></td>
    <td colspan="<%if trim(utility)="2" then%>5<%else%>2<%end if%>" align="center"><b>Billed</b></td>
<%if trim(utility)="2" then%>
    <td colspan="2" align="center"><b>System</b></td>
<%end if%>
	<td style="font-family:arial; font-size:10;" rowspan="2" valign="bottom" align="center">Load Factor</td>
</tr>
<tr style="font-family:arial; font-size:10; color:black" bgcolor="#0099FF">
<%if trim(utility)="2" then%>
	<td align="center">On Peak KWH</td>
	<td align="center">Off Peak KWH</td>
<%end if%>
	<td align="center">Total Consumption</td>
	<td align="center">Total Bill</td>
<%if trim(utility)="2" then%>
	<td align="center">Peak Demand</td>
	<td align="center">Peak Demand</td>
	<td align="center">Bill Variance</td>
<%end if%>
</tr>
<%do until rs.EOF%>
  <tr style="font-size:10">
  <td align="right"><%=rs("acctid")%></td>
<%if trim(utility)="2" then%>
  <td align="right"><%=formatnumber(rs("onpeak"))%></td>
  <td align="right"><%=formatnumber(rs("offpeak"))%></td>
<%end if%>
  <td align="right"><%=formatnumber(rs("totalUsage"))%></td>
  <td align="right"><%=formatcurrency(rs("totbillamt"))%></td>
  <td align="right"><%=formatnumber(rs("totalDemand"))%></td>
<%if trim(utility)="2" then%>
  <%if cdbl(rs("sdemand")) <= 1 then%>
  	<td align="right">NA</td>
	  <td align="right">NA</td>
  <%else%>
	  <td align="right"><%=formatnumber(rs("Sdemand"))%></td>
	  <td align="right"><%=abs(formatnumber(rs("variancekw")))%>%</td>
  <%end if%>
    <td align="right"><%=formatpercent(rs("loadfactor"))%></td>
<%end if%>
</tr>
<%rs.movenext
loop%>
</table><p>
<blockquote>Building Square Ft.: <%=formatnumber(Bsqft,0)%><br>
Occupied Square Ft.: <%=formatnumber(Osqft,0)%>*<br>
Watts per Square Ft.: <%=formatnumber(wsqft,2)%></blockquote>
<table cellspacing="0" cellpadding="0" width="100%">
  <tr>
    <td align="right"><font size="2" face="arial"><b><a href="javascript:history.back()">Back</a></b></font></td>
  </tr>
</table>
<font size="2"><i><font face="Arial, Helvetica, sans-serif" size="1">*Note: Occupied 
Square Footage is based on available data for ERI &amp; Submetered tenants. 
</font> </i> </font> 
<tr style="font-size:10"> 
</body>
</html>