<%option explicit%>
<!--#include file="adovbs.inc"-->
<%
'http params
dim ypid, b, m, luid
ypid = request("Y")
luid = request("luid")
b = request("b")
m = request("m")
dim billingtype
if trim(luid)<>"" then
    billingtype = "Tenant "
elseif trim(b)<>"" then
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
cnn.Open application("cnnstr_genergy1")
cnn.CursorLocation = adUseClient

' specify stored procedure to run
cmd.CommandText = "sp_abc"
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("ypid", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bldgnum", adChar, adParamInput, 5)
cmd.Parameters.Append prm

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

'return set to recordset rs
cnn.test ypid,b, rs


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
    <td bgcolor="#000000" width="46%" height="2">
      <div align="right"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><a href="javascript:document.location='opt_accthist.asp?b=<%=b%>&m=<%=m%>&luid=<%=luid%>'" style="text-decoration:none;color:white" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Billing History</a></b></font></div>
    </td>
  </tr>
  <tr>
    <td width="46%">&nbsp;</td>
    <td width="46%">&nbsp;</td>
  </tr>
</table>

<%
if not(rs.EOF) then%>
<table border="1" width="100%">
<tr style="font-family:arial; font-size:12; color:black" bgcolor="#0099FF">
    <td style="font-size:10" rowspan="2" align="center">From</td>
    <td style="font-size:10" rowspan="2" align="center">To</td>
    <td colspan="5" align="center"><b>Billed</b></td>
    <td colspan="2" align="center"><b>System</b></td>
</tr>
<tr style="font-family:arial; font-size:10; color:black" bgcolor="#0099FF">
<td align="center">On Peak KWH</td>
<td align="center">Off Peak KWH</td>
<td align="center">Total KWH</td>
<td align="center">Total Bill</td>
<td align="center">Peak Demand</td>
<td align="center">Peak Demand</td>
<td align="center">Bill Variance</td>
</tr>
<tr style="font-size:10"> 
  <%
response.write "<td>"& rs("datefrom") &"</td>"
response.write "<td>"& rs("dateto") &"</td>"
response.write "<td align=""right"">"& formatnumber(rs("onpeak")) &"</td>"
response.write "<td align=""right"">"& formatnumber(rs("offpeak")) &"</td>"
response.write "<td align=""right"">"& formatnumber(rs("totkwh")) &"</td>"
response.write "<td align=""right"">"& formatcurrency(rs("totbillamt")) &"</td>"
response.write "<td align=""right"">"& formatnumber(rs("totkw")) &"</td>"
'response.write "<td align=""right"">"& formatnumber(rs("Sonpeak")) &"</td>"
'response.write "<td align=""right"">"& formatnumber(rs("Soffpeak")) &"</td>"
'response.write "<td align=""right"">"& formatnumber(rs("Stotkwh")) &"</td>"
response.write "<td align=""right"">"& formatnumber(rs("Sdemand")) &"</td>"
response.write "<td align=""right"">"& abs(formatnumber(rs("variancekw"))) &"%</td>"
response.write "</tr>"
response.write "</table><p>"
response.write "<blockquote>Load Factor: "& formatpercent(rs("loadfactor")) &"<BR>"
response.write "Building Square Ft.: "& formatnumber(rs("Bsqft"),0) &"<br>"
response.write "*Occupied Square Ft.: "& formatnumber(rs("Osqft"),0) &"<br>"
response.write "Watts per Square Ft.: "& formatnumber(rs("wsqft"),2) &"</blockquote>"

else
    response.write "No available data."
end if
%>
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