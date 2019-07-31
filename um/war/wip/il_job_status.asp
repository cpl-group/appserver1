<%option explicit%>
<!--#include file="adovbs.inc"-->
<%
dim cnn, cmd, rs,jid,invoiceid,prm1,prm2,c
c=trim(request("c"))
jid=trim(request("jid"))
invoiceid=trim(request("invoiceid"))
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")

' open connection
cnn.Open getConnect(0,0,"Intranet")

' Set command properties
cmd.CommandType = adCmdStoredProc
select case c
case "IL"
cmd.CommandText = "il_job_status"
case "GY"
cmd.CommandText = "gy_job_status"
case "NY"
cmd.CommandText = "ny_job_status"
end select

' Specify connection
cmd.ActiveConnection = cnn

'return data
set rs=cmd.execute
dim i
%>
<title>Job Status Report</title>
<script language="JavaScript">
function toggleHelp(){
  if (document.all.quickhelp.style.display == "none") {
    document.all.quickhelp.style.display = "inline"
  } else {
    document.all.quickhelp.style.display = "none"
  }
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#ffffff">

<table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc" width="100%">
<tr bgcolor="#ffffff">
  <td colspan="10">
  <div align="right"><a href="javascript:toggleHelp();" style="text-decoration:none;"><b>Quick Help</b></a></div>
  <div id="quickhelp" style="display:none;">
  Shaded rows indicate job hours over estimate.<br>
  The <b>index</b> column shows actual hours as a percentage of estimated hours; shaded rows will always have an index greater than 100%.
  </div>
  </td>
</tr>
<tr bgcolor="#dddddd" style="font-weight:bold;"><td>Job #</td><td>Type</td><td>Description</td><td>Address</td><td>Floor</td><td>% Complete</td><td>Estimated Hrs</td><td>Hours</td><td>OverT</td><td>Index</td></tr>
<%
while not rs.eof
  if rs("diff")>1 then%>	
<tr bgcolor="#f0f6e3"><%
  else%>
<tr bgcolor="#ffffff"><%
end if%>
<td><%=rs("job")%></td><td><%=rs("type")%></td><td><%=rs("description")%></td><td><%=rs("address_1")%></td><td><%=rs("floor")%></td><td align="center"><%=cint(rs("percent_complete"))&"%"%></td><td align="right"><%=rs("total_labor_units_est")%></td><td align="right"><%=rs("hours")%></td><td align="right"><%=rs("overt")%></td><td align="right"><%=formatpercent(rs("diff"),0)%></td></tr>
<%	
  rs.movenext()
wend
%>
<tr bgcolor="#eeeeee">
  <td colspan="10">
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td><div style="position:inline;width:18px;height:12px;background:#f0f6e3;border:1px solid #999999;">&nbsp;</div></td>
    <td width="6">&nbsp;</td>
    <td>Shaded rows indicate job hours over estimate</td>
  </tr>
  </table>  
  </td>
</tr>
</table>
<br>
</body>
</html>