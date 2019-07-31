<%option explicit%>
<!--#include file="adovbs.inc"-->
<%
'http params
dim crdate,d,c,o,total,currentjob,currentcustomer,gtotal

d = request("d")
c = request("c")
o = request ("o")

gtotal=0
'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")

' open connection
cnn.Open "driver={SQL Server};server=10.0.7.20;uid=sa;pwd=!general!;database=main;"
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_activity_ara_status'",cnn
crdate=rs(0)
rs.close


' specify stored procedure to run based on company
if c="IL" then
cmd.CommandText = "sp_outstanding"
else
cmd.CommandText = "sp_GY_outstanding"
end if
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("day", adinteger, adParamInput)
cmd.Parameters.Append prm

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

'return set to recordset rs
cnn.test  d, rs


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


<script type="text/javascript">

function openWindow(jobno,company)
{

// Append jobno to http link
var urlLink     = "https://appserver1.genergy.com/um/war/jc/jc.asp?c=" + company + "&j=" + jobno

// Open new window and customize window settings
window.open(urlLink,"window","scrollbars=no,width=900,height=600,resizeable")

}

</script>



<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">

<tr style="font-family:arial; font-size:12; color:black" bgcolor="#0099FF">
    
  <td style="font-size:10" rowspan="2" align="center">
  

    <table width="99%" border="1">
      <tr> 
        <td width="8%">job </td>
        <td width="22%">
          <div align="center">Customer</div>
        </td>
        <td width="20%">
          <div align="center">Job Name</div>
        </td>
        <td width="10%">Invoice #</td>
        <td width="15%">Invoice Date</td>
        <td width="10%">Days out</td>
        <td width="15%">
          <div align="right">Amount</div>
        </td>
      </tr>
     </table>
	 <table width="99%" border="0">
        <%
total = 0
currentjob = cstr(rs("job"))
currentcustomer = rs("customer")

while not rs.EOF 



if cstr(rs("job")) = cstr(currentjob) and o = "J" then 
total = total + rs("outstanding_amt")
gtotal=gtotal + rs("outstanding_amt")
else

if cstr(rs("customer")) = currentcustomer and o = "C" then 
total = total + rs("outstanding_amt")
gtotal=gtotal + rs("outstanding_amt")
else



%><tr><td colspan="6">&nbsp;</td><td align="right"><b><%=formatnumber(total,2)%></b></td></tr>
<%if total=0 then%></table><%end if%>
<hr><table width="99%" border="0">
<%
gtotal=gtotal + rs("outstanding_amt")
total =  rs("outstanding_amt")
currentjob = cstr(rs("job"))
currentcustomer = rs("customer")
end if
end if


%>
	<tr><%if rs("job") = "" then %>
        <td width="8%" height="20" bgcolor="#FF0000"><%=rs("job")%></td>
		<% else %>
		<td width="8%" height="20" > <a href="javascript:openWindow('<%=rs("job")%>','<%=c%>')"><%=rs("job")%></a></td>
		<% end if%>
        <td width="22%" height="20"> <%=rs("customer")%> </td>
        <td width="20%" height="20"> <%=rs("jobname")%> </td>
        <td width="10%" height="20"> <%=rs("invoice#")%> </td>
        <td width="15%" height="20"> <%=rs("invoice_date")%> </td>
        <td width="10%" height="20"> <%=rs("past_due_days")%> </td>
        <td width="15%" height="20" align="right"><%=formatnumber(rs("outstanding_amt"),2)%></td>
        
      </tr>
      <%
 rs.movenext


wend
%>
<hr> 
<tr><td colspan="6">&nbsp;</td><td align="right"><b><%=formatnumber(total,2)%></b></td></tr>
    </table><hr> 
    <p>&nbsp;</p>
    
	<table width="99%" border="0" dwcopytype="CopyTableRow">
      <tr> 
        <td height="21" width="46%"> 
          <div align="left"><font size="2">updated as of <%=crdate%></font></div>
        </td>
        <td height="21" width="46%">
          <div align="right"><b>Outstanding Grand Total over <%=d%> days </b></div>
        </td>
        <td height="21" width="8%"> 
          <div align="right"><b><%=formatnumber(gtotal,2)%></b></div>
        </td>
      </tr>
    </table>
	
	
      <tr> 
        <td>
          <div align="right"></div>
  </td>
      </tr> 
	  <%
	  set cnn = nothing %>
</body>
</html>
