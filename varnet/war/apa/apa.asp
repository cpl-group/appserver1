<%option explicit%>
<!--#include file="adovbs.inc"-->
<%
'http params
dim crdate,d,c,o,total,currentjob,currentcustomer,gtotal

d = request("d")
c = request("c")
'o = request ("o")

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
cmd.CommandText = "IL_APA"
else
cmd.CommandText = "GY_APA"
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

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">

<tr style="font-family:arial; font-size:12; color:black" bgcolor="#0099FF">
    
  <td style="font-size:10" rowspan="2" align="center">
  

    <table width="99%" border="1">
      <tr> 
        <td width="8%">Cust. #</td>
        <td width="35%"> 
          <div align="center">Customer Name</div>
        </td>
        <td width="33%"> 
          <div align="center">Total Cash Collected</div>
        </td>
        <td width="24%"> 
          <div align="right">Average of payments (days)</div>
        </td>
      </tr>
    </table>
	 
    <table width="99%" border="0">
      <%
total = 0

while not rs.EOF 
total = total + rs("tot")


%>
      
    </table>

    <table width="99%" border="0">
        <td width="8%" height="20"> <%=rs("customer")%> </td>
        <td width="35%" height="20"> <%=rs("name")%> </td>
        <td align="right" width="33%" height="20"> <%=formatnumber(rs("tot"),2)%> </td>  
		<td align="right" width="24%" height="20"> <%=rs("avg_diff")%> </td>
	
      <%
 rs.movenext


wend
%>
      
    </table>
      <p>&nbsp;</p>
    
	<table width="99%" border="0" dwcopytype="CopyTableRow">
      <tr> 
        <td height="23" width="35%"> 
          <div align="left"><font size="2">update as of   <%=formatdatetime(crdate,0)%></font></div>
        </td>
        <td height="23" width="35%"><b>Collected within the last <%=d%> days </b></td>
        <td height="23" width="6%"> 
          <div align="right"><b><%=formatnumber(total,2)%></b></div>
        </td>
        <td height="23" width="24%">&nbsp;</td>
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