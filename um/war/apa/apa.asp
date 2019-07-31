<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = '"&c&"_activity_ara_status'",cnn
crdate=rs(0)

rs.close

' specify stored procedure to run based on company
select case c
'case "IL"
'cmd.CommandText = "IL_APA"
case "GY"
cmd.CommandText = "GY_APA"
'case "NY"
'cmd.CommandText = "NY_APA"
case "GE"
cmd.CommandText = "GE_APA"
end select

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
<script language="JavaScript1.2">
function cash_detail(c,customer,d) {
	theURL="https://appserver1.genergy.com/um/war/apa/cash_detail.asp?c="+c+"&customer="+customer+"&d=" +d
	openwin(theURL,800,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scrollbars=yes, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>

<title>Genergy War Room - Cash Receipt</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
td { font-size:smaller; }
</style>
</head>
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">


<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
<tr bgcolor="#228866">
  <td width="8%"><span class="standardheader">Cust. #</span></td>
  <td width="35%"><span class="standardheader">Customer Name</span></td>
  <td width="33%"><span class="standardheader">Total Cash Collected</span></td>
</tr>

      <%
total = 0

while not rs.EOF 
total = total + rs("tot")

%>
      
<tr bgcolor="#ffffff">
  <td><%=rs("customer")%> </td>
  <td><%=rs("name")%> </td>
  <td align="right"><a href="javascript:cash_detail('<%=c%>','<%=rs("customer")%>','<%=d%>')"><%=formatnumber(rs("tot"),2)%></a> </td>
</tr>
  <%
 rs.movenext


wend
%>
    </table>
      <p>&nbsp;</p>
    
	<table width="99%" border="0" dwcopytype="CopyTableRow">
      <tr> 
        <td height="23" width="35%"> 
          <div align="left"><font size="2">update as of <%=formatdatetime(crdate,0)%></font></div>
        </td>
        <td height="23" width="35%">
          <div align="right"><b>Collected within the last <%=d%> days </b></div>
        </td>
        <td height="23" width="6%"> 
          <div align="right"><b><%=formatnumber(total,2)%></b></div>
        </td>
      </tr>
    </table>
	
	<br><br>
	  <%
	  set cnn = nothing %>
	  
</body>
</html>