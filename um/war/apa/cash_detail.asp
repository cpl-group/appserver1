<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,d,c,o,total,currentjob,gtotal,customer

d = request("d")
c = request("c")
customer = request ("customer")

gtotal=0
'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_activity_ara_status'",cnn
crdate=rs(0)
rs.close

' specify stored procedure to run based on company

select case c 
case "IL" 
rs.open "select * from il_activity_ara_activity where customer='" & customer & "' and activity_date >= getdate()- " & cstr(d) & " and amount < 0",cnn
case "GY"
rs.open "select * from gy_activity_ara_activity where customer='" & customer & "' and activity_date >= getdate()- " & cstr(d) & " and amount < 0",cnn
case "NY"
rs.open "select * from ny_activity_ara_activity where customer='" & customer & "' and activity_date >= getdate()- " & cstr(d) & " and amount < 0",cnn
case "GE"
rs.open "select * from ge_activity_ara_activity where customer='" & customer & "' and activity_date >= getdate()- " & cstr(d) & " and amount < 0",cnn
end select
%>
<html>
<head>
<script language="JavaScript1.2">
function cash_detail(c,customer,d) {
	theURL="https://appserver1.genergy.com/um/war/apa/cash_detail.asp?c="+c+"&customer="+customer+"&d=" +d
	openwin(theURL,800,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scroll=on, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>

<title>Genergy War Room - Cash Receipt Detail</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
td { font-size:smaller; }
</style>
</head>
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">


<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
<tr bgcolor="#339966">
  <td width="8%"><span class="standardheader">Cust. #</span></td>
  <td width="35%"><span class="standardheader">Invoice</span></td>
  <td width="33%"><span class="standardheader">Activity Date</span></td>
  <td width="24%"><span class="standardheader">Amount</span></td>
</tr>

<%
total = 0

while not rs.EOF 
total = total + rs("amount")

%>
      
<tr bgcolor="#ffffff">
  <td><%=rs("customer")%> </td>
  <td><%=rs("invoice")%></td>
  <td><%=rs("activity_date")%></td>
  <td align="right"><b><%=formatcurrency(rs("amount"),2)%></b></td>
</tr>
      <%
 rs.movenext


wend
%>
</table>
    
<table width="99%" border="0" dwcopytype="CopyTableRow">
<tr> 
  <td width="35%">Update as of   <%=formatdatetime(crdate,0)%></td>
  <td width="35%"><b>Collected within the last <%=d%> days </b></td>
  <td align="right"><b><%=formatnumber(total,2)%></b></td>
</tr>
</table>
<p>&nbsp;</p>
	
	
	  <%
	  set cnn = nothing %>
	  
</body>
</html>