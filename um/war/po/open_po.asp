<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,total,p,j,c,sql

if request("ji")="" then 
  j = request ("jg")
else
  j = request ("ji")
end if

c = request("c")


'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_master_po'",cnn
crdate=rs(0)
rs.close

' specify stored procedure to run based on company
  sql="select distinct c.name as vendor, a.commitment as po_num,a.description as description,a.amount as amount,a.amount_invoiced as amount_invoiced,a.amount_paid as amount_paid,(a.amount_invoiced-a.amount)as amount_open from "&c&"_master_po_item a,"&c&"_master_po b,"&c&"_master_apm_vendor c where b.closed=0 and a.job='"&j&"' and a.commitment=b.commitment and b.vendor=c.vendor"
'response.write sql
'response.end
rs.open sql,cnn

%>
<html>
<head>
<script language="JavaScript1.2">
function po_item(c,p,j) {
  theURL="/um/war/po/po_item.asp?c="+c+"&p="+p+"&j=" +j
  openwin(theURL,800,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scrollbars=yes, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<title>PO Details</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#ffffff">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Open Purchase Orders</span></td>
</tr>
</table>

  
<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr bgcolor="#dddddd" style=font-weight:bold;"> 
    <td width="20%">PO number</td>
    <td width="20%">Vendor</td>
    <td width="20%">Description</td>
    <td width="20%">Amount</td>
    <td width="20%">Amount Invoiced</td>
    <td width="20%">Amount Balance</td>
  </tr>
  <%
total = 0

while not rs.EOF 
total = total + (rs("amount")-rs("amount_invoiced"))


%>
  <tr bgcolor="#ffffff"> 
    <td><a href="/um/opslog/poview.asp?po=<%=rs("po_num")%>"><%=rs("po_num")%></a></td>
    <td><a href="/um/opslog/poview.asp?po=<%=rs("po_num")%>"><%=rs("vendor")%></a></td>
    <td><%=rs("description")%></td>
    <td align="right"><%=formatcurrency(rs("amount"),2)%></td>
    <td align="right"><%=formatcurrency(rs("amount_invoiced"),2)%></td>
   <td align="right"><%=formatcurrency(rs("amount_open"),2)%></td>
    
    <%
 rs.movenext


wend
%>
  <tr bgcolor="#ffffff"> 
    <td colspan="5" align="right"><b>Total Balance:</b></td>
    <td align="right"><b><%=formatcurrency(total,2)%></b></td>
  </tr>
</table>

<table border=0 cellpadding="3" cellspacing="1">
<tr>
  <td>Updated as of <%=formatdatetime(crdate,0)%></td>
</tr>
<tr><td><input type="button" value="Close Window" onclick="window.close();"></td></tr>
</table>
<% set cnn = nothing %>
    
</body>
</html>