<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,c,t_outstanding,j,t_amount,t_adj,t_net

'd = request("d")
c = request("c")
if request("ji")="" then 
  j = request ("jg")
else
  j = request ("ji")
end if
'o = request ("o")
if trim(c)="" then 
	response.write "<font face=""arial"" size=""2"" color=""red"">This section can not be displayed because <b>no billing contact</b> is specified for this job. </font>"
	response.end
end if

'adodb vars
dim cnn, cmd, rs, prm,tally,rst
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

rs.open "SELECT distinct isnull(tax,0) as tax,s.customer,s.status,s.invoice,a.amount, adjustment,outstanding_amount, cash_receipt, a.Activity_Date as generation_date,s.generation_date FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_ACTIVITY_ARA_ACTIVITY a on s.invoice=a.invoice  left join "&c&"_BILLED_BLI_INVOICE i on i.invoice=s.invoice WHERE a.amount > 0 and a.job='"&j&"' order by s.generation_date desc",cnn 'took out order by s.invoice in order to sort by date
'rs.open "SELECT distinct isnull(tax,0) as tax,s.customer,s.status,s.invoice,a.amount, adjustment,outstanding_amount, cash_receipt, a.Activity_Date as generation_date, isNull(Billed, 0) as Billed FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_ACTIVITY_ARA_ACTIVITY a on s.invoice=a.invoice  left join "&c&"_BILLED_BLI_INVOICE i on i.invoice=s.invoice WHERE a.amount > 0 and a.job='"&j&"' order by s.generation_date desc",cnn 'took out order by s.invoice in order to sort by date
'response.write "SELECT s.customer,s.status,s.invoice,a.amount, adjustment,outstanding_amount, cash_receipt, a.Activity_Date as generation_date, isNull(Billed, 0) as Billed FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_ACTIVITY_ARA_ACTIVITY a on s.invoice=a.invoice  left join "&c&"_BILLED_BLI_INVOICE i on i.invoice=s.invoice WHERE a.amount > 0 and a.job='"&j&"' order by s.invoice, s.generation_date"
'response.write "SELECT distinct s.customer,s.status,s.invoice,a.amount, adjustment,outstanding_amount, cash_receipt, a.Activity_Date as generation_date, isNull(Billed, 0) as Billed FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_ACTIVITY_ARA_ACTIVITY a on s.invoice=a.invoice  left join "&c&"_BILLED_BLI_INVOICE i on i.invoice=s.invoice WHERE a.amount > 0 and a.job='"&j&"' order by s.invoice, s.generation_date"
'response.write "SELECT distinct tax,s.customer,s.status,s.invoice,a.amount, adjustment,outstanding_amount, cash_receipt, a.Activity_Date as generation_date, isNull(Billed, 0) as Billed FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_ACTIVITY_ARA_ACTIVITY a on s.invoice=a.invoice  left join "&c&"_BILLED_BLI_INVOICE i on i.invoice=s.invoice WHERE a.amount > 0 and a.job='"&j&"' order by s.invoice, s.generation_date"
'response.end
%>
<html>
<head>
<script language="JavaScript1.2">
function po_item(c,p,j) {
  theURL="/um/war/po/po_item.asp?c="+c+"&p="+p+"&j=" +j
  openwin(theURL,800,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scroll=on, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<title>Invoice detail</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">

  <td><span class="standardheader">Amount Billed</span></td>
</tr>
</table>

<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr bgcolor="#dddddd" style="font-weight:bold;"> 
    <td>Customer</td>
    <td>Invoice</td>
    <td>Generation Date</td>
    <td>Status</td>
    <td>Amount w/Tax</td>
	<td>Amount</td>
    <td>Adjustment</td>
    <td nowrap>Cash Receipt</td>
    <td>Outstanding</td>
  </tr>
  <%
t_outstanding = 0
t_amount = 0
t_adj = 0
voidamt=0
voidamttax=0

while not rs.EOF 
'on error resume next
Dim t_cr,tax,voidamt,voidamttax
t_outstanding = t_outstanding + rs("outstanding_amount")
t_adj = t_adj + rs("adjustment")

t_amount = t_amount + rs("amount")

t_cr = t_cr + rs("cash_receipt")
tax=tax +rs("tax")

if rs("status") = "Voided" then 
voidamt=voidamt + formatcurrency(rs("amount"),2)
voidamttax=voidamttax + formatcurrency(rs("tax"),2) 
end if
'response.end
%>
 
  <%if rs("status") = "Paid" then %>
  <tr bgcolor="#ffffff"> 
    <% else %>
  <tr bgcolor="#f0f6e3"> 
    <%end if%>
    <td> <%=rs("customer")%></td>
    <td> <a href="bliview.asp?jid=<%=j&"&invoiceid="&rs("invoice")&"&invoicedate="&rs("generation_date")&"&invoicecustomer="&rs("customer")&"&c="&c%>"><%=rs("invoice")%></a> 
    </td>
    <td> <%=rs("generation_date")%> </td>
    <td> <%=rs("status")%></td>
    <td align="right">&nbsp;<%=formatcurrency(rs("amount"),2)%></td>
    <td align="right">&nbsp;<%=formatcurrency(rs("amount")-rs("tax"),2)%></td>
	<td align="right">&nbsp;<%=formatcurrency(rs("adjustment"),2)%></td>
    <td align="right"><%=formatcurrency(rs("cash_receipt"),2)%></td>
    <td align="right">&nbsp;<%=formatcurrency(rs("outstanding_amount"),2)%></td>
    <%
 rs.movenext


wend
%>
  <tr bgcolor="#ffffff"> 
    <td style="border-top:1px solid #000000;border-left:1px solid #000000;border-bottom:1px solid #000000;" colspan="4"><b>Job 
      # <%=j%></b></td>
    
     <td style="border-top:1px solid #000000;border-bottom:1px solid #000000;" align="right">&nbsp;<b><%=formatcurrency(t_amount-voidamt,2)%></b></td>
	<td style="border-top:1px solid #000000;border-bottom:1px solid #000000;" align="right">&nbsp;<b><%=formatcurrency(t_amount-tax-voidamttax,2)%></b></td>
	<td style="border-top:1px solid #000000;border-bottom:1px solid #000000;" align="right">&nbsp;<b><%=formatcurrency(t_adj,2)%></b></td>
    <td style="border-top:1px solid #000000;border-bottom:1px solid #000000;" align="right"><b><%=formatcurrency(t_cr,2)%></b></td>
    <td style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;" align="right">&nbsp;<b><%=formatcurrency(t_outstanding,2)%></b></td>
  </tr>
</table>
<table border=0 cellpadding="3" cellspacing="0">
<tr>
 <td><b>Net Billed Updated as of <%=formatdatetime(crdate,0)%>: <font color="#CC0000"><%=formatcurrency(t_amount-voidamt+t_adj,2)%></font></b></td>

</tr>
<tr>
  <td>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td><div style="position:inline;width:18px;height:12px;background:#f0f6e3;border:1px solid #999999;">&nbsp;</div></td>
    <td width="6">&nbsp;</td>
    <td>Shaded rows indicate voided or unpaid invoices</td>
  </tr>
  </table>  
  </td>
</tr>
<tr>
  <td><input type="button" value="Close Window" onclick="window.close();"></td>
</tr>
</table>  
<%

set cnn = nothing %>
    
</body>
</html>