<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

dim c,j,t_amount,jid


c = request("c")
jid =request("jid") 
j=request("j") 
if trim(c)="" then 
	response.write "<font face=""arial"" size=""2"" color=""red"">This section can not be displayed because <b>no billing contact</b> is specified for this job. </font>"
	response.end
end if


dim cnn, cmd, rs, prm,tally,rst
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient



'rs.open "SELECT s.customer,s.status,s.invoice,a.amount, adjustment,outstanding_amount, cash_receipt, a.Activity_Date as generation_date, isNull(Billed, 0) as Billed FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_ACTIVITY_ARA_ACTIVITY a on s.invoice=a.invoice  left join "&c&"_BILLED_BLI_INVOICE i on i.invoice=s.invoice WHERE a.amount > 0 and a.job='"&j&"' order by s.invoice, s.generation_date",cnn
'rs.open "select * from invoice_submission where submitted=1 and  flag=0 and closed=0 and jobno = '"&jid&"'",cnn

rs.open "select distinct a.invoice_amt,a.invoice_date,a.submittedby,mj.description,mj.customer,a.flag from invoice_submission a inner join master_job mj on a.jobno = mj.id  where submitted=1 and   a.jobno= '" & jid &"'" ,cnn'and mj.id='"& jid &"'",cnn
'response.write "select  distinct * from invoice_submission a inner join master_job mj on a.jobno = mj.id  where submitted=1 and  flag=0 and closed=0 and a.jobno= '" & jid &"'"
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

  <td><span class="standardheader">Review Invoices</span></td>
</tr>
</table>

<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr bgcolor="#dddddd" style="font-weight:bold;"> 
    <td>Invoice Date</td>
    <td>Invoice Amount</td>
    <td>Submitted By</td>
	</tr>
  <%

t_amount = 0


while not rs.EOF 
Dim t_cr
t_amount = t_amount + rs("invoice_amt")

%>

 <%if rs("flag") then %>
 <tr bgcolor="#ffffff">
    <% else %>
     <tr bgcolor="#f0f6e3">
    <%end if%>
   <td> <%=rs("Invoice_date")%></td>
   <td> <a href="javascript:openwin('/um/opslog/corpmain.asp?source=corpinvoice&day=<%=rs("invoice_date")&"&job="& j & "&description="&rs("description")&"&customer=" & rs("customer")%>',1024,768)"><%=formatcurrency(rs("invoice_amt"),2)%></a> 
	</td>
   <% response.write "<td>" & userfullname(rs("submittedby"))&"</td>"%>
 
    <%
 rs.movenext


wend
%>
  <tr bgcolor="#ffffff"> 
    <td style="border-top:1px solid #000000;border-left:1px solid #000000;border-bottom:1px solid #000000;" colspan="4"><b>Job 
      # <%=j%></b></td>
       </tr>
</table>
<table border=0 cellpadding="3" cellspacing="0">
<tr>
  <td><b>Updated as of <%=formatdatetime(date,0)%>: <font color="#CC0000"><%=formatcurrency(t_amount,2)%></font></b></td>
</tr>
<tr>
  <td>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>


<tr>
	<td><div style="position:inline;width:18px;height:12px;background:#ffffff;border:1px solid #999999;">&nbsp;</div></td>
    <td width="6">&nbsp;</td>
    <td>Shaded rows indicate approved invoice(s)</td>
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