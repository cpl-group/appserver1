<%@Language="VBScript"%>
<% 'COMMMENTS 
'2/4/2008 N.Ambo added column for invoice amount
%> %>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function process(item, flag, d, job){
	document.location="accinvoicefilter.asp?job="+job+"&flag="+flag+"&date="+d+"&item="+item
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")
d=request("date")
item=request("item")
dim pageLabel
pageLabel = "Approved Invoices"
sql="select invoice_submission.invoice_date, invoice_submission.jobno, master_job.job,master_job.description,master_job.customer,sum(billable)as total, invoice_submission.invoice_amt from invoice_submission join master_job on invoice_submission.jobno=master_job.id where "
if  item= "day" then
	sql=sql & "invoice_date>='"& d &"' and flag=1  and  invoice_submission.jobno=master_job.id group by invoice_submission.invoice_date, invoice_submission.invoice_comment, invoice_submission.jobno, master_job.job,master_job.description,master_job.customer, invoice_submission.invoice_amt"
	pageLabel = "Invoice Search Results"
else
	if item = "job" then
		sql=sql & "jobno='" &d& "' and invoice_submission.jobno=master_job.id group by invoice_submission.invoice_date, invoice_submission.jobno, master_job.job,invoice_submission.invoice_comment, master_job.description,master_job.customer, invoice_submission.invoice_amt"
		pageLabel = "Invoice Search Results"
	else   'default view- all submitted( but not closed ) jobs
        sql=sql & "flag=1 and closed=0  and invoice_submission.jobno=master_job.id group by invoice_submission.invoice_date, invoice_submission.jobno, master_job.job,invoice_submission.invoice_comment, master_job.description,master_job.customer, invoice_submission.invoice_amt"
	end if
end if
'response.write sql

rst1.Open sql, cnn1, 0, 1, 1
%>

<body bgcolor="#FFFFFF" text="#000000">
<%
if not rst1.eof then
%>
<!--
<table width="100%" border="0" bgcolor="#3399CC">
  <tr > 
    <td height="1" width="71%"> <b><font face="Arial, Helvetica, sans-serif">&nbsp<font color="#FFFFFF">Current 
      Invoices</b><i><b><font face="Arial, Helvetica, sans-serif"><font color="#FFFFFF"> 
      </b></i> </td>
  </tr>
</table> 
<div align="right"> </div>
-->

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#cccccc">
	<tr bgcolor="#dddddd">
		<td colspan="4" align="center"  style="border-bottom:1px solid #aaaaaa">
			<b><%=pageLabel%></b>
		</td>
	</tr>
  <tr bgcolor="#dddddd"> 
    <td style="border-top:1px solid #ffffff;">Invoice Date</td>
    <td style="border-top:1px solid #ffffff;" align="right">Invoice Amount ($)</td>
    <td style="border-top:1px solid #ffffff;">Job Number</td>
    <td style="border-top:1px solid #ffffff;">Description</td>
    <td style="border-top:1px solid #ffffff;">&nbsp;</td>
  </tr> 
  <%
  
  dim d,job,longjob,description,customer, invoiceAmt
  
  do until rst1.eof
  d=rst1("invoice_date")
  job=rst1("jobno")
  longjob=rst1("job")
  description=rst1("description")
  'contact=rst1("contact name")
  customer=rst1("customer")
  invoiceAmt = rst1("invoice_amt") '2/4/2008 N.Ambo added invocie amount
  %>
  <form>
    <tr bgcolor="#ffffff"> 
      <td><a href="corpmain.asp?source=accinvoice&day=<%=d%>&job=<%=longjob%>&description=<%=description%>&customer=<%=customer%>"><%=d%></a></td>
       <td align="right"><%=invoiceAmt%></td>
      <td><%=longjob%></td>
      <td><%=rst1("description")%></td>
      <% 
    total=rst1("total")
	if total>=0 then
	else
		total=0
	end if
	total=formatcurrency(total, 2)    
%>
      <input type="hidden" name="d" value="<%=d%>">
      <input type="hidden" name="flag" value="<%=flag%>">
      <input type="hidden" name="item" value="<%=item%>">
      <input type="hidden" name="job" value="<%=job%>">
      <td>
	  <%if item<>"day" and item <> "job" then %>
        <input type="button" name="b1" value="Close" onclick="process(item.value, flag.value, d.value, job.value)">
	  <%end if %>
      </td>
    </tr>
  </form>
  <%
  rst1.movenext
  loop  
  %>
</table>

<%
else
	rst1.close
	if(item = "job") then
		sql="select id from master_job where id='" &d& "'"
		rst1.Open sql, cnn1, 0, 1, 1
		if rst1.eof then
%>
<br><div class="notetext" style="padding:10px;">No such job</div>
<%
 		else
%> 
<br><div class="notetext" style="padding:10px;">No invoices found</div> 
<%
		end if
	else
%>
<br><div class="notetext" style="padding:10px;">No invoices found</div> 
<%
	end if
end if
%>
 
</body>
</html>