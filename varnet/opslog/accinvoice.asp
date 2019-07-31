<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function process(item, flag, d, job){
	document.location="accinvoicefilter.asp?job="+job+"&flag="+flag+"&date="+d+"&item="+item
}
</script>
</head>
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
d=request("date")
item=request("item")
if  item= "day" then
	sql="select invoice_submission.invoice_date, invoice_submission.jobno, [job log].description,customers.companyname,[job log].[contact name],sum(billable)as total from invoice_submission join [job log] on invoice_submission.jobno=[job log].[entry id]join customers on [job log].customer=customers.customerid where invoice_date>='"& d &"' and flag=1 and closed=1 and  invoice_submission.jobno=[job log].[entry id] group by invoice_submission.invoice_date, invoice_submission.jobno, [job log].description,customers.companyname,[job log].[contact name]"
else
	if item = "job" then
		sql="select invoice_submission.invoice_date, invoice_submission.jobno, [job log].description,customers.companyname,[job log].[contact name],sum(billable)as total from invoice_submission join [job log] on invoice_submission.jobno=[job log].[entry id]join customers on [job log].customer=customers.customerid where jobno='" &d& "'and flag=1 and closed=1 and  invoice_submission.jobno=[job log].[entry id] group by invoice_submission.invoice_date, invoice_submission.jobno, [job log].description,customers.companyname,[job log].[contact name]"
	else
    sql="select invoice_submission.invoice_date, invoice_submission.jobno, [job log].description,customers.companyname,[job log].[contact name],sum(billable)as total from invoice_submission join [job log] on invoice_submission.jobno=[job log].[entry id]join customers on [job log].customer=customers.customerid where flag=1 and closed=0  and invoice_submission.jobno=[job log].[entry id] group by invoice_submission.invoice_date, invoice_submission.jobno, [job log].description,customers.companyname,[job log].[contact name]"

	end if
end if
'response.write sql

rst1.Open sql, cnn1, 0, 1, 1
%>

<body bgcolor="#FFFFFF" text="#000000">
<%
if not rst1.eof then
%>
<table width="100%" border="0" bgcolor="#3399CC">
  <tr > 
    <td height="1" width="71%"> <b><font face="Arial, Helvetica, sans-serif">&nbsp<font color="#FFFFFF">Current 
      Invoices</font></font></b><i><b><font face="Arial, Helvetica, sans-serif"><font color="#FFFFFF"> 
      </font></font></b></i> </td>
  </tr>
</table> 
<div align="right"> </div>

<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td><font face="Arial, Helvetica, sans-serif">Invoice Date</font></td>
    <td><font face="Arial, Helvetica, sans-serif">Job Number</font></td>
    <td><font face="Arial, Helvetica, sans-serif">Description</font></td>
    <td>&nbsp;</td>
  </tr>
  <%
  do until rst1.eof
  d=rst1("invoice_date")
  job=rst1("jobno")
  description=rst1("description")
  contact=rst1("contact name")
  customer=rst1("companyname")
  %>
  <form>
    <tr> 
      <td><font face="Arial, Helvetica, sans-serif"><a href="corpmain.asp?day=<%=d%>&job=<%=job%>&description=<%=description%>&contact=<%=contact%>&customer=<%=customer%>"><%=d%></a></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><a href="opslogview.asp?job=<%=job%>"><%=job%></a></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><%=rst1("description")%></font></td>
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
<font face="Arial, Helvetica, sans-serif">
<%
else
	rst1.close
	if(item = "job") then
		sql="select [entry id] from [job log] where [entry id]='" &d& "'"
		rst1.Open sql, cnn1, 0, 1, 1
		if rst1.eof then
%>
<br><center><b><h2>No Such Job</b></center>
<%
 		else
%> 
<br><center><b><h2>No Invoices Found</b></center> 
<%
		end if
	else
%>
	<br><center><b><h2>No Invoices Found</b></center> 
<%
	end if
end if
%>
</font> 
</body>
</html>
