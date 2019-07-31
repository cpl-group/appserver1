<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function process(item, flag, d, job){
	var count=document.temp.count.value-1
	document.location="corpinvoicefilter.asp?job="+job+"&flag="+flag+"&date="+d+"&item="+item
}
function rejectinv(jobnum, d) {

	var temp = "rejectinv.asp?job=" + jobnum + "&d="+ d
	window.open(temp,"", "scrollbars=yes,width=600, height=300, status=no, menubar=no" );


}
</script>
</head>
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open application("cnnstr_main")
flag=request("date")
item=request("item")
if  item= "day" then
	sql="select invoice_date, jobno, [job log].description,customers.companyname,[job log].[contact name],sum(billable) as total from [job log], invoice_submission,customers where submitted=1 and flag=0  and invoice_date>='" & flag & "'and customers.customerid=[job log].customer and [entry id]=jobno group by invoice_date, jobno, [job log].description,customers.companyname,[job log].[contact name],[job log].customer"
	
else
	if item = "job" then
    	sql="select distinct jobno, [job log].description, invoice_date,customers.companyname,[job log].[contact name],sum(billable) as total from [job log], invoice_submission,customers where submitted=1 and flag=0 and jobno='" &flag& "'and customers.customerid=[job log].customer and jobno=[entry id] group by invoice_date, jobno , [job log].description,customers.companyname,[job log].[contact name],[job log].customer"
	
	else
    sql="select invoice_submission.invoice_date, invoice_submission.jobno, [job log].description ,customers.companyname,[job log].[contact name],sum(billable)as total from invoice_submission join [job log] on invoice_submission.jobno=[job log].[entry id]join customers on [job log].customer=customers.customerid where submitted=1 and  flag=0 and invoice_submission.invoice_date > isnull( [job log].billdate,'1/1/2001') and invoice_submission.jobno=[job log].[entry id] group by invoice_submission.invoice_date, invoice_submission.jobno, [job log].description,customers.companyname,[job log].[contact name],[job log].customer"
	end if
end if

'response.write sql
'response.end
rst1.Open sql, cnn1, 0, 1, 1
%>

<body bgcolor="#FFFFFF" text="#000000">
<%
if not rst1.eof then
%>

<table width="100%" border="0" bgcolor="#3399CC">
  <tr >
    <td height="2" width="71%"><font face="Arial, Helvetica, sans-serif"><i><b> 
      &nbspCurrent Invoices </b></i></font> </td>
  </tr>
</table> 

<div align="right"> </div>

<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td><font face="Arial, Helvetica, sans-serif">Invoice Date</font></td>
    <td><font face="Arial, Helvetica, sans-serif">Job Number</font></td>
    <td><font face="Arial, Helvetica, sans-serif">Description</font></td>
    <td><font face="Arial, Helvetica, sans-serif">&nbsp </font></td>
  </tr>
  <%
  count=0
  do until rst1.eof
  d=rst1("invoice_date")
  job=rst1("jobno")
  description=rst1("description")
  contact=rst1("contact name")
  customer=rst1("companyname")
  %>
  <tr>
    <form>
      <td><a href="corpmain.asp?day=<%=d%>&job=<%=job%>&description=<%=description%>&contact=<%=contact%>&customer=<%=customer%>"><font face="Arial, Helvetica, sans-serif"><%=d%></font></a></td>
      <input type="hidden" name="d" value="<%=d%>">
	  <input type="hidden" name="job" value="<%=job%>">
      <input type="hidden" name="flag" value="<%=flag%>">
      <input type="hidden" name="item" value="<%=item%>">
      <td><font face="Arial, Helvetica, sans-serif"><a href="opslogview.asp?job=<%=job%>"><%=job%></a></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><%=description%></font></td>
      <% 
    total=rst1("total")
	if total>=0 then
	else
		total=0
	end if
	total=formatcurrency(total, 2)    
%>
      <td> 
        <div align="center"> 
          <input type="button" name="Submit" value="APPROVE" onclick="process(item.value, flag.value, d.value, <%=job%>)">
          <input type="button" name="Button" value="REJECT" onclick="rejectinv(job.value, d.value)">
        </div>
      </td>
    </form>
  </tr>
  <%
  rst1.movenext
  count=count+1  
  loop
  
  %>
</table>
<form name="temp">
<input type="hidden" name="count" value="<%=count%>">
</form>
<font face="Arial, Helvetica, sans-serif">
<%
else
	rst1.close
	if(item = "job") then
		sql="select [entry id] from [job log] where [entry id]='" &flag& "'"
		rst1.Open sql, cnn1, 0, 1, 1
		if rst1.eof then
%>
<br><center><b><h2>No Such Job</b></center>
<%
 		else
%> 
<br><center><b><h2>No Invoices are waiting</b></center> 
<%
		end if
	else
%>
	<br><center><b><h2>No Invoices are waiting</b></center> 
<%
	end if
end if
%>
</font> 
</body>
</html>
