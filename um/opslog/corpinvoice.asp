<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%'1/15/2008 N.Ambo modified page to add new column on screen for the invoice amount %>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
function process(item, flag, d, job){
	var count=document.temp.count.value-1
	document.location="corpinvoicefilter.asp?job="+job+"&flag="+flag+"&date="+d+"&item="+item
}
function rejectinv(jobnum, d) {

	var temp = "rejectinv.asp?job=" + jobnum + "&d="+ d
	window.open(temp,"", "scrollbars=yes,width=600, height=300, status=no, menubar=no" );


}

//visual feedback functions for img buttons
function buttonOver(obj,clr){
  if (arguments.length == 1) { clr = "#336699"; }
  obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
  if (arguments.length == 1) { clr = "#000000"; }
  obj.style.border = "1px solid " + clr;
}

function buttonOut(obj,clr){
  if (arguments.length == 1) { clr = "#eeeeee"; }
  obj.style.border = "1px solid " + clr;
}

function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")

' no search needed for submitted invoice, so flag handlers are disabled

'flag=request("date")
'item=request("item")
'if  item= "day" then
'	sql="select invoice_date, jobno, master_job.description,master_job.job,master_job.customer_name,sum(billable) as total from master_job, invoice_submission where submitted=1 and flag=0  and invoice_date>='" & flag & "'and master_job.id=jobno group by invoice_date, jobno, master_job.description,master_job.job,master_job.customer_name"
'else
'	if item = "job" then
'   	sql="select distinct jobno, master_job.description, master_job.job,master_job.customer_name,invoice_date,customers.companyname,sum(billable) as total from master_job, invoice_submission where submitted=1 and flag=0 and jobno='" &flag& "' and jobno=master_job.id group by invoice_date, jobno , master_job.description,master_job.job,customer_name"	
'	else
        sql="select invoice_submission.submittedby,invoice_submission.invoice_comment, invoice_submission.invoice_date, invoice_submission.jobno, master_job.job,master_job.description ,master_job.job,master_job.customer,sum(billable)as total, invoice_submission.invoice_amt from invoice_submission , master_job where invoice_submission.jobno=master_job.[id] and submitted=1 and  flag=0 and closed=0 and invoice_submission.jobno=master_job.[id] group by invoice_submission.invoice_date, invoice_submission.invoice_comment, invoice_submission.jobno, master_job.job,master_job.description,master_job.job,master_job.customer,invoice_submission.submittedby, invoice_submission.invoice_amt"
     
'	end if
'end if

'response.write sql
'response.end
rst1.Open sql, cnn1, 0, 1, 1
%>

<body bgcolor="#FFFFFF" text="#000000">

<%
if not rst1.eof then
%>

<table border="0" cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
	<tr bgcolor="dddddd">
		<td colspan="4" align="center" style="border-bottom:1px solid #aaaaaa">
			<b>Submitted Invoices</b>
		</td>
	</tr>
	<tr bgcolor="#dddddd"> 
		<td style="border-top:1px solid #ffffff;">Invoice Date</td>
		<td style="border-top:1px solid #ffffff;" align="right">Invoice Amount ($)</td> <%'Added 1/15/2008 N.Ambo%>
		<td style="border-top:1px solid #ffffff;">Job Number</td>
		<td style="border-top:1px solid #ffffff;">Description</td>
		<td style="border-top:1px solid #ffffff;">&nbsp; </td>
	</tr>

  <%
  dim count,d,longjob,job,description,customer,userR, invoiceAmt 'invoice amt added 1/15/2008 N.Ambo to add column to screen
  count=0
  do until rst1.eof
  d=rst1("invoice_date")
  job=rst1("jobno")
  longjob=rst1("job")
  description=rst1("description")
  'contact=rst1("contact name")
  customer=rst1("customer")
  submUser=rst1("submittedby")
  invoiceAmt = rst1("invoice_amt") 'Added 1/15/2008 N.Ambo
 if not rst1.eof then 
  dim  rstt, strsql,stat,moneyamt,submUser
	'set cnn = server.createobject("ADODB.connection")
	set rstt = server.createobject("ADODB.recordset")
	'cnn.open getConnect(0,0,"intranet")
    
	strsql = "SELECT * FROM MASTER_JOB mj inner join taxstatuslist t on  t.id = mj.taxstatusid WHERE mj.id='"&split(longjob,"-")(1)&"'"

	rstt.Open strsql,cnn1
	if not rstt.EOF then
    moneyamt=rstt("amt_1")
    stat= rstt("Status")
  
    end if
	
end if
  
  %>
  <tr bgcolor="#ffffff">
    <form>
      <td><a href="corpmain.asp?source=corpinvoice&day=<%=d%>&job=<%=longjob%>&description=<%=description%>&customer=<%=customer%>"><%=d%></a></td>
      <input type="hidden" name="d" value="<%=d%>">
	  <input type="hidden" name="job" value="<%=job%>">
      <input type="hidden" name="flag" value="<%=flag%>"> 
      <input type="hidden" name="item" value="<%=item%>">
      <!--<td><%'=longjob%></td>-->
    <td align="right"><%=invoiceAmt%></td> <%'Added 1/15/2008 N.Ambo%>
   <td> <a href="javascript:openwin('/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=<%=split(longjob,"-")(1)%>',850,500)" onmouseover="this.T_WIDTH=180;this.T_OPACITY=80;this.T_TEXTALIGN='center';this.T_SHADOWWIDTH=5;this.T_TITLE='<%=longJob%>';return escape('Job Amount:<%=formatcurrency(moneyamt,2)%><br>Job Status:<%=stat%><br>Submitted By:<%=userfullname(submUser)%>');"><%=longjob%></a></td>
	  <td><%=description%></td>
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
			<% if (allowgroups("Genergy_Corp")) then %>
          		<input type="button" name="Submit" value="APPROVE" onclick="process(item.value, flag.value, d.value, <%=job%>)">
          		<input type="button" name="Button" value="REJECT" onclick="rejectinv(job.value, d.value)">
			<% end if%>
        </div>
      </td>
    </form>
  </tr>
  <%
  rstt.movenext
  rst1.movenext
  count=count+1  
  loop
  rstt.close 
  %>

</table>
<form name="temp">
<input type="hidden" name="count" value="<%=count%>">
</form>

<%
else
	rst1.close
	if(item = "job") then
		sql="select id from master_job where id='" &flag& "'"
		rst1.Open sql, cnn1, 0, 1, 1
		if rst1.eof then
%>
<br><div class="notetext" style="padding:10px;">No such job</div>
<%
 		else
%> 
<br><div class="notetext" style="padding:10px;">No invoices are waiting</div> 
<%
		end if
	else
%>
	<br><div class="notetext" style="padding:10px;">No invoices pending approval</div> 
<%
	end if
end if
%>
<script language="JavaScript" type="text/javascript" src="/genergy2/wz_tooltip.js"></script>
</body>
</html>