<html>
<head>
<%@Language="VBScript"%>
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim customer_name,contact,job,icomment,buttontxt,buttonaction
customer=Request("customer")
tjob= Request("job")
job=cint(right(trim(tjob),6))
flag= Request("day")
if request("viewtimes")<>"yes" then
  buttontxt=""
  buttonaction="?customer="&customer&"&job="&tjob&"&day="&flag&"&description="&description
else
  buttontxt="View Filed Invoices Status"
  buttonaction="?customer="&customer&"&job="&tjob&"&day="&flag&"&description="&description
end if
description=Request("description")

dim sourceF
sourceF=request("source")


ReDim Categorys(5)
ReDim Categorysbh(5)
ReDim Categorysot(5)

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")
'sqlstr = "select i.* ,j.job as tjob from invoice_submission i join [master_job]j on i.jobno=j.[id] where i.jobno='"& job &"' and invoice_date='" & flag &"'" 
'rst1.open "SELECT DISTINCT a.*, b.job,b.Customer, b.Customer_Name, b.amt_1, b.amt_2, d.Contact, d.Contact_Name " & _
'		  "FROM invoice_submission as a INNER JOIN MASTER_JOB as b ON a.Jobno = b.id " & _
'		  "INNER JOIN gy_MASTER_ARM_CUSTOMER as c ON b.Customer = c.Customer INNER JOIN GY_STANDARD_ARS_CONTACT as d ON c.Billing_Contact = d.Contact " & _
'		  "WHERE a.Jobno = "&job
dim somesql
somesql = "SELECT a.Customer_Name,a.job,a.company,a.amt_1, a.amt_2, c.Contact, c.Contact_Name FROM MASTER_JOB as a INNER JOIN gy_MASTER_ARM_CUSTOMER as b ON a.Customer = b.Customer INNER JOIN GY_STANDARD_ARS_CONTACT as c ON b.Billing_Contact = c.Contact WHERE a.id = "&job
'response.write somesql
rst1.open somesql,cnn1
'response.write "SELECT a.Customer_Name,a.job,a.company,a.amt_1, a.amt_2, c.Contact, c.Contact_Name FROM MASTER_JOB as a INNER JOIN gy_MASTER_ARM_CUSTOMER as b ON a.Customer = b.Customer INNER JOIN GY_STANDARD_ARS_CONTACT as c ON b.Billing_Contact = c.Contact WHERE a.id = "&job
'response.end
		  'amounts not used yet

if not rst1.EOF then
  customer_name=rst1("customer_name")
  contact=rst1("contact_name")
  company=rst1("company")
  
else
  customer_name=""
  contact=""
  company=""
end if

'response.write "company:" & company
%>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function updateEntry(key, d, job, flag){
	document.frames.detail.location="corpinvoicemodify.asp?id="+key+"&day="+d+"&job="+job+"&flag="+flag
}
function printinvoice(tjob,job,invoice_date,customer,contact,description, cphone, cEmail){
	var temp="invoicereport.asp?tjob="+tjob+"&job="+job+"&day="+invoice_date+"&customer="+customer+"&contact="+contact+"&d="+description+"&cphone="+cphone+"&cEmail="+cEmail+"&company=<%=company%>"
	window.open(temp,"", "scrollbars=yes,width=800, height=600, status=no, menubar=no" );
}

function invoice(job){
    var currdate = new Date()
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	document.location="corpinvoiceupdate.asp?job="+job+"&day="+currdate
}
function commentinv(jobno,invoice_date){
	document.location="invoiceupdatecmt.asp?jobno="+jobno+"&invoice_date="+invoice_date
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#eeeeee" text="#000000">


  <% 	
  rst1.close
  
  sqlstr2 = "select DISTINCT cat.category as category, sum(cat.hours) as hours, sum(cat.billh) as billh, sum(overt) as overt from (select employees.category as category, sum(hours) as hours, sum(hours_bill) as billh, sum(overt) as overt from invoice_submission, [job log], [employees] where jobno=[entry id] and username=matricola and invoice_submission.invoice_date = '" & flag & "' and invoice_submission.jobno='" & job & "' and invoice_submission.category = 0 group by employees.category union select invoice_submission.category as category, sum(hours) as hours, sum(hours_bill) as billh, sum(overt) as overt from invoice_submission, [job log], [employees] where jobno=[entry id] and username=matricola and invoice_submission.invoice_date = '" & flag & "' and invoice_submission.jobno='" & job & "' group by invoice_submission.category) as cat group by cat.category"
'  response.write sqlstr2
'  response.end
   rst1.Open sqlstr2, cnn1, 0, 1, 1
  
  
  
    
    while not rst1.eof
    
      categorys(rst1("category")) = rst1("hours")
      categorysbh(rst1("category")) = rst1("billh")
      categorysot(rst1("category")) = rst1("overt")
      
      rst1.movenext
    
    wend
    rst1.close
    sqlstr2 = "select invoice_comment, isnull(invoice_amt,0) as invoice_amt, cName, cTelephone, cEmail, sum(hours) as hours, sum(hours_bill) as hours_bill from invoice_submission where jobno='"&job&"' and invoice_date='"& flag &"' group by invoice_comment, cName, cTelephone, cEmail, invoice_amt"
    
    rst1.Open sqlstr2, cnn1, 0, 1, 1
      if not rst1.eof then
        icomment=rst1("invoice_comment")
      	hours=Trim(rst1("hours"))
      	totalbillhours=Trim(rst1("hours_bill"))
	  	invCname = rst1("cname")
	  	invCtelephone = rst1("ctelephone")
	  	invCemail = rst1("cemail")
		total_amt = rst1("invoice_amt")
      end if
    rst1.close
  %>
  <form name="frm" method="post" action="invoiceupdatecmt.asp">
  <input type="hidden" name="description" value="<%=description%> ">
  <input type="hidden" name="invoice_date" value="<%=flag%>">
  <input type="hidden" name="jobno" value="<%=job%>">
  <input type="hidden" name="tjobno" value="<%=tjob%>">
  <input type="hidden" name="customer" value="<%=customer_name%>">
  <input type="hidden" name="contact" value="<%=contact%>">
  <input type="hidden" name="showhist" value=1>
  	<input type="hidden" name="source" value="<%=sourceF%>">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#eeeeee">
  <td height="1" width="71%"><b><a href="javascript:openwin('/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=<%=job%>',550,400)">Invoice for Job # <%=job%></a></b></td>
  <td align="right"><% if buttontxt<>"" then %><input type="button" value="<%=buttontxt%>" onClick="document.location='corpmain.asp<%=buttonaction%>'"><% end if %><input type="button" name="Print" value="Print Invoice" onclick="printinvoice(tjobno.value,jobno.value,invoice_date.value, customer.value, contact.value, description.value, invCtelephone.value, invCemail.value)" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
</tr>
<tr bgcolor="#eeeeee">
  <td colspan="2">
  
  <table border=0 cellpadding="3" cellspacing="0">
        <tr valign="top"> 
          <td width="80">Invoice Date:</td>
          <td width="326"><%=flag%></td>
        </tr>
        <tr valign="top"> 
          <td>Customer:</td>
          <td> 
            <%
    if customer_name="" then
      response.write("Customer or Billing Contact Not Designated")
    else %>
            <%=customer_name%>&nbsp;(<%=customer%>)<br>
            <%end if%>
          </td>
        </tr>
        <tr valign="top">
          <td>Contact:</td>
          <td><table width="100%" border="0">
              <tr>
                <td width="26%">Name</td>
                <td width="74%"><input name="invCname" type="text" value="<%=invCname%>"></td>
              </tr>
              <tr>
                <td>Telephone</td>
                <td><input name="invCtelephone" type="text" value="<%=invCtelephone%> "></td>
              </tr>
              <tr>
                <td>Email</td>
                <td><input name="invCemail" type="text" value="<%=invCemail%>"></td>
              </tr>
            </table></td>
          <td rowspan="6" width="20">&nbsp;</td>
          <td width="339">&nbsp;</td>
        </tr>
        <tr valign="top"> 
          <td>Description:</td>
          <td><%=description%></td>
          <td>Comment</td>
        </tr>
        <tr valign="top"> 
          <td>Hours:</td>
          <td> 
            <!-- begin color coded hours table -->
            <table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc"width="300">
              <tr> 
                <td align="right" bgcolor="#dddddd"></td>
                <td bgcolor="#66ff66" width="20%">Entry</td>
                <td bgcolor="#339999" width="20%">Junior</td>
                <td bgcolor="#ff9900" width="20%">Mid</td>
                <td bgcolor="#cc0000" width="20%"><span style="color:#ffffff;">Senior</span></td>
                <td bgcolor="#666699" width="20%"><span style="color:#ffffff;">Admin</span></td>
              </tr>
              <tr bgcolor="#ffffee"> 
                <td align="right" bgcolor="#dddddd">Regular:</td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="detail"><%=Categorys(1)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#2" target="detail"><%=Categorys(2)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#3" target="detail"><%=Categorys(3)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#4" target="detail"><%=Categorys(4)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="detail"><%=Categorys(5)%></a>&nbsp;</td>
              </tr>
              <tr bgcolor="#f0f0e0"> 
                <td align="right" bgcolor="#dddddd">Billable:</td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="detail"><%=Categorysbh(1)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#2" target="detail"><%=Categorysbh(2)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#3" target="detail"><%=Categorysbh(3)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#4" target="detail"><%=Categorysbh(4)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="detail"><%=Categorysbh(5)%></a>&nbsp;</td>
              </tr>
              <tr bgcolor="#e3e3d3"> 
                <td align="right" bgcolor="#dddddd">Overtime:</td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="detail"><%=Categorysot(1)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#2" target="detail"><%=Categorysot(2)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#3" target="detail"><%=Categorysot(3)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#4" target="detail"><%=Categorysot(4)%></a></td>
                <td><a href="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="detail"><%=Categorysot(5)%></a>&nbsp;</td>
              </tr>
            </table>
            <!-- end color coded hours table -->
          </td>
          <td rowspan="3"> <textarea name="invoicecomment" cols="50" rows="5"><%=icomment%></textarea>
            <br> <input type="hidden" name="comment11" value="<%=icomment%>"> 
            <input type="hidden" name="day" value="<%=flag%>"> <input type="hidden" name="job" value="<%=job%>">
          </td>
        </tr>
        <tr> 
          <td>Total Hours:</td>
          <td><%=hours%> </td>
        </tr>
        <tr> 
          <td>Total Billable:</td>
          <td><b><%=totalbillhours%></b></td>
        </tr>
		<tr>
			<td>Total Amount:</td>			
			<td>
				<%if allowgroups("Genergy_Corp") then%>
					<input type="text" name="tot_amt" value="<%=formatcurrency(total_amt,2)%>" size="4">
				<%else%>
					<%=formatcurrency(total_amt,2)%><input type="hidden" name="tot_amt" value="<%=total_amt%>">
				<%end if%>
			</td>
		</tr>
      </table>
  </td>
</tr>
<tr>
<td>
<input type="submit" name="com" value="Update Invoice Details" style="border:1px outset #ddffdd;background-color:ccf3cc;"></td>
</tr>
<tr>
  <td colspan="2">
  <IFRAME name="detail" width="100%" height="130" src="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>&comment=<%=icomment%>&customer=<%=customer%>" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 style="border:1px solid #cccccc;"></IFRAME>
  <IFRAME name="list" width="100%" height="130" src="/um/war/ara/invoice.asp?c=<%=company%>&j<%=left(company,1)&"="&tjob%>" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 style="border:1px solid #cccccc;"></IFRAME>
 
  </td>
</tr> 
</table> 
</form>
</body>
</html>