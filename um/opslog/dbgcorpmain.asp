<html>
<head>
<%@Language="VBScript"%>
<%
customer=Request("customer")
contact=Request("contact")
job= Request("job")
flag= Request("day")
description=Request("description")
company=Request.Form("company")
ReDim Categorys(5)
ReDim Categorysbh(5)
ReDim Categorysot(5)

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")
sqlstr = "select i.* ,case when j.[entry id] > 6283 then left([entry type],2)+'-00'+convert(varchar(4),[entry id]) else '00-00'+convert(varchar(4),[entry id]) end as tjob from invoice_submission i join [job log]j on i.jobno=j.[entry id] where i.jobno='"& job &"' and invoice_date='" & flag &"'" 

rst1.Open sqlstr, cnn1, 0, 1, 1
if rst1.eof then
  response.Write(sqlstr)
  response.End()
end if
tjob=rst1("tjob")
%>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function updateEntry(key, d, job, flag){
	document.frames.detail.location="corpinvoicemodify.asp?id="+key+"&day="+d+"&job="+job+"&flag="+flag
}
function printinvoice(tjob,job,invoice_date,customer,contact,description){
	var temp="invoicereport.asp?tjob="+tjob+"&job="+job+"&day="+invoice_date+"&customer="+customer+"&contact="+contact+"&d="+description
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

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">


<table width="100%" border="0" cellspacing="0">
  <tr bgcolor="#3399CC">
    <td width="57%" height="2" > 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b>Invoice 
        for 
        <% if rst1("contract") then 
	  response.write "Contract" 
	  else
	  response.write "T&M" 
	  end if 
	  %>
        Job No. <%=job%> Invoice date: <%=flag%></b></font></div>
      
    </td>
</tr></table>

 
<% 	
rst1.close

sqlstr2 = "select DISTINCT cat.category as category, sum(cat.hours) as hours, sum(cat.billh) as billh, sum(overt) as overt from (select employees.category as category, sum(hours) as hours, sum(hours_bill) as billh, sum(overt) as overt from invoice_submission, [job log], [employees] where jobno=[entry id] and username=matricola and invoice_submission.invoice_date = '" & flag & "' and invoice_submission.jobno='" & job & "' and invoice_submission.category = 0 group by employees.category union select invoice_submission.category as category, sum(hours) as hours, sum(hours_bill) as billh, sum(overt) as overt from invoice_submission, [job log], [employees] where jobno=[entry id] and username=matricola and invoice_submission.invoice_date = '" & flag & "' and invoice_submission.jobno='" & job & "' group by invoice_submission.category) as cat group by cat.category"

 rst1.Open sqlstr2, cnn1, 0, 1, 1


	
	while not rst1.eof
	
		categorys(rst1("category")) = rst1("hours")
		categorysbh(rst1("category")) = rst1("billh")
		categorysot(rst1("category")) = rst1("overt")
		
		rst1.movenext
	
	wend
	rst1.close
	sqlstr2 = "select invoice_comment, sum(hours) as hours, sum(hours_bill) as hours_bill from invoice_submission where jobno='"&job&"' and invoice_date='"& flag &"' group by invoice_comment"
	
	rst1.Open sqlstr2, cnn1, 0, 1, 1
    if not rst1.eof then
		hours=Trim(rst1("hours"))
		totalbillhours=Trim(rst1("hours_bill"))
		end if
%>
<font face="Arial, Helvetica, sans-serif"><i><b> 
<input type="hidden" name="description" value="<%=description%> ">
<input type="hidden" name="invoice_date" value="<%=flag%>">
<input type="hidden" name="jobno" value="<%=job%>">
<input type="hidden" name="tjobno" value="<%=tjob%>">
<input type="hidden" name="customer" value="<%=customer%>">
<input type="hidden" name="contact" value="<%=contact%>">
</b></i></font> 
<table width="100%">
<tr>
    <td width="83%" height="2" > <font face="Arial, Helvetica, sans-serif"><b><i>Customer: 
      <%=customer%>, <%=contact%></i></b></font></td>
	  
    <td width="17%"> 
      <div align="right">
        <input type="button" name="Print" value="Print Invoice" onclick="printinvoice(tjobno.value,jobno.value,invoice_date.value, customer.value,contact.value,description.value)">
      </div>
    </td>
</tr>
</table>
<form name="frm" method="post" action="file:///K|/um/opslog/invoiceupdatecmt.asp">

  <table width="100%">
    <tr> 
      <td width="50%"> <font face="Arial, Helvetica, sans-serif"><i><b><%=description%> 
        </b></i> </font></td>
      <td width="50%"><font face="Arial, Helvetica, sans-serif">Invoice Comment 
        &nbsp </font></td>
    </tr>
    <tr> 
      <td width="50%"> 
        <table width="100%" border="0">
          <tr> 
            <td width="84%"><font face="Arial, Helvetica, sans-serif">Total Hours</font></td>
            <td bgcolor="#CCCCCC" width="16%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif"><%=hours%> 
                </font></div>
            </td>
          </tr>
          <tr> 
            <td width="84%"><font face="Arial, Helvetica, sans-serif">Total Billable</font></td>
            <td bgcolor="#CCCCCC" width="16%"> 
              <div align="right"> <font face="Arial, Helvetica, sans-serif"><%=totalbillhours%> 
                </font></div>
            </td>
          </tr>
          <tr> 
            <td height="24" width="84%">
              <table width="100%" border="0">
                <tr> 
                  <td bgcolor="#339999" width="20%"> 
                    <div align="center"><b><a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="list"><%=Categorys(5)%></a> 
                      | <a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="list"><%=Categorysbh(5)%></a> 
                      | <a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="list"><%=Categorysot(5)%></a> 
                      </b></div>
                  </td>
                  <td bgcolor="#00FF00" width="20%"> 
                    <div align="center"><b><a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="list"><%=Categorys(1)%></a> 
                      | <a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="list"><%=Categorysbh(1)%></a> 
                      | <a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#1" target="list"><%=Categorysot(1)%></a></b></div>
                  </td>
                  <td bgcolor="#00CC00" width="20%"> 
                    <div align="center"><b><a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#2" target="list"><%=Categorys(2)%> 
                      | <%=Categorysbh(2)%> | <%=Categorysot(2)%></a></b></div>
                  </td>
                  <td bgcolor="#3399CC" width="20%"> 
                    <div align="center"><b><a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#3" target="list"><%=Categorys(3)%></a> 
                      | <a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#3" target="list"><%=Categorysbh(3)%></a> 
                      | <a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#3" target="list"><%=Categorysot(3)%></a></b></div>
                  </td>
                  <td bgcolor="#FF0000" width="20%"> 
                    <div align="center"><b><a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#4" target="list"><%=Categorys(4)%></a> 
                      | <a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#4" target="list"><%=Categorysbh(4)%></a> 
                      | <a href="file:///K|/um/opslog/corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>#4" target="list"><%=Categorysot(4)%></a></b></div>
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#339999" width="20%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Admin</font></b></div>
                  </td>
                  <td bgcolor="#00FF00" width="20%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Entry</font></b></div>
                  </td>
                  <td bgcolor="#00CC00" width="20%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Junior</font></b></div>
                  </td>
                  <td bgcolor="#3399CC" width="20%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Mid</font></b></div>
                  </td>
                  <td bgcolor="#FF0000" width="20%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Senior</font></b></div>
                  </td>
                </tr>
              </table>
            </td>
            <td bgcolor="#CCCCCC" height="24" width="16%"> 
              <div align="right"></div>
            </td>
          </tr>
        </table>
      </td>
      <td width="50%" bgcolor="#CCCCCC" valign="top"> <font face="Arial, Helvetica, sans-serif"> 
        <textarea name="invoicecomment" cols="50" rows="5"><%=rst1("invoice_comment")%></textarea>
        <input type="hidden" name="comment11" value="<%=rst1("invoice_comment")%>">
		<input type="hidden" name="day" value="<%=flag%>">
		<input type="hidden" name="job" value="<%=job%>">
        <input type="submit" name="com" value="Update" >
        </font></td>
    </tr>
  </table>
</form>
<IFRAME name="list" width="100%" height="200" src="corpinvoicedetail.asp?day=<%=flag%>&job=<%=job%>&description=<%=description%>&comment=<%=rst1("invoice_comment")%>&customer=<%=customer%>&contact=<%=contact%>" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>
<IFRAME name="detail" width="100%" height="150" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>
</body>
</html>
