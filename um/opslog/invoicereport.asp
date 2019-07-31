<html>
<head>
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
tjob= Request("tjob")
job= Request("job")
flag= Request("day")
description= Request("d")
comment= Request("day")
customer=Request("customer")
contact=Request("contact")
cPhone=Request("cPhone")
cEmail=Request("cEmail")
company=Request("company")
'response.end


ReDim Categorys(5)
ReDim Categorysbh(5)
ReDim Categorysot(5)

Dim cnn1,icomment
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")
%>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>


function invoice(job){
    var currdate = new Date()
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	document.location="corpinvoiceupdate.asp?job="+job+"&day="+currdate
}

</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
.fineprint td { font-size:8pt; }
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" onload="print();">


<table border=0 cellpadding="3" cellspacing="0" width="550">
<tr>
  <td colspan="2" style="border:1px solid #000000;"><b>Invoice for job <%=tjob%></b></td>
</tr>
<tr><td colspan="2" height="8"></td></tr>
<tr>
  <td width="25%">Invoice date:</td>
  <td width="75%"><%=flag%></td>
</tr>
 
<% 	
sqlstr2 = "select DISTINCT cat.category as category, sum(cat.hours) as hours, sum(cat.billh) as billh, sum(overt) as overt from (select employees.category as category, sum(hours) as hours, sum(hours_bill) as billh, sum(overt) as overt from invoice_submission, master_job, [employees] where jobno=master_job.id and username=matricola and invoice_submission.invoice_date = '" & flag & "' and invoice_submission.jobno='" & job & "' and invoice_submission.category = 0 group by employees.category union " & _
          "select invoice_submission.category as category, sum(hours) as hours, sum(hours_bill) as billh, sum(overt) as overt from invoice_submission, master_job, [employees] where jobno=master_job.id and username=matricola and invoice_submission.invoice_date = '" & flag & "' and invoice_submission.jobno='" & job & "' group by invoice_submission.category) as cat group by cat.category"

rst1.Open sqlstr2, cnn1, 0, 1, 1

		while not rst1.eof
	
		categorys(rst1("category")) = rst1("hours")
		categorysbh(rst1("category")) = rst1("billh")
		categorysot(rst1("category")) = rst1("overt")
		
		rst1.movenext
	
	wend
	rst1.close
	
	sqlstr2 = "select invoice_comment, sum(hours) as hours, sum(hours_bill) as hours_bill from invoice_submission where jobno='"&job&"' and invoice_date='"& flag &"' group by invoice_comment"
'	response.write sqlstr2
'  response.end
	rst1.Open sqlstr2, cnn1, 0, 1, 1
  if not rst1.eof then
		hours=Trim(rst1("hours"))
		totalbillhours=Trim(rst1("hours_bill"))
  'Screen invoice comment for problematic characters
  icomment=rst1("invoice_comment")
  icomment=replace(icomment,"<","&lt;")
  icomment=replace(icomment,">","&gt;")
  Set rst2 = Server.CreateObject("ADODB.recordset")
  if trim(company)<>"" then
  rst2.open "SELECT * FROM "&company&"_STANDARD_ARS_CONTACT WHERE Contact_name='"&contact&"'", cnn1
  if not rst2.eof then
    cPhone = rst2("Telephone")
    cEmail = rst2("Email_Address")
  end if
  rst2.close
  end if
%>
<tr>
  <td>Customer:</td>
  <td><%=customer%></td>
</tr>
<tr>
  <td>Contact:</td>
  <td><%=contact%></td>
</tr>
<tr>
  <td>Phone:</td>
  <td><%=cPhone%></td>
</tr>
<tr>
  <td>Email:</td>
  <td><%=cEmail%></td>
</tr>
<tr>
  <td>Description:</td>
  <td>
  <%=description%> 
  <input type="hidden" name="invoice_date" value="<%=flag%>">
  <input type="hidden" name="jobno" value="<%=job%>">
  </td>
</tr>
<form action="invoiceupdatecmt.asp" name="frm">
<tr valign="top"> 
  <td>Invoice Comment:</td>
  <td><%=icomment%> </td>
</tr>
<tr><td colspan="2"><hr noshade size="1"></td></tr>
<tr>
  <td><b>Total Hours:</b></td>
  <td><b><%=hours%><b></td>
</tr>
<tr>
  <td><b>Total Billable:</b></td>
  <td><b><%=totalbillhours%></b></td>
</tr>
<tr>
  <td colspan="2">
  <br>Breakdown of hours by employee rate:<br>
  <table border=0 cellpadding="3" cellspacing="0" style="border:1px solid #000000;" width="100%">
  <tr align="center" style="font-weight:bold;">
    <td width="20%" style="border-bottom:1px solid #000000;">&nbsp;</td>
    <td width="16%" style="border-bottom:1px solid #000000;">Admin</td>
    <td width="16%" style="border-bottom:1px solid #000000;">Entry</td>
    <td width="16%" style="border-bottom:1px solid #000000;">Junior</td>
    <td width="16%" style="border-bottom:1px solid #000000;">Mid</td>
    <td width="16%" style="border-bottom:1px solid #000000;">Senior</td>
  </tr>
  <tr align="right">
    <td bgcolor="#ffffff" align="right" style="border-bottom:1px solid #cccccc;">Regular:</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorys(5)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorys(1)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorys(2)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorys(3)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorys(4)%>&nbsp;</td>
  </tr>
  <tr align="right">
    <td bgcolor="#ffffff" align="right" style="border-bottom:1px solid #cccccc;">Billable:</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorysbh(5)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorysbh(1)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorysbh(2)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorysbh(3)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=Categorysbh(4)%>&nbsp;</td>
  </tr>
  <tr align="right">
    <td bgcolor="#ffffff" align="right">Overtime:</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;"><%=Categorysot(5)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;"><%=Categorysot(1)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;"><%=Categorysot(2)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;"><%=Categorysot(3)%>&nbsp;</td>
    <td bgcolor="#ffffff" style="border-left:1px solid #cccccc;"><%=Categorysot(4)%>&nbsp;</td>
  </tr>
  </table>	  
  </td>
</tr>
	<%end if
	rst1.close%>
</form>
<%

ReDim Category(5)

Category(1) = "#FFFFFF"
Category(1) = "#00FF00"
Category(2) = "#00CC00"
Category(3) = "#3399CC"
Category(4) = "#FF0000"
Category(5) = "#339999"

'sqlstr = "select invoice_submission.* from invoice_submission where jobno='"& job &"' and invoice_date='" & flag &"'" 
sqlstr = "select invoice_submission.*, employees.category as empcat from invoice_submission join employees on employees.username = invoice_submission.matricola, master_job where jobno=master_job.id and invoice_date='" & flag &"' and jobno='" & job &"' order by employees.category desc, date desc" 

rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<tr>
  <td colspan="2">
  <br>Breakdown of hours by employee:<br>
  <table border="0" cellpadding="3" cellspacing="0" width="550" style="border:1px solid #000000;" class="fineprint">
  <tr bgcolor="#ffffff">
    <td colspan="3">
    <input type="hidden" name="job" value="<%=job%>">
    <input type="hidden" name="flag" value="<%=flag%>">
    <input type="hidden" name="description" value="<%=description%>">
    <input type="hidden" name="comment" value="<%=comment%>">
    <input type="hidden" name="contact2" value="<%=contact%>">
    <input type="hidden" name="customer" value="<%=customer%>">
    &nbsp;
    </td>
    <td colspan="3" align="center"><b>Hours</b></td>
    <td colspan="2" align="center"><b>Expense</b></td>
  </tr>
  <tr bgcolor="#ffffff" valign="bottom"> 
    <td width="10%" style="border-bottom:1px solid #000000;"><b>User</b></td>
    <td width="10%" style="border-bottom:1px solid #000000;"><b>Date</b></td>
    <td width="35%" style="border-bottom:1px solid #000000;"><b>Description</b></td>
    <td width="9%" style="border-left:1px solid #cccccc;border-top:1px solid #cccccc;border-bottom:1px solid #000000;">Regular</td>
    <td width="9%" style="border-top:1px solid #cccccc;border-bottom:1px solid #000000;">Billable</td>
    <td width="9%" style="border-right:1px solid #cccccc;border-top:1px solid #cccccc;border-bottom:1px solid #000000;">Overtime</td>
    <td width="9%" style="border-bottom:1px solid #000000;">Descr.</td>
    <td width="9%" style="border-bottom:1px solid #000000;">Cost</td>
  </tr>
  <%
    do until rst1.eof
    d=rst1("date")
  %>
  <tr valign="top">
    <td style="border-bottom:1px solid #cccccc;"> 
    <input type="hidden" name="key" value="<%=rst1("id")%>">
    <input type="hidden" name="d" value="<%=d%>">
    <input type="hidden" name="job" value="<%=job%>">
    <input type="hidden" name="flag" value="<%=flag%>">
    <input type="hidden" name="description" value="<%=description%>">
    <input type="hidden" name="contact" value="<%=contact%>">
    <input type="hidden" name="customer" value="<%=customer%>">
    <%
    ValueArray = split(rst1("matricola"), "\")
    username=ValueArray(1)
    %>
    <%=username%>
    </td>
    <td style="border-bottom:1px solid #cccccc;"><%=d%></td>
    <td style="border-bottom:1px solid #cccccc;"><%=rst1("description")%></td>
    <td align="right" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=rst1("hours")%>&nbsp;</td>
    <td align="right" style="border-left:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=rst1("hours_bill")%>&nbsp;</td>
    <td align="right" style="border-left:1px solid #cccccc;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"><%=rst1("overt")%>&nbsp;</td>
    <td style="border-bottom:1px solid #cccccc;"><%=rst1("expense")%>&nbsp;</td>
    <td align="right" style="border-bottom:1px solid #cccccc;">  
    <%
    value=rst1("value")
    if value>=0  then
    else
        value=0
    end if
    value=formatcurrency(value, 2)
    %>
    <%=value%>
    </td>
  </tr>
  <%
    rst1.movenext
    loop
    rst1.close
  %>
  </table>
</td>
</tr>
</table>
</body>
</html>