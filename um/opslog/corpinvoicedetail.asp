<html> <head> 
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
job= Request("job")
flag= Request("day")
description=Request("description")
comment=Request("comment")
customer=Request("customer")
'contact=Request("contact")
ReDim Category(5)

Category(1) = "#FFFFFF"
Category(1) = "#66ff66"
Category(2) = "#339999"
Category(3) = "#ff9900"
Category(4) = "#cc0000"
Category(5) = "#666699"

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")
sqlstr = "select invoice_submission.* from invoice_submission where jobno='"& job &"' and invoice_date='" & flag &"'" 
sqlstr = "select invoice_submission.*, employees.category as empcat from invoice_submission join employees on employees.username = invoice_submission.matricola, [master_job] where jobno=master_job.[id] and invoice_date='" & flag &"' and jobno='" & job &"' order by employees.category desc, date desc" 
rst1.Open sqlstr, cnn1, 0, 1, 1
if not rst1.eof then
	if rst1("flag") then dbflag = true else dbflag = false
else
	dbflag = false
end if

%>
<script>
function updateEntry(key, d, job, flag, description,customer,contact){
    if(parent.document.frm.showhist.value==1) {
	  parent.document.frm.showhist.value=0
	  parent.document.frames.list.location=parent.document.frames.detail.location
	}
    parent.document.frames.detail.location="corpinvoicemodify.asp?id="+key+"&day="+d+"&job="+job+"&flag="+flag+"&description="+description+"&customer="+customer+"&contact="+contact
}
function addtime(job,invday,des, comment,customer,contact){
	var temp = "corptimeadd.asp?job="+ job + "&invday=" +invday +"&des="+des + "&comment="+comment+"&customer="+customer+"&contact="+contact;
	if(parent.document.frm.showhist.value==1) {
	  parent.document.frm.showhist.value=0
	  parent.document.frames.list.location=parent.document.frames.detail.location
	}
	parent.document.frames.detail.location= temp
}
function invoice(job){
    var currdate = new Date()
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	document.location="corpinvoiceupdate.asp?job="+job+"&day="+currdate
}

</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc">
<tr valign="bottom" bgcolor="#dddddd">
  <td>&nbsp;</td>
  <td>
  <input type="hidden" name="job" value="<%=job%>">
  <input type="hidden" name="flag" value="<%=flag%>">
  <input type="hidden" name="description" value="<%=description%>">
  <input type="hidden" name="comment" value="<%=comment%>">
  <input type="hidden" name="contact" value="<%=contact%>">
  <input type="hidden" name="customer" value="<%=customer%>">
  <% if not dbflag then %>
  <input type="submit" name="addtime" value="Add Time" onClick="addtime(job.value, flag.value, description.value,comment.value,customer.value, contact.value)">
  <%end if %>
  </td>
  <td>User Name</td>
  <td>Date</td>
  <td>Description</td>
  <td width="6%">Hours</td>
  <td width="6%">Billable Hours</td>
  <td width="6%">Overtime</td>
  <td width="6%">Expense</td>
  <td width="6%">Value</td>
</tr>
<%
  do until rst1.eof
  d=rst1("date")
if rst1("category") = 0 then
%>
<tr valign="top" bgcolor="<%=Category(rst1("empcat"))%>"> 
  <% else %>
<tr valign="top" bgcolor="<%=Category(rst1("category"))%>"> 
  <% end if %>
  <form name="form1">
  <td width="2%" bgcolor="<%=Category(rst1("category"))%>">&nbsp;
  <input type="hidden" name="key" value="<%=rst1("id")%>">
  <input type="hidden" name="d" value="<%=d%>">
  <input type="hidden" name="job" value="<%=job%>">
  <input type="hidden" name="flag" value="<%=flag%>">
  <input type="hidden" name="description" value="<%=description%>">
  <input type="hidden" name="contact" value="<%=contact%>">
  <input type="hidden" name="customer" value="<%=customer%>">
  </td>
  <td bgcolor="#ffffff">
  <a name="<%=rst1("Category")%>"></a> 
  <% if not dbflag then %>
  <input type="button" name="edit" value="edit" onClick="updateEntry(key.value, d.value, job.value, flag.value, description.value,customer.value,contact.value)">
  <%end if %>
  </td>
  <td bgcolor="#ffffff">  
  <%
  ValueArray = split(rst1("matricola"), "\")
  username=ValueArray(1)
  %>
  <%=username%>
  </td>
  <td bgcolor="#ffffff"><%=d%></td>
  <td bgcolor="#ffffff"><%=rst1("description")%></td>
  <td bgcolor="#ffffee"><%=rst1("hours")%></td>
  <td bgcolor="#f0f0e0"><%=rst1("hours_bill")%></td>
  <td bgcolor="#e3e3d3"><%=rst1("overt")%></td>
  <td bgcolor="#ffffff"><%=rst1("expense")%></td>
  <td bgcolor="#ffffff">  
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
  </form>
</tr>
<%
rst1.movenext
loop
rst1.close
%>
</table>
<br><br>
<br><br>
<br><br>
<br><br>
<br><br>
<br><br>
<br><br>
</body>
</html>
