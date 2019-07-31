<html> <head> 
<%@Language="VBScript"%>
<%
job= Request("job")
flag= Request("day")
description=Request("description")
comment=Request("comment")
customer=Request("customer")
contact=Request("contact")
ReDim Category(5)

Category(1) = "#FFFFFF"
Category(1) = "#00FF00"
Category(2) = "#00CC00"
Category(3) = "#3399CC"
Category(4) = "#FF0000"
Category(5) = "#339999"

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
sqlstr = "select invoice_submission.* from invoice_submission where jobno='"& job &"' and invoice_date='" & flag &"'" 
sqlstr = "select invoice_submission.*, employees.category as empcat from invoice_submission join employees on employees.username = invoice_submission.matricola, [job log] where jobno=[entry id] and invoice_date='" & flag &"' and jobno='" & job &"' order by employees.category desc, date desc" 


rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<script>
function updateEntry(key, d, job, flag, description,customer,contact){
	parent.document.frames.detail.location="corpinvoicemodify.asp?id="+key+"&day="+d+"&job="+job+"&flag="+flag+"&description="+description+"&customer="+customer+"&contact="+contact
}
function addtime(job,invday,des, comment,customer,contact){
	var temp = "corptimeadd.asp?job="+ job + "&invday=" +invday +"&des="+des + "&comment="+comment+"&customer="+customer+"&contact="+contact;
	parent.document.frames.detail.location= temp
}
function invoice(job){
    var currdate = new Date()
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	document.location="corpinvoiceupdate.asp?job="+job+"&day="+currdate
}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" >
  <tr > 
    <td width="6%"><font size="1"> 
      <input type="hidden" name="job" value="<%=job%>">
      <input type="hidden" name="flag" value="<%=flag%>">
      <input type="hidden" name="description" value="<%=description%>">
      <input type="hidden" name="comment" value="<%=comment%>">
      <input type="hidden" name="contact" value="<%=contact%>">
      <input type="hidden" name="customer" value="<%=customer%>">
	  <% if not rst1("flag") then %>
      <input type="submit" name="addtime" value="Add Time" onClick="addtime(job.value, flag.value, description.value,comment.value,customer.value, contact.value)">
	  <%end if %>
      </font></td>
    <td width="5%"><font face="Arial, Helvetica, sans-serif" size="2" bgcolor="#CCCCCC" color="#000000">User 
      Name</font></td>
    <td width="5%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Date</font></td>
    <td width="56%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Description</font></td>
    <td width="4%" bgcolor="#00CCFF"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Hours</font></b></font></div>
    </td>
    <td width="3%" bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Bill 
        H</font></b></font></div>
    </td>
    <td width="4%" bgcolor="#0033FF"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Over 
        T</font></b></font></div>
    </td>
    <td width="5%" bgcolor="#0066CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Expense</font></b></font></div>
    </td>
    <td width="4%" bgcolor="#3300CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2">Expense 
        Cost</font></b></font></div>
    </td>
  </tr>
  <%
		do until rst1.eof
		d=rst1("date")
	if rst1("category") = 0 then
	%>
  <tr bgcolor="<%=Category(rst1("empcat"))%>"  > 
    <% else %>
  <tr  > 
    <% end if %>
    <form name="form1">
      <td width="6%" bgcolor="<%=Category(rst1("category"))%>"> 
        <input type="hidden" name="key" value="<%=rst1("id")%>">
        <input type="hidden" name="d" value="<%=d%>">
        <input type="hidden" name="job" value="<%=job%>">
        <input type="hidden" name="flag" value="<%=flag%>">
        <input type="hidden" name="description" value="<%=description%>">
        <input type="hidden" name="contact" value="<%=contact%>">
        <input type="hidden" name="customer" value="<%=customer%>">
        <a name="<%=rst1("Category")%>"></a> <font size="1">
        <% if not rst1("flag") then %>
        </font>
<input type="button" name="edit" value="edit" onClick="updateEntry(key.value, d.value, job.value, flag.value, description.value,customer.value,contact.value)">
        <font size="1">
        <%end if %>
        </font> </td>
      <td width="5%" bgcolor="<%=Category(rst1("category"))%>"> <font color="#000000" size="2"> 
        <%
	  ValueArray = split(rst1("matricola"), "\")
	  username=ValueArray(1)
	  %>
        <%=username%> </font></td>
      <td width="5%" bgcolor="<%=Category(rst1("category"))%>"> <font color="#000000" size="2"><%=d%> 
        </font></td>
      <td width="56%" bgcolor="<%=Category(rst1("category"))%>"> <font color="#000000" size="2"><%=rst1("description")%> 
        </font></td>
      <td width="4%" bgcolor="#00CCFF" height="34" bordercolor="#999999" valign="middle"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("hours")%> 
          </font></div>
      </td>
      <td width="4%" bgcolor="#3399CC" height="34" bordercolor="#999999" valign="middle"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("hours_bill")%> 
          </font></div>
      </td>
      <td width="3%" bgcolor="#0033FF" height="34" bordercolor="#999999" valign="middle"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("overt")%> 
          </font></div>
      </td>
      <td width="4%" bgcolor="#0066CC" height="34" bordercolor="#999999" valign="middle"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("expense")%> 
          </font></div>
      </td>
      <td width="4%" bgcolor="#3300CC"> <font color="#FFFFFF" size="2"> 
        <%
	  value=rst1("value")
	  if value>=0  then
	  else
	  	  value=0
	  end if
	  value=formatcurrency(value, 2)
	  %>
        <%=value%> </font></td>
    </form>
  </tr>
  <%
  	   	  rst1.movenext
		loop
		
    rst1.close
  %>
</table>
</body>
</html>
