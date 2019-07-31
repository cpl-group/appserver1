<html>
<head>
<%@Language="VBScript"%>
<%
job= Request.Querystring("job")
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
sqlstr = "select invoice_submission.*, employees.category from invoice_submission join employees on employees.username = invoice_submission.matricola, [job log] where jobno=[entry id] and date > [job log].last_invoice and jobno='" & job &"' order by employees.category desc, date desc" 

sqlstr2 = "select sum(hours) as hours, sum(hours_bill) as hours_bill, sum(billable) as billable from invoice_submission join [job log] on invoice_submission.jobno=[job log].[entry id] where invoice_submission.date> [job log].last_invoice and invoice_submission.jobno='"& job &"'"

rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function updateEntry(key, d, job){
	parent.document.frames.detail.location="opstimesheet.asp?id="+key+"&day="+d+"&job="+job
}

function invoice(job){

    var currdate = new Date()
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	document.location="invoiceupdate.asp?job="+job+"&day="+currdate
	
}
function clear(){

	document.forms[0].invoice.value="";
	
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
if not rst1.eof then

%>
  
<table width="100%" border="0">
  <tr > 
      <td width="6%"><font size="1"></font></td>
      <td width="5%"><font face="Arial, Helvetica, sans-serif" size="2" bgcolor="#CCCCCC" color="#000000">User 
        Name</font></td>
      <td width="5%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Date</font></td>
      <td width="56%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Description</font></td>
      
    <td width="4%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Hours</font></td>
      <td width="3%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Bill 
        H</font></td>
      <td width="4%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Over 
        T</font></td>
      <td width="5%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Expense</font></td>
      <td width="4%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Value</font></td>
      <td width="8%"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Total</font></td>
    </tr>
    <%
		do until rst1.eof
		d=rst1("date")
	%>
    
  <tr bgcolor="<%=Category(rst1("Category"))%>"> 
    <form name="form1">
      
      <td width="6%"> 
        <input type="hidden" name="key" value="<%=rst1("id")%>">
		<input type="hidden" name="d" value="<%=d%>">
		<input type="hidden" name="job" value="<%=job%>">
        <input type="button" name="edit" value="edit" onClick="updateEntry(key.value, d.value, job.value)">
      </td>
      <td width="5%"> <font color="#000000" size="2"> 
        <%
	    ValueArray = split(rst1("matricola"), "\")
		username=ValueArray(1)
		%>
        <%=username%> </font></td>
      <td width="5%"> <font color="#000000" size="2"><%=d%> </font></td>
      <td width="56%"> <font color="#000000" size="2"><%=rst1("description")%> 
        </font></td>
      <td width="4%"> <font color="#000000" size="2"><%=rst1("hours")%> </font></td>
      <td width="3%"> <font color="#000000" size="2"><%=rst1("hours_bill")%> </font></td>
      <td width="4%"> <font color="#000000" size="2"><%=rst1("overt")%> </font></td>
      <td width="5%"> <font color="#000000" size="2"><%=rst1("expense")%> </font></td>
      <td width="4%"> <font color="#000000" size="2"> 
        <%
		  value=rst1("value")
		  if value=""  then
  		      value=0
          end if
		  value=formatcurrency(value, 2)
		%>
        <%=value%> </font></td>
      <td width="8%"> <font color="#000000" size="2"> 
        <%
		  billable=Trim(rst1("billable"))
		  if isnull(billable)  then
		      billable=0
	      end if
'		  billable=formatcurrency(billable, 2)
		%>
        <%=billable%> </font></td>
	  </form>
    </tr>
    <%
  		rst1.movenext
		loop
		
    rst1.close
  %>
  </table>
<%

end if
%>
</body>
</html>
