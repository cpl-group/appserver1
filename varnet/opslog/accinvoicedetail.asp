<html>
<head>
<%@Language="VBScript"%>
<%
job= Request("job")
flag= Request("day")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
sqlstr = "select invoice_submission.* from invoice_submission where jobno='"& job &"' and invoice_date='" & flag &"'" 

sqlstr2 = "select sum(hours) as hours, sum(hours_bill) as hours_bill, sum(billable) as billable from invoice_submission where invoice_date='"& flag &"'"
rst1.Open sqlstr, cnn1, 0, 1, 1

%>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function updateEntry(key, d, job, flag){
	document.frames.detail.location="corpinvoicemodify.asp?id="+key+"&day="+d+"&job="+job+"&flag="+flag
}

function invoice(job){
    var currdate = new Date()
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	document.location="corpinvoiceupdate.asp?job="+job+"&day="+currdate
}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
if not rst1.eof then

%>
<table width="100%" border="0" height="6%" cellspacing="0">
  <tr>
    <td width="100%" bgcolor="#3399CC">
      <div align="center"><i><font face="Georgia, Times New Roman, Times, serif" size="+1">Job No. <%=job%> 
        Invoice date: <%=flag%></font></i></div>
    </td></tr></table>
<div align="right">
<input type="button" name="Submit" value="Back" onclick="javascript:history.back()">
</div>
<table width="100%" border="0" >
  <tr > 
    <td width="6%"><font size="1"></font></td>
    <td width="5%"><font face="Arial, Helvetica, sans-serif" size="2" bgcolor="#CCCCCC" color="#000000">User Name</font></td>
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
  <tr> 
	<form name="form1">
      
    <td width="6%"> 
	  <input type="hidden" name="key" value="<%=rst1("id")%>">
	  <input type="hidden" name="d" value="<%=d%>">
	  <input type="hidden" name="job" value="<%=job%>">
	  <input type="hidden" name="flag" value="<%=flag%>">
      <input type="button" name="edit" value="edit" onClick="updateEntry(key.value, d.value, job.value, flag.value)">
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
	  if value>=0  then
	  else
	  	  value=0
	  end if
	  value=formatcurrency(value, 2)
	  %>
      <%=value%> </font></td>
    <td width="8%"> <font color="#000000" size="2"> 
      <%
	  billable=rst1("billable")
	  if billable>=0  then
	  else
	      billable=0
      end if
	  billable=formatcurrency(billable, 2)
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
    rst1.Open sqlstr2, cnn1, 0, 1, 1
  %>
  
  <%
    if not rst1.eof then
        hourlyrate=0.00
	    hours=Trim(rst1("hours"))
		totalbillhours=Trim(rst1("hours_bill"))
		billable=Trim(rst1("billable"))
		if billable >=0 then
		else
			billable=0
		end if
		billable=formatnumber(billable, 2)
		if(totalbillhours>0) then
			hourlyrate=formatcurrency(billable/totalbillhours, 2)
		end if
		billable=formatcurrency(billable, 2)
	end if
  %>
  <table width="50%" border="0">
 
  <tr>
    <td>Total Hours</td>
    <td bgcolor="#CCCCCC"><div align="right"><%=hours%> 
      </div>
    </td>
  </tr>
  <tr>
    <td>Total Billable</td>
    <td bgcolor="#CCCCCC"><div align="right"> <%=totalbillhours%> 
      </div>
    </td>
  </tr>
  <tr>
    <td>Total</td>
    <td bgcolor="#CCCCCC">
      <div align="right"><%=billable%></div>
    </td>
  </tr>
  <tr>
    <td>Hourly Rate</td>
    <td bgcolor="#CCCCCC"><div align="right"><%=hourlyrate%></div></td>
  </tr>
</table>
<%
end if
%>
<IFRAME name="detail" width="100%" height="150" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>

</body>
</html>
