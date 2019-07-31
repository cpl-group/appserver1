<html>
<head>
<%@Language="VBScript"%>
<%
job= Request.Querystring("job")
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
dim edit,hgt
if request("edit")="no" then
  edit=false
  hgt=" height=""25"""
else
  edit=true
  hgt=""
end if

cnn1.Open getConnect(0,0,"intranet")
sqlstr = "select times.*, employees.category from times join employees on employees.username = times.matricola, master_job where jobno=master_job.id and date > master_job.last_invoice and jobno='" & job &"' order by employees.category desc, date desc" 

' Not referenced below -> sqlstr2 = "select sum(hours) as hours, sum(hours_bill) as hours_bill, sum(billable) as billable from times join master_job on times.jobno=master_job.id where times.date> master_job.last_invoice and times.jobno='"& job &"'"

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
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
if not rst1.eof then 

%>
  
<table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc" width="100%">
<tr valign="bottom" bgcolor="#dddddd">
  <td width="2%"></td>
  <%if edit then%><td>&nbsp;</td><%end if%>
  <td>User Name</td>
  <td>Date</td>
  <td>Description</td>
  <td width="6%">Hours</td>
  <td width="6%">Billable Hours</td>
  <td width="6%">Overtime</td>
  <td width="6%">Expense</td>
  <td width="6%">Value</td>
  <td>Total</td>
</tr>
<%
do until rst1.eof
d=rst1("date")
%>
    
<tr valign="top" bgcolor="#ffffff"> 
  <form name="form1">
  <input type="hidden" name="key" value="<%=rst1("id")%>">
  <input type="hidden" name="d" value="<%=d%>">
  <input type="hidden" name="job" value="<%=job%>">
  <td width="2%" bgcolor="<%=Category(rst1("Category"))%>">&nbsp;</td>
  <%if edit then%><td><input type="button" name="edit" value="edit" onClick="updateEntry(key.value, d.value, job.value)"></td><%end if%>
  <td> 
  <%
  ValueArray = split(rst1("matricola"), "\")
  username=ValueArray(1)
  %>
  <%=username%> </td>
  <td><%=d%> </td>
  <td><%=rst1("description")%></td>
  <td align="right"><%=rst1("hours")%> </td>
  <td align="right"><%=rst1("hours_bill")%></td>
  <td align="right"><%=rst1("overt")%> </td>
  <td align="right"><%=rst1("expense")%> </td>
  <td> 
        <%
		  value=rst1("value")
		  if value=""  then
  		      value=0
          end if
		  value=formatcurrency(value, 2)
		%>
        <%=value%> </td>
      <td> 
        <%
		  billable=Trim(rst1("billable"))
		  if isnull(billable)  then
		      billable=0
	      end if
'		  billable=formatcurrency(billable, 2)
		%>
        <%=billable%> </td>
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