<html>
<head>
<%@Language="VBScript"%>
<%
tjob= Request("tjob")
job= Request("job")
flag= Request("day")
description= Request("d")
comment= Request("day")
customer=Request("customer")
contact=Request("contact")
'response.write Request("job")
'response.write Request("day")
'response.write Request("d")
'response.end


ReDim Categorys(5)
ReDim Categorysbh(5)
ReDim Categorysot(5)

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

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
</head>

<body bgcolor="#FFFFFF" text="#000000">


<table width="100%" border="0" cellspacing="0">
  <tr >
    <td width="57%" height="2" > 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><b>Invoice for 
        Job No. <%=tjob%> Invoice date: <%=flag%></b></font></div>
    </td>
</tr></table>

 
<% 	
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
if not rst1.eof then
%>
<table width="100%">
<tr>
    <td width="35%" height="2" > <font face="Arial, Helvetica, sans-serif"><b><i>Customer: <%=customer%>, 
    <%=contact%></i></b></font></td>
	<tr>
	<td><font face="Arial, Helvetica, sans-serif"><i><b><%=description%> 
    <input type="hidden" name="invoice_date" value="<%=flag%>">
    <input type="hidden" name="jobno" value="<%=job%>">
    </b></i></font></td>
	</tr>

  
</table>
<form action="invoiceupdatecmt.asp" name="frm">
  <table width="100%" >
    <tr> 
	  <td width="99%" height="24"><font face="Arial, Helvetica, sans-serif"><b>Invoice 
        Comment </b>:</font> <font face="Arial, Helvetica, sans-serif"><%=rst1("invoice_comment")%></font> 
    </tr>
	<tr>
	  <td width="99%" height="21"> <font face="Arial, Helvetica, sans-serif"><b>Total 
        Hours:</b> <%=hours%><b> Total Billable:</b> <%=totalbillhours%> </font></td>
	</tr>
	<tr>
	  <td><font face="Arial, Helvetica, sans-serif"><b>Admin:</b>[ <%=Categorys(5)%> 
        ] [ <%=Categorysbh(5)%> ] [ <%=Categorysot(5)%> ]<b> Entry: </b>[ <%=Categorys(1)%> 
        ] [ <%=Categorysbh(1)%> ] [ <%=Categorysot(1)%> ] <b>Junior</b></font><b> 
        :</b> [ <%=Categorys(2)%> ] [ <%=Categorysbh(2)%> ] [ <%=Categorysot(2)%> 
        ]<b> </b> <font face="Arial, Helvetica, sans-serif"><b>Mid: </b>[ <%=Categorys(3)%> 
        ] [ <%=Categorysbh(3)%> ] [ <%=Categorysot(3)%> ]<b> </b> <b>Senior:</b>[ 
        <%=Categorys(4)%> ] [ <%=Categorysbh(4)%> ] [ <%=Categorysot(4)%> ] </font></td>
	</tr>
	<%end if
	rst1.close%>
  </table>
</form>
<%

ReDim Category(5)

Category(1) = "#FFFFFF"
Category(1) = "#00FF00"
Category(2) = "#00CC00"
Category(3) = "#3399CC"
Category(4) = "#FF0000"
Category(5) = "#339999"

sqlstr = "select invoice_submission.* from invoice_submission where jobno='"& job &"' and invoice_date='" & flag &"'" 
sqlstr = "select invoice_submission.*, employees.category as empcat from invoice_submission join employees on employees.username = invoice_submission.matricola, [job log] where jobno=[entry id] and invoice_date='" & flag &"' and jobno='" & job &"' order by employees.category desc, date desc" 

rst1.Open sqlstr, cnn1, 0, 1, 1
%>

<table width="100%" border="0" >
  <tr > 
    <td width="7%" height="47"><font size="1"> 
      <input type="hidden" name="job" value="<%=job%>">
      <input type="hidden" name="flag" value="<%=flag%>">
      <input type="hidden" name="description" value="<%=description%>">
      <input type="hidden" name="comment" value="<%=comment%>">
      <input type="hidden" name="contact2" value="<%=contact%>">
      <input type="hidden" name="customer" value="<%=customer%>">
      </font></td>
    <td width="4%" height="47"><font face="Arial, Helvetica, sans-serif" size="2" ><b>User 
      Name</b></font></td>
    <td width="3%" height="47"><font face="Arial, Helvetica, sans-serif" size="2" ><b>Date</b></font></td>
    <td width="51%" height="47"><font face="Arial, Helvetica, sans-serif" size="2" ><b>Description</b></font></td>
    <td width="5%" height="47" > 
      <div align="center"><font ><b><font face="Arial, Helvetica, sans-serif" size="2">Hours</font></b></font></div>
    </td>
    <td width="5%" height="47" > 
      <div align="center"><font ><b><font face="Arial, Helvetica, sans-serif" size="2">Bill 
        H</font></b></font></div>
    </td>
    <td width="5%" height="47"> 
      <div align="center"><font ><b><font face="Arial, Helvetica, sans-serif" size="2">Over 
        T</font></b></font></div>
    </td>
    <td width="6%" height="47"> 
      <div align="center"><font><b><font face="Arial, Helvetica, sans-serif" size="2">Expense</font></b></font></div>
    </td>
    <td width="14%" height="47"> 
      <div align="center"><font ><b><font face="Arial, Helvetica, sans-serif" size="2">Expense 
        Cost</font></b></font></div>
    </td>
  </tr>
  <%
		do until rst1.eof
		d=rst1("date")
	if rst1("category") = 0 then
	%>
  <tr  > 
    <% else %>
	
  <tr  > <form name="form2" method="post" action="">
    <% end if %>
    

      <td width="7%" height="46" > 
        <input type="hidden" name="key" value="<%=rst1("id")%>">
        <input type="hidden" name="d" value="<%=d%>">
        <input type="hidden" name="job" value="<%=job%>">
        <input type="hidden" name="flag" value="<%=flag%>">
        <input type="hidden" name="description" value="<%=description%>">
        <input type="hidden" name="contact" value="<%=contact%>">
        <input type="hidden" name="customer" value="<%=customer%>">
        <a name="<%=rst1("Category")%>"></a> <font size="1">
        
        </font> </td>
      <td width="4%" height="46" > <font  size="2"> 
        <%
	  ValueArray = split(rst1("matricola"), "\")
	  username=ValueArray(1)
	  %>
        <%=username%> </font></td>
      <td width="3%" height="46" > <font size="2"><%=d%> </font></td>
      <td width="51%" height="46" > <font size="2"><%=rst1("description")%> </font></td>
      <td width="5%"  height="46" bordercolor="#999999" valign="middle"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" ><%=rst1("hours")%> 
          </font></div>
      </td>
      <td width="5%" height="46" bordercolor="#999999" valign="middle"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" ><%=rst1("hours_bill")%> 
          </font></div>
      </td>
      <td width="5%"  height="46" bordercolor="#999999" valign="middle"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("overt")%> 
          </font></div>
      </td>
      <td width="6%"  height="46" bordercolor="#999999" valign="middle"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("expense")%> 
          </font></div>
      </td>
      <td width="14%" height="46" > <font color="#FFFFFF" size="2"> 
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

