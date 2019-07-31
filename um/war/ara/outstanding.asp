<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,d,c,o,total,currentjob,currentcustomer,gtotal,url,currentmanager,currentcustomerno

d = request("d")
c = request("c")
o = request ("o")


gtotal=0
'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")

' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = '"&c&"_activity_ara_status'",cnn
crdate=rs(0)
rs.close

' specify stored procedure to run based on company
select case c
case "IL" 
cmd.CommandText = "sp_outstanding"
case "GY" 
 cmd.CommandText = "sp_GY_outstanding"
case "NY"
 cmd.CommandText = "sp_NY_outstanding"
case "GE"
 cmd.CommandText = "sp_GE_outstanding"
end select
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("day", adinteger, adParamInput)
cmd.Parameters.Append prm
'response.write crdate
'response.end

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

'return set to recordset rs
cnn.test  d,  rs

select case o
case "J" rs.sort = "job"
case "C" rs.sort = "customer"
case "A" rs.sort = "managername desc"
end select

%>
<html>
<head>
<title></title>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>


<script type="text/javascript">

function openWindow(jobno,company,jid)
{

// Append jobno to http link

if (company=="IL") {
var urlLink     = "http<%if instr(request.servervariables("SERVER_NAME"),"vmgbox")=0 then response.write "s"%>://<%=request.servervariables("SERVER_NAME")%>/um/war/jc/jc.asp?c=" + company + "&ji=" + jobno + "&jid=" + jid
}
else {
var urlLink     = "http<%if instr(request.servervariables("SERVER_NAME"),"vmgbox")=0 then response.write "s"%>://<%=request.servervariables("SERVER_NAME")%>/um/war/jc/jc.asp?c=" + company + "&jg=" + jobno + "&jid=" + jid
}

// Open new window and customize window settings
openwin(urlLink,600,900)

}
function openwin(url,h,w) {

window.open(url,"window","scrollbars=yes,width="+w+",height="+h+",resizeable=no")
}
</script>



<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">

<tr style="font-family:arial; font-size:12; color:black" bgcolor="#0099FF">
    
  <td style="font-size:10" rowspan="2" align="center">
  

    <table width="99%" border="0">
      <tr bgcolor="#6699cc"> 
        <td width="8%"><span class="standardheader">Job ID </span></td>
        <td width="22%" align="center"> <span class="standardheader">Customer</span> 
        </td>
        <td width="20%" align="center"> <span class="standardheader">Job Name</span> 
        </td>
        <td width="10%" align="center"><span class="standardheader">Invoice #</span></td>
        <td width="15%" align="center"><span class="standardheader">Invoice Date</span></td>
        <td width="5%" align="center"><span class="standardheader">Days Out</span></td>
        <td width="10%" align="center"> <span class="standardheader">Amount</span> 
        <td width="20%" align="center"> <span class="standardheader">Account Manager</span> 
        </td>
        <td width="10%" align="center"><span class="standardheader">Request<br>Update</span></td>
      </tr>
<!--      </table>
	 <table width="99%" border="1"> -->
<%
if rs.eof then 
response.Write("record not found")
response.End()
end if

total = 0
currentjob = cstr(rs("job"))
currentcustomer = rs("customer")
currentmanager = rs("managername")
currentcustomerno =  rs("customerno")
while not rs.EOF 
'response.write cstr(rs("job")) &"," & o &"," & currentcustomer
if cstr(rs("job")) = cstr(currentjob) and o = "J" then 
total = total + rs("outstanding_amt")
gtotal=gtotal + rs("outstanding_amt")
else

	if cstr(rs("customerno")) = trim(currentcustomerno) and o = "C" then 
	'if cstr(rs("customer")) = trim(currentcustomer) and o = "C" then 
		total = total + rs("outstanding_amt")
		gtotal=gtotal + rs("outstanding_amt")
	else 
		if trim(rs("managername")) = trim(currentmanager) and o = "A" then 
			total = total + rs("outstanding_amt")
			gtotal=gtotal + rs("outstanding_amt")
		else
		if o="C" then
	%>
	<tr><td colspan="7" align='right'><font color='#666666'><b>Customer#:<%=currentcustomerno%></b></font></td></td><td align="right"><b><%=formatnumber(total,2)%></b></td></tr>
	<% else%>
	<tr><td colspan="7">&nbsp;</td></td><td align="right"><b><%=formatnumber(total,2)%></b></td></tr>
	<%end if%>
	<%

	
	if total=0 then%><!-- </table> --><%end if%>
	<tr><td colspan="8"><hr></td></tr>
		<!-- <table width="99%" border="0"> -->
	<%
	gtotal=gtotal + rs("outstanding_amt")
	total =  rs("outstanding_amt")
	currentjob = cstr(rs("job"))
	currentcustomer = rs("customer")
	currentmanager = rs("managername")
	currentcustomerno = rs("customerno")
	end if
	end if
end if
dim  rss,rssSql,custno,invoicedate,amount,InvNum,j,lastcustno ' to display cust# at bottom
	
	 
set rss = server.createobject("ADODB.Recordset")
InvNum=rs("invoice#")
j=rs("job")
rssSql="select " &c&"_BILLED_BLI_INVOICE.invoice_date as Invoicedate,customer,Invoice_amount from " &c&"_BILLED_BLI_INVOICE where invoice='"&rs("invoice#")&"'"
rss.open rssSql, cnn,1,1
if not rss.EOF then 
custno= rss("customer")
invoicedate =rss("Invoicedate")
amount=rss("Invoice_amount")
rss.close
end if
%>
      <tr>
        <%if rs("job") = "" then %>
        <td width="8%" height="20" bgcolor="#FF0000"><font size="2"><%=rs("job")%></font></td>
        <% else %>
        <td width="8%" height="20" > <font size="2"><a href="javascript:openWindow('<%=rs("job")%>','<%=c%>','<%response.write split(rs("job"),"-")(1)%>')"><%=rs("job")%></a></font></td>
        <% end if%>
        <td width="22%" height="20"> <font size="2"><%=rs("customer")%> </font></td>
        <td width="20%" height="20"> <font size="2"><%=rs("jobname")%></font><br><font size="1"><%=rs("jobaddress")%></font></td>
		<% url= "/um/war/ara/bliview.asp?jid=" & rs("job") & "&c=" & c & "&invoiceid=" & rs("invoice#")%> 
        <td width="10%" height="20" align="right"> <font size="2"><a href="javascript:openwin('<%=url%>',600,900)"><%=rs("invoice#")%></a>
		<% url= "/um/war/notes/notes.asp?c="&c&"&iid=" & rs("invoice#")&"&manager=" & rs("muserid") &"&customer="&rs("customer")&"&custno="&custno&"&invoicedate="&invoicedate&"&amount="&amount&"&InvNum="&InvNum&"&j="&j%>
          <a href="javascript:openwin('<%=url%>',400,800)"><img src="/images/notes.gif" border="0"></a>
          <%if rs("ticket") = 1 then%>*<%else%>&nbsp;<% end if %></font></td>
        <td width="15%" height="20" align="right"> <font size="2"><%=rs("invoice_date")%></font></td>
        <td width="5%" height="20" align="right" bgcolor="#C2C0C1"> <font size="2"><%=rs("past_due_days")%></font></td>
        <td width="10%" height="20" align="right" bgcolor="#CCCCCC"><font size="2"><%=formatnumber(rs("outstanding_amt"),2)%></font></td>
        <td width="20%" height="20" align="right" bgcolor="#CCCCCC"><font size="2"><%if trim(currentmanager) <> "" then%><%=rs("managername")%><%else%>NA<%end if%></font></td>
        <td width="20%" height="20" align="right" bgcolor="#CCCCCC"><input type="button" value="Email" onClick="window.open('emailprocess.asp?jobid=<%if instr(rs("job"),"-")>0 then response.write split(rs("job"),"-")(1)%>&invoiceid=<%=rs("invoice#")%>&comapny=<%=c%>','processjobemail', 'width=400,height=100, scrollbars=no');"></td>
      </tr>
      <%
	 
	  lastcustno=rs("customerno")
 rs.movenext


wend
%>
      <tr>
       <% if o="C" then%>
	<td colspan="7" align='right'><font color='#666666'><b>Customer#:<%=lastcustno%></b></font></td></td>
	  <%else%>
		<td colspan="6" align="right">&nbsp;</td>
		<% end if%>
		<!--<td>Customer#:</td>-->
        <td align="right" ><b><%=formatnumber(total,2)%></b></td>
      </tr>
    </table>
    <hr> 
    <p>&nbsp;</p>
    <%response.Write("*  This asterix signifies that there is a trouble ticket attached to this invoice.")  %>
	<table width="99%" border="0" dwcopytype="CopyTableRow">
      <tr> 
        <td height="21" width="46%"> 
          <div align="left"><font size="2">Updated as of <%=crdate%></font></div>
        </td>
        <td height="21" width="46%">
          <div align="right"><b>Outstanding Grand Total over <%=d%> days </b></div>
        </td>
        <td height="21" width="8%"> 
          <div align="right"><b><%=formatnumber(gtotal,2)%></b></div>
        </td>
      </tr>
    </table>
	
	
      <tr> 
        <td>
          <div align="right"></div>
  </td>
      </tr> 
	  <%
	  set cnn = nothing %>
</body>
</html>
