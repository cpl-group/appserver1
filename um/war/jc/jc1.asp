<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,j,c,avg,wip,tcost,OPENPO,link,amt_paid,aTable,sql, prefix,hours, jid
jid = Request("jid")

'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")
'open connection
cnn.open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT company FROM MASTER_JOB WHERE id="&jid, cnn
if not rs.eof then c = rs("company") else c = request("c")
rs.close
if c="GY" then 
	'j = request ("jg")
	aTable = "GY_MASTER_JOB"
	prefix = "jg"
else
	'j = request ("ji")
	aTable = "GE_MASTER_JOB"
	prefix = "ji"
end if
j = jid

sql = "select job,customer_name,description,address_1,address_2 from " & atable & " where job like '%" & j & "%' or description like '%" & j & "%' or Address_1 like '%" & j & "%' or Address_2 like '%" & j & "%'  order by job"
'response.write sql
'response.end
rs.open sql,cnn

if rs.EOF then 
	rs.close
	response.write "NO RECORDS FOUND"
	set cnn = nothing 
	response.end
end if

if rs.recordcount > 1 then 
	%>
	<title>Job Cost Report</title>
	<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
	</head>
	
	<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">
	<table border=0 celpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
		<tr bgcolor="#dddddd" style="font-weight:bold">
			<td>Job Number</td>
			<td>Description</td>
			<td colspan="2">Address</td>
		</tr>
		<% while not rs.EOF %>
			<tr bgcolor="#ffffff" valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" onClick="javascript:document.location='jc.asp?<%=prefix%>=<%=rs("job")%>&c=<%=c%>'"> 
			 <td width="11%"><%=rs("job")%></td>
			 <td width="34%"><%=rs("description")%></td>
			 <td width="32%"><%=rs("address_1")%></td>
			 <td width="23%"><%=rs("address_2")%></td>
			</tr>
			<% 
			rs.movenext
		wend
		%>
	</table>  <%
	response.write "</body>"
	response.write "</html>"	%>
	<%
	rs.close
else
	j = rs("job")
	rs.close
	
	rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_activity_ara_status'",cnn
	crdate=rs(0)
	rs.close
	
	if c="GY" then
		'get sum of hours from time sheet
		rs.open "SELECT isnull(sum(hours)+sum(overt),0) FROM times WHERE jobno = '"&split(j,"-")(1)&"'",cnn
		hours=rs(0)
		rs.close
		sql="select isnull(sum(a.amount)-sum(a.amount_invoiced),0) from gy_master_po_item a,gy_master_po b where b.closed=0 and a.job='"&j&"' and a.commitment=b.commitment"
	else
		rs.open "SELECT isnull(sum(hours)+sum(overt),0) FROM times WHERE jobno = '"&split(j,"-")(1)&"'",cnn
		hours=rs(0)
		rs.close
		sql="select isnull(sum(a.amount)-sum(a.amount_invoiced),0) from il_master_po_item a,il_master_po b where b.closed=0 and a.job='"&j&"' and a.purchase_order=b.purchase_order"
	end if
	
	rs.open sql,cnn
	openpo=rs(0)
	rs.close
	
	rs.open "SELECT isnull(SUM(AMOUNT)+sum(adjustment),0) FROM " + c + "_ACTIVITY_ARA_STATUS WHERE status='Paid' and job = '" + j + "'",cnn
	amt_paid=rs(0)
	rs.close
	
	dim totalInvoices
	'rs.open "SELECT isnull(sum(invoice_amt),0) as bignumber, jobno from invoice_submission where jobno='"&j&"' group by jobno"
	rs.open "select sum(s.amt) as bignumber from (select invoice_amt amt from invoice_submission where jobno = '"&split(j,"-")(1)&"' and flag=0 group by jobno, invoice_date, invoice_amt) s"
	if rs.eof then totalInvoices = 0 else totalInvoices = rs("bignumber")
	rs.close
	' specify stored procedure 
	
	cmd.CommandText = "sp_job_cost"
	cmd.CommandType = adCmdStoredProc
	
	Set prm = cmd.CreateParameter("c", adchar, adParamInput,2)
	cmd.Parameters.Append prm
	
	Set prm = cmd.CreateParameter("j", advarchar, adParamInput,9)
	cmd.Parameters.Append prm
	
	' assign internal name to stored procedure
	cmd.Name = "test"
	Set cmd.ActiveConnection = cnn
	
	'return set to recordset rs
	cnn.test c,j, rs
	
	if rs.eof then
		response.write "no record found"
		set cnn = nothing 
		response.end
	else
	if rs("jtd_labor_units") = 0 then
	response.write "no hours/times posted within the account system"
	else
	avg = (rs("jtd_labor_cost") + rs("jtd_overhead_cost") + rs("jtd_other_cost") )/ rs("jtd_labor_units")
	wip =  (rs("Revised_Contract_Amount") * rs("percent_complete")/100)-rs("jtd_work_billed")
	'tcost = ( cdbl(hours) * cdbl(avg)) + cdbl(rs("jtd_subcontract_cost"))+ cdbl(rs("jtd_material_cost"))
	end if
	%>
<html>
<head>
<script language="JavaScript1.2">
function po(c,j) {
if (c=="IL") {
var urlLink     = "/um/war/apa/material_cost.asp?c=" + c + "&ji=" + j
}
else {
var urlLink     = "/um/war/apa/material_cost.asp?c=" + c + "&jg=" + j
}
  openwin(urlLink,900,400)
}

function open_po(c,j) {
if (c=="IL") {
var urlLink     = "/um/war/po/open_po.asp?c=" + c + "&ji=" + j
}
else {
var urlLink     = "/um/war/po/open_po.asp?c=" + c + "&jg=" + j
}
  openwin(urlLink,800,400)
}



function time(c,j) {
if (c=="IL") {
var urlLink     = "/um/war/ts/ts.asp?c=" + c + "&ji=" + j
}
else {
var urlLink     = "/um/war/ts/ts.asp?c=" + c + "&jg=" + j
}	openwin(urlLink,800,400)
}

function invoice(c,j) {

if (c=="IL") {
var urlLink     = "/um/war/ara/invoice.asp?c=" + c + "&ji=" + j
}
else {
var urlLink     = "/um/war/ara/invoice.asp?c=" + c + "&jg=" + j
}
  openwin(urlLink,800,400)
}


function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scrollbars=yes, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

  //visual feedback functions for img buttons
  function buttonOver(obj,clr){
    if (arguments.length == 1) { clr = "#336699"; }
    obj.style.border = "1px solid " + clr;
  }
  
  function buttonDn(obj,clr){
    if (arguments.length == 1) { clr = "#000000"; }
    obj.style.border = "1px solid " + clr;
  }
  
  function buttonOut(obj,clr){
    if (arguments.length == 1) { clr = "#eeeeee"; }
    obj.style.border = "1px solid " + clr;
  }

</script>

<title>Job Cost Report</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#eeeeee">
	
  
<table border=0 cellpadding="3" cellspacing="0" id="genjobtable" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #cccccc;border-right:1px solid #cccccc;">
<%
dim sqlstr, rst3, tcolor,Desc,company,job,jtype,cstatus, rst2
set rst3 = server.createobject("ADODB.Recordset")
sqlstr = "SELECT description, type, company, job, status FROM MASTER_JOB WHERE id='"&jid&"'"
rst3.open sqlstr, cnn

if not rst3.EOF then
  Desc = rst3("description")
  company = rst3("company")
  job = rst3("job")
  jtype=left(rst3("type"),6)
  cStatus = rst3("status")
  Select Case cStatus
    case "In progress"
      tcolor = "#66ff66"
    case "Unstarted"
      tcolor = "#ffcc00"
    case "Closed"
      tcolor = "#cc0033"
  end select 
end if
rst3.close
%>
<tr> 
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee" nowrap>
  <table border="0" cellspacing="0" cellpadding="3" width="100%">
  <tr>
    <td nowrap>
    <a href="/gEnergy2_Intranet/opsmanager/joblog/<%="viewjob.asp?jid=" & jid %>">General Info</a> &nbsp;|&nbsp; <a href="/gEnergy2_Intranet/opsmanager/joblog/jobtime.asp?jid=<%=jid%>">Job Time</a> &nbsp;|&nbsp; <a href="/gEnergy2_Intranet/opsmanager/joblog/<%="jobfolder.asp?jid="&jid%>">Job Folder</a>&nbsp;|&nbsp; <a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requestion Forms</a>&nbsp;|&nbsp; <b>Job Cost</b>&nbsp;|&nbsp; <a href="/gEnergy2_Intranet/opsmanager/joblog/viewchange.asp?jid=<%=jid%>">Change Orders</a>
    </td>
  </tr>
  </table>
  </td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
<tr> 
  <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <table border=0 cellspacing="0" cellpadding="3" width="100%">
  <tr>
    <td width="130">Job:</td>
    <td><%=Desc%>&nbsp; (<%=company%>: &nbsp;<%=job%>) &nbsp;<%=jtype%></td>
  </tr>
  <tr>
    <td>Status:</td>
    <td>
    <table border=0 cellpadding="0" cellspacing="0">
    <tr>
      <td><div style="position:inline;width:18px;height:12px;background:<%=tcolor%>;border:1px solid #999999;">&nbsp;</div></td>
      <td width="6">&nbsp;</td>
      <td><%=cStatus%></td>
    </tr>
    </table>
    </td>
  </tr>
  </table>
  </td>
</tr>
<tr>
  <td colspan="2" style="border-top:1px solid #ffffff;"><b>Job Cost Report</b></td>
</tr>
<tr> 
  <td width="50%" valign="top">
  <!-- begin column 1 -->
  <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr valign="top" bgcolor="#336699"> 
    <td colspan="2"><span class="standardheader">Job-to-Date Costs</span></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td width="70%">Labor Hours</td>
    <td width="30%" align="right"><%=formatnumber(rs("jtd_labor_units"),2)%></td>
  </tr>

  <tr valign="top" bgcolor="#ffffff">   
    <td><a href=<%="javascript:po('" & c & "','" & rs("job") & "')"%>>Materials Cost</a></td>
    <td align="right"><%=formatcurrency(rs("jtd_material_cost"),2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td><a href=<%="javascript:open_po('" & c & "','" & rs("job") & "')"%>>Open POs</a></td>
    <td align="right"><%=formatcurrency(openpo,2)%></td>
  </tr>
  
  <tr valign="top" bgcolor="#ffffff"> 
    <td style="border-top:1px solid #000000;border-left:1px solid #000000;border-bottom:1px solid #000000;"><b>Job-to-Date Cost</b></td>
    <td style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;" align="right"><b><%=formatcurrency(rs("jtd_cost")+openpo,2)%></b></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td>Current Hours</td>
    <td align="right"><%=formatnumber(hours,2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td>Log Contract Amt 1</td>
    <td align="right"><%=formatcurrency(rs("msc_job_amt_1"),2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td>Log Contract Amt 2</td>
    <td align="right"><%=formatcurrency(rs("msc_job_amt_2"),2)%></td>
  </tr>
  </table>
  <!-- end column 1 -->
  </td>
  <td width="50%" valign="top">
  <!-- begin column 2 -->
  <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr><td colspan="2" bgcolor="#dddddd"><b>Contract &amp; Billing</b></td></tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td>Original Contract</td>
    <td align="right"><%=formatcurrency(rs("Original_contract_amount"),2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td>Change Orders</td>
    <td align="right"><%=formatcurrency(rs("JTD_Aprvd_Contract_Chgs"),2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td>Revised Contract</td>  
    <td align="right"><%=formatcurrency(rs("Revised_Contract_Amount"),2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td>% Complete</td>
    <td align="right"><%=formatnumber(rs("percent_complete"),0)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">
  	<td>Current Pending Amount to Bill</td>
	<%if isNULL(totalInvoices) then totalInvoices = 0%>
	<td align="right"><%=formatcurrency(totalInvoices,2)%></td>
</tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td><a href=<%="javascript:invoice('" & c & "','" & rs("job") & "')"%>>Amount Billed</a></td>
    <td align="right"><%=formatcurrency(rs("JTD_work_billed"),2) %></td>
  </tr>
  </table>
  <!-- end column 2 -->
  </td>
</tr>
<tr>
	
</tr>
<tr>
  <td style="border-bottom:1px solid #cccccc;">Updated as of <%=formatdatetime(crdate,0)%><br>&nbsp;</td>
  <!--<td align="right" style="border-bottom:1px solid #cccccc;"><font color="red"><b>*</b></font>  Only includes invoices as of 10/1/2003</td>-->
</tr>
</table>
		   
<%
set cnn = nothing 
end if %>

</body>
</html>
<% end if %>