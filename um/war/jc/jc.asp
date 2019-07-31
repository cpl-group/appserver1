<%option explicit%>
<%'1/10/2008 N.Ambo added functionality to allow only users in the group 'Job Status Admins' to view the detailed amounts of the job.
'For other users these amounts will not be visible (requested by Danny)
%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<title>Job Cost Report</title>
<script language="JavaScript" type="text/javascript">
//<!--

//visual feedback functions for img buttons
function buttonOver(obj,clr){
  if (arguments.length == 1) { clr = "#336699"; }
  obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
  if (arguments.length == 1) { clr = "#000000"; }
  obj.style.border = "1px solid " + clr;
}

function newTask(jid){	
	theURL=	"/genergy2_intranet/opsmanager/joblog/edittasks.asp?mode=new&jobnum=" + jid	
	//window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,400,230)
	}

function buttonOut(obj,clr){
  if (arguments.length == 1) { clr = "#eeeeee"; }
  obj.style.border = "1px solid " + clr;
}

//-->
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<%
'http params
dim crdate,j,c,avg,wip,tcost,OPENPO,link,amt_paid,aTable,sql, prefix,hours, jid, permissionflag
dim cnn1, cmd, rs, prm
set cnn1 = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")
' open connection
cnn1.open getConnect(0,0,"Intranet")
cnn1.CursorLocation = adUseClient

jid = Request("jid")
c = request("c")

'check the rights of the user, the user's rigths determines what can be viewed on the screen regarding amount values
dim approved
if allowgroups("Job Status Admins") then
   approved = true
else 
    approved = false
end if

dim totalInvoices
rs.open "select sum(s.amt) as bignumber from (select invoice_amt amt from invoice_submission  where jobno = '"&jid&"' and flag=0 group by jobno, invoice_date, invoice_amt) s", cnn1

if rs.eof then 
	totalInvoices = 0 
else 
	if isnull(rs("bignumber")) then 
		totalInvoices = 0 
	else
		totalInvoices = rs("bignumber")
	end if
end if
rs.close
%>
  	<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">
	<table border=0 cellpadding="3" cellspacing="0" id="genjobtable" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #cccccc;border right:1px solid #cccccc;">
	<%
		'if request("caller")="joblog" then
			Dim tcolor,Desc,company,job,jtype,cstatus, rst2
			Set rst2 = Server.CreateObject("ADODB.recordset")
			rst2.Open "SELECT description, type, company, job, status FROM MASTER_JOB WHERE id='"&jid&"'", cnn1
			if not rst2.EOF then
				Desc = rst2("description")
				company = rst2("company")
				job = rst2("job")
				jtype=left(rst2("type"),6)
				cStatus = rst2("status")
				
				Select Case lCase(cStatus)
				case "in progress"
					tcolor = "#66ff66"
				case "unstarted"
					tcolor = "#ffcc00"
				case "closed"
					tcolor = "#cc0033"
				end select 
			end if
			rst2.close
			%>
			<tr> 
				<td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
					<table border="0" cellspacing="0" cellpadding="3" width="100%">
						<tr>
							
          <td> <a href="/gEnergy2_Intranet/opsmanager/joblog/viewjob.asp?jid=<%=jid %>">General 
            Info</a> &nbsp;|&nbsp;<a href="/gEnergy2_Intranet/opsmanager/joblog/jobestimates.asp?jid=<%=jid%>&job=<%=job%>&c=<%=c%>&mode=view">Job 
            Estimates</a> &nbsp;|&nbsp;<a href="/gEnergy2_Intranet/opsmanager/joblog/jobtime.asp?jid=<%=jid%>">Job 
            Time</a> &nbsp;|&nbsp; <a href="/gEnergy2_Intranet/opsmanager/joblog/jobfolder.asp?jid=<%=jid%>">Job 
            Folder</a>&nbsp;|&nbsp; <a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requestion 
            Forms</a>&nbsp;|&nbsp; <b>Job Cost</b>&nbsp;|&nbsp; <a href="/gEnergy2_Intranet/opsmanager/joblog/viewchange.asp?jid=<%=jid%>">Change 
            Orders</a>&nbsp;|&nbsp; <a href="/genergy2_Intranet/opsmanager/joblog/jobtasks.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Job
              Tasks</a> </td>
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
<%' end if %>
<%


Select case ucase(c)
case "GE"
	j = request ("jg")
	aTable = "GE_MASTER_JOB"
	prefix = "jg"
case "GY"
	j = request ("jg")
	aTable = "GY_MASTER_JOB"
	prefix = "jg"
case "IL"
		j = request ("ji")
		aTable = "IL_MASTER_JOB"
		prefix = "ji"
case "NY"
		j = request ("jn")
		aTable = "NY_MASTER_JOB"
		prefix = "jn"
CASE "GS"
     j = request ("jg")
	 aTable = "GS_MASTER_JOB"
	 prefix = "jg"
CASE "EM"
     j = request ("je")
	 aTable = "EM_MASTER_JOB"
	 prefix = "je"
end select

if session("corp")=5 then
  permissionflag=""
else
  permissionflag="1"
end if

'adodb vars



sql = "select job, description,address_1,address_2 from master_job where job like '%" & j & "%' or description like '%" & j & "%' or Address_1 like '%" & j & "%' or Address_2 like '%" & j & "%'  order by job"
rs.open sql,cnn1
'response.write sql
'response.end
if rs.eof or not allowGroups("Genergy_Corp,AR_Admin,gAccounting,IT Services") then
	%>
		<tr>
			<td colspan = 2>No records found or you do not have permissions to this module.</td>
		</tr>
	<%
	if rs.eof then 
		set cnn1 = nothing 
		rs.close
	end if 
end if

if rs.recordcount > 1 then 
	%>
<tr>
  <td style="border-top:1px solid #ffffff;"><b>Job Cost Report</b></td>
  <td align="right" style="border-top:1px solid #ffffff;" bgcolor="#eeeeee"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr bgcolor="#dddddd">
    <td width="11%"><b>Job Number</b></td>
    <td width="34%"><b>Description</b></td>
    <td><b>Address</b></td>
    <td></td>
  </tr>
  <% 
  while not rs.EOF 
  %>
  
  <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:document.location='jc.asp?<%=prefix%>=<%=rs("job")%>&c=<%=c%>'" bgcolor="#ffffff"> 
    <td width="11%"><%=rs("job")%></td>
    <td width="34%"><%=rs("description")%></td>
    <td width="32%"><%=rs("address_1")%></td>
    <td width="23%"><%=rs("address_2")%></td>
  </tr>
  <% 
  rs.movenext
  wend
  %>
</table>
</body>
</html>
<%
  rs.close
else
  j = rs("job")
  rs.close
  
  rs.open "SELECT crdate FROM sysobjects WHERE name = '"&c&"_activity_ara_status'",cnn1
  crdate=rs(0)
  rs.close
  
  if c="GY" or c="GE" or c="EM" then
  'get sum of hours from time sheet
  rs.open "SELECT isnull(sum(hours)+sum(overt),0) FROM times WHERE jobno = '"&split(j,"-")(1) &"'",cnn1
  hours=rs(0)
  rs.close



  sql="select distinct isnull(a.amount-a.amount_invoiced,0),a.commitment from "&c&"_master_po_item a,"&c&"_master_po b where b.closed=0 and a.job='"&j&"' and a.commitment=b.commitment order by a.commitment"
  else
  
  'sql="select isnull(sum(a.amount)-sum(a.amount_invoiced),0) from "&c&"_master_po_item a,"&c&"_master_po b where b.closed=0 and a.job='"&j&"' and  a.commitment=b.commitment"
  sql="select distinct isnull(a.amount-a.amount_invoiced,0),a.commitment from "&c&"_master_po_item a,"&c&"_master_po b where b.closed=0 and a.job='"&j&"' and  a.commitment=b.commitment order by a.commitment"
'response.write sql
'response.end
  rs.open "SELECT isnull(sum(hours)+sum(overt),0) FROM times WHERE jobno = '"&split(j,"-")(1) &"'",cnn1
  hours=rs(0)
  rs.close
  
  end if

  rs.open sql,cnn1


  while not rs.EOF 
openpo = openpo + rs(0)

rs.movenext
wend
 
  rs.close

 ' rs.open "SELECT isnull(SUM(cash_receipt),0) FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_activity_ara_activity a on  a.invoice=s.invoice WHERE  a.job = '" + j + "'  and a.amount > 0",cnn1
  rs.open "SELECT distinct isnull(cash_receipt,0),a.invoice FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_activity_ara_activity a on  a.invoice=s.invoice WHERE  a.job = '" + j + "'  and a.amount > 0",cnn1
   ' response.write "SELECT distinct isnull(cash_receipt,0),a.invoice FROM "&c&"_ACTIVITY_ARA_STATUS s inner join "&c&"_activity_ara_activity a on  a.invoice=s.invoice WHERE  a.job = '" + j + "'  and a.amount > 0"
   ' response.end
   while not rs.EOF 
   amt_paid=amt_paid + abs(rs(0))
   rs.movenext
   wend
  rs.close
  
  response.write(c)
  response.Write(j)
  response.end
  ' specify stored procedure 
  
  cmd.CommandText = "sp_job_cost"
  cmd.CommandType = adCmdStoredProc
  
  Set prm = cmd.CreateParameter("c", adchar, adParamInput,2)
  cmd.Parameters.Append prm
  
  Set prm = cmd.CreateParameter("j", advarchar, adParamInput,9)
  cmd.Parameters.Append prm
  
  ' assign internal name to stored procedure
  cmd.Name = "test"
  Set cmd.ActiveConnection = cnn1
  
  'return set to recordset rs
  cnn1.test c,j, rs
  
  dim hoursMsg
  if rs.eof then
  %>
  
  </head>
  <body bgcolor="#ffffff">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td colspan="2">No records found.</td>
  </tr>
  </table>
  </body>
  </html>
  <%
    set cnn1 = nothing 
    response.end
  else
  if rs("jtd_labor_units") = 0 then
    hoursMsg = "<br><span class=""notetext"">No hours/times for this job posted within the account system</span>"
  else
  avg = (rs("jtd_labor_cost") + rs("jtd_overhead_cost") + rs("jtd_other_cost") )/ rs("jtd_labor_units")
  wip =  (rs("Revised_Contract_Amount") * rs("percent_complete")/100)-rs("jtd_work_billed")
  'tcost = ( cdbl(hours) * cdbl(avg)) + cdbl(rs("jtd_subcontract_cost"))+ cdbl(rs("jtd_material_cost"))
  end if
  %>
  <script language="JavaScript1.2">
  function po(c,j,costcode) {
  if (c=="IL" && c=="NY") {
  var urlLink     = "/um/war/apa/material_cost.asp?c=" + c + "&ji=" + j + "&costcode="+costcode
  }
  else {
  var urlLink     = "/um/war/apa/material_cost.asp?c=" + c + "&jg=" + j + "&costcode="+costcode
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
  } openwin(urlLink,800,400)
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

  function edit_job(jid) {
    theURL="/genergy2_intranet/opsmanager/joblog/updatejob.asp?jid=" + jid
    //window.document.all['genjobtable'].Border ="1"
    window.document.all['genjobtable'].bgColor ="#dddddd"
    openwin(theURL,750,400)
  }
  
  function newPO(){
    <% if request("caller")="joblog" then %>
    document.location="/um/opslog/newpo.asp?jid=<%=jid%>&caller=joblog"
    <% else %>
    document.location="/um/opslog/newpo.asp?jid=<%=jid%>"
    <% end if %>
  }
  
  function viewinvoice(jid) {
    theURL="/um/opslog/timesheetmain.asp?flag=0&job=" + jid
    //window.document.all['genjobtable'].Border ="1"
    window.document.all['genjobtable'].bgColor ="#999999"
    openwin(theURL,750,550)
  }

  </script>
  

</head>
  
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#eeeeee">
  
<table border=0 cellpadding="3" cellspacing="0" id="genjobtable" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #cccccc;border-right:1px solid #cccccc;">
<%
dim sqlstr, rst3, jtosearch
set rst3 = server.createobject("ADODB.Recordset")
if request("caller")="joblog" then
  jtosearch = jid
else 
  jtosearch = split(j,"-")(1) 
end if
sqlstr = "SELECT description, type, company, job, status FROM MASTER_JOB WHERE id='"&jtosearch&"'"
rst3.open sqlstr, cnn1

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
'Retrieve estimate numbers for this job
sqlstr = "select sum(est_labor_units) as est_LU, sum(est_labor_cost) as est_LC,sum(est_overhead_cost) as est_OC,sum(est_burden_cost) as est_BC,max(est_labor_unit_cost) as est_LUC,sum(est_material_cost) as est_MC from "&c&"_JOB_Estimates where job = '"&job&"'"
rst3.open sqlstr, cnn1

if not rst3.EOF then
	Dim est_LU, est_LC, est_OC, est_BC, est_LUC, est_MC
  
	est_LU = rst3("est_LU")
	est_LC = rst3("est_LC")
	est_OC = rst3("est_OC")
	est_BC = rst3("est_BC")
	est_LUC = rst3("est_LUC")
	est_MC = rst3("est_MC")
end if
rst3.close
%>
<tr>
  <td><b>Job Cost Report</b><% if not request("caller")="joblog" then %><br><%=Desc%>&nbsp; (<%=company%>: &nbsp;<%=job%>) &nbsp;<%=jtype%><% end if %><%=hoursMsg%></td>
  <td align="right" bgcolor="#eeeeee"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
<tr> 
  <td width="50%" valign="top">
  <!-- begin column 1 -->
  <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr valign="top" bgcolor="#336699"> 
    <td colspan="4"><span class="standardheader">Job-to-Date Costs</span></td>
  </tr>
  <tr valign="top" bgcolor="#cccccc"> 
    <td width="25%">&nbsp;</td>
    <td width="25%" align="center"><span class="standardheader">Actual</span></td>
    <td width="25%" align="center"><span class="standardheader">Estimate</span></td>
	<td width="25%" align="center"><span class="standardheader">Difference</span></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td width="25%" nowrap>Labor Hours</td>
    <td width="25%" align="right"><%=formatnumber(rs("jtd_labor_units"),2)%></td>
    <td width="25%" align="right"><%if isnumeric(est_LU) Then%><%=formatnumber(est_LU,2)%><%else%>NA<%end if%></td>
	<td width="25%" align="right"><%if isnumeric(est_LU) and trim(est_LU) <> "0" Then%><%=formatpercent(cdbl(rs("jtd_labor_units"))/est_lu,2)%><%else%>NA<%end if%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td nowrap>Labor Cost</td>
    <td width="25%" align="right"><%if approved then%><%=formatcurrency(rs("jtd_labor_cost"),2)%><%else%>**********<%end if%></td>
    <td width="25%" align="right"><%if isnumeric(est_LU) Then%><%=formatcurrency(est_LC,2)%><%else%>NA<%end if%></td>
	<td width="25%" align="right"><%if isnumeric(est_LU) and trim(est_LU) <> "0" Then%><%=formatpercent(cdbl(rs("jtd_labor_Cost"))/est_LC,2)%><%else%>NA<%end if%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">
    <td nowrap>Overhead Cost</td>  
    <td width="25%" align="right"><%if approved then %><%=formatcurrency(rs("jtd_overhead_cost"),2)%><%else%>**********<%end if%></td>
    <td width="25%" align="right"><%if isnumeric(est_OC) Then%><%=formatcurrency(est_OC,2)%><%else%>NA<%end if%></td>
	<td width="25%" align="right"><%if isnumeric(est_OC) and trim(est_OC) <>"0" Then%><%=formatpercent(cdbl(rs("jtd_overhead_cost"))/est_OC,2)%><%else%>NA<%end if%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td nowrap>Burden Cost</td>
    <td width="25%" align="right"><%if approved then %><%=formatcurrency(rs("jtd_other_cost"),2)%><%else%>**********<%end if%></td>
    <td width="25%" align="right"><%if isnumeric(est_BC) Then%><%=formatcurrency(est_BC,2)%><%else%>NA<%end if%></td>
	<td width="25%" align="right"><%if isnumeric(est_BC) and trim(est_BC) <>"0" Then%><%=formatpercent(cdbl(rs("jtd_other_cost"))/est_BC,2)%><%else%>NA<%end if%></td>
  </tr>
  <tr valign="top" bgcolor="#eeeeee">
    <td nowrap>Average labor cost per hour:</td>
    <td width="25%" align="right"><%if approved then%><%=formatcurrency(avg,2)%><%else%>**********<%end if%></td>
    <td width="25%" align="right"><%if isnumeric(est_LUC) Then%><%=formatcurrency(est_LUC,2)%><%else%>NA<%end if%></td>
	<td width="25%" align="right"><%if isnumeric(est_LUC) and trim(est_LUC) <>"0" Then%><%=formatpercent(cdbl(avg)/est_LUC,2)%><%else%>NA<%end if%></td>
  </tr>

  <tr valign="top" bgcolor="#ffffff">   
    <td nowrap><a href=<%="javascript:po('" & c & "','" & rs("job") & "','004')"%>>Materials Cost</a></td>
    <td width="25%" align="right"><%=formatcurrency(rs("jtd_material_cost"),2)%></td>
    <td width="25%" align="right"><%if isnumeric(est_MC) Then%><%=formatcurrency(est_MC,2)%><%else%>NA<%end if%></td>
	<td width="25%" align="right"><%if (isnumeric(est_MC) and trim(est_MC)<>"0") and isnumeric(rs("jtd_material_cost")) Then%><%=formatpercent(cdbl(rs("jtd_material_cost"))/est_MC,2)%><%else%>NA<%end if%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td colspan=3><a href=<%="javascript:open_po('" & c & "','" & rs("job") & "')"%>>Open POs</a></td>
    <td align="right"><%=formatcurrency(openpo,2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td colspan=3><a href=<%="javascript:po('" & c & "','" & rs("job") & "','005')"%>>Subcontractor Cost</a></td>
    <td align="right"><%=formatcurrency(rs("jtd_subcontract_cost"),2)%></td>
  </tr>
  
  <tr valign="top" bgcolor="#ffffff"> 
    <td colspan=3 style="border-top:1px solid #000000;border-left:1px solid #000000;border-bottom:1px solid #000000;"><b>Job-to-Date Cost</b></td>
    <td style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;" align="right"><b><%=formatcurrency(rs("jtd_cost")+openpo,2)%></b></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td colspan=3>Current Hours</td>
    <td align="right"><%=formatnumber(hours,2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td colspan=3>Log Contract Amt 1</td>
    <td align="right"><%=formatcurrency(rs("msc_job_amt_1"),2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td colspan=3>Log Contract Amt 2</td>
    <td align="right"><%=formatcurrency(rs("msc_job_amt_2"),2)%></td>
  </tr>
  <!--<tr valign="top" bgcolor="#ffffff">  
    <td><b>Projected Cost</b></td>
    <td align="right"><b><%=formatcurrency(tcost,2)%></b></td>
  </tr>-->
  </table>
  <!-- end column 1 -->
  </td>
  <td width="50%" valign="top">
  <!-- begin column 2 -->
  <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr><td colspan="2" bgcolor="#dddddd"><b>Contract &amp; Billing</b></td></tr>
  <!--<tr valign="top" bgcolor="#ffffff"> 
    <td>Original Contract</td>
    <td align="right"><%=formatcurrency(rs("Original_contract_amount"),2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td>Change Orders</td>
    <td align="right"><%=formatcurrency(rs("JTD_Aprvd_Contract_Chgs"),2)%></td>
  </tr>-->
  <tr valign="top" bgcolor="#ffffff">  
    <td>Revised Contract</td>  
    <td align="right"><%=formatcurrency(rs("Revised_Contract_Amount"),2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td>% Complete</td>
    <td align="right"><%=formatnumber(rs("percent_complete"),0)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td>Job Value</td>
    <td align="right"><%=formatcurrency(rs("Revised_Contract_Amount") * rs("percent_complete")/100,2) %></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">
  	<td><a href=<%="javascript:openwin('/um/war/ara/ReqdInv.asp?c="& c &"&jid="& jid &"&j="& job &"',800,400)"%>>Current Pending Amount to Bill</a></td>
	<td align="right"><%=formatcurrency(totalInvoices,2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td><a href=<%="javascript:invoice('" & c & "','" & rs("job") & "')"%>>Amount Billed</a> (Amount w/o Tax)</td>
    <td align="right"><%=formatcurrency(rs("JTD_work_billed"),2) %></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td>Amt. paid</td>
    <td align="right"><%=formatcurrency(amt_paid,2)%></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff">  
    <td>Work-in-Progress (WIP)</td>
    <td align="right"><%=formatcurrency(wip,2)%></td>
  </tr>
  </table>
  <!-- end column 2 -->
  </td>
</tr>
<tr>
  <td style="border-bottom:1px solid #cccccc;">Updated as of <%=formatdatetime(crdate,0)%><br>&nbsp;</td>
 <!-- <td align="right" style="border-bottom:1px solid #cccccc;"><font color="red"><b>*</b></font>  Only includes invoices as of 10/1/2003</td>-->
</tr>
</table>

<% if request("caller")="joblog" then %>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #cccccc;">
<tr>
		<td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;"> 
			<!-- #include virtual="/includes/jobfooterbuttons.asp" -->
		</td>
  <td align="right" style="border-top:1px solid #ffffff;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
<% end if %>
<br><br>

<%
set cnn1 = nothing 
end if %>
</body>
</html>
<% end if %>