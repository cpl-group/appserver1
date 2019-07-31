<%@Language="VBScript"%>
<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<%
'7/5/2008 N.Ambo added link to job tasks page

dim item, var
item= Request("select")
var= Request("findvar")
dim jid,permissionflag
jid = ""

'if session("corp")=5 then
permissionflag=""
'else
'permissionflag="1"
'end if

dim msg
if isempty(var) then
	msg=""
	'Write a browser-side script to update another frame (named
	'detail) within the same frameset that displays this page.
	Response.Write "<script>" & vbCrLf
	Response.Write "parent.location = " & _
	Chr(34) & "poindex.asp?msg=" & msg & Chr(34) & vbCrLf
	Response.Write "</script>" & vbCrLf
end if

Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

dim vendorSelect, vendor, comp
rst1.open "select * from companycodes where active = 1 and code <> 'AC' order by name", cnn1
do until rst1.eof
	vendorSelect = vendorSelect & "SELECT [name], vendor, '"&rst1("code")&"' as comp FROM ["&rst1("code")&"_MASTER_APM_VENDOR] UNION all "
	rst1.movenext
loop
rst1.close
vendorSelect = "(SELECT distinct * FROM (" & vendorSelect
vendorSelect = left(vendorSelect,len(vendorSelect)-10) & ") v)"
'vendorSelect = "(SELECT distinct * FROM (SELECT [name], vendor FROM [" & replace(vendorSelect,"|","_MASTER_APM_VENDOR] UNION all SELECT [name], vendor FROM [")&"_MASTER_APM_VENDOR]) v)"

sqlstr= "select po.submitted, po.accepted, po.approved, po.closed, po.po_total,employees.[first name]+ ' ' + employees.[last name] as req, ISNULL(app.[First name] + ' ' + app.[Last name], 'N/A') AS appuser,ltrim(str(po.Jobnum)) " _
& " + '.' + ltrim(str(po.POnum)) as ponumber, po.vendor,po.po_total,po.podate,po.requistioner, case when po.vid<>'0' then vs.name else po.vendor end as vendorname from employees join po on " _
& "substring(employees.username,7,20)=po.requistioner INNER JOIN master_job m ON po.jobnum=m.id left outer join employees app on  app.username = { fn CONCAT('ghnet\', po.approved_user) } "_
& "LEFT JOIN "&vendorSelect&" vs ON vs.vendor=po.vid and vs.comp=m.company "
if item="jobnum" then
	sqlstr = sqlstr	& "where jobnum = " &	var & "order by podate desc "
	jid = trim(var)
elseif item = "vendor" and not isempty(var) then
	sqlstr = sqlstr	& "where (po.vendor like '%" & var & "%' or vs.name like '%" & var & "%') order by podate desc "
elseif item="description" and not isempty(var) then
	sqlstr = sqlstr	& "where description like '%" & var & "%'order by podate desc "
elseif item="requistioner" and not isempty(var) then
	sqlstr = sqlstr	& "where employees.[first name]+ ' ' + employees.[last name] like '%" & var & "%'order by podate desc "
elseif item="date" and not isempty(request("todate")) and not isempty(request("fromdate")) then
	sqlstr = sqlstr	& "where po.podate >= '" & request("fromdate") & "' and po.podate <= '" & request("todate") & "' order by podate desc"
end if
'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1
%>
<html>
<head>
<title><%
if request("printview") = "yes" then
	%>Requisition Forms From <%=request("fromdate")%> to <%=request("todate")%><%
else%>
	 View Requisition Form<%
end if%>
</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
//<!--
function highlight(tRow){
	tRow.style.backgroundColor = "lightgreen";
}

function unlight(tRow){
	tRow.style.backgroundColor = "white";

}
//jid is <%=jid%>
function edit_job(jid) {
	theURL="/genergy2_intranet/opsmanager/joblog/updatejob.asp?jid=" + jid
	//window.document.all['genjobtable'].Border ="1"
	window.document.all['genjobtable'].bgColor ="#dddddd"
	openwin(theURL,750,400)
}
function openwin(url,mwidth,mheight){
	newwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function newTask(jid){	
	theURL=	"/genergy2_intranet/opsmanager/joblog/edittasks.asp?mode=new&jobnum=" + jid	
	//window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,400,230)
	}

function newPO(){
	<% if request("caller")="joblog" then %>
		document.location="newpo.asp?jid=<%=jid%>&caller=joblog"
	<% else %>
		document.location="newpo.asp?jid=<%=jid%>"
	<% end if %>
}

function viewinvoice(jid) {
	theURL="timesheetmain.asp?flag=0&job=" + jid
	//window.document.all['genjobtable'].Border ="1"
	window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,750,550)
}

var loaded = 0;
	function preloadImages(){
	edit_jobOn = new Image(); edit_jobOn.src = "images/btn-edit_job-1.gif";
	edit_jobOff = new Image(); edit_jobOff.src = "images/btn-edit_job.gif";
	new_poOn = new Image(); new_poOn.src = "images/btn-new_po-1.gif";
	new_poOff = new Image(); new_poOff.src = "images/btn-new_po.gif";
	invoiceOn = new Image(); invoiceOn.src = "images/btn-invoice-1.gif";
	invoiceOff = new Image(); invoiceOff.src = "images/btn-invoice.gif";
	
	loaded = 1;
}

function msover(img){
	if ((loaded) && (document.all)) {
		img.src = eval (img.id + "On.src");
	}
}

function msout(img){
	if ((loaded) && (document.all)) {
		img.src = eval (img.id + "Off.src");
	}
}

//visual feedback functions for img buttons
function buttonOver(obj,clr){
	if (arguments.length == 1) {
		clr = "#336699";
	}
	obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
	if (arguments.length == 1) {
		clr = "#000000";
	}
	obj.style.border = "1px solid " + clr;
}

function buttonOut(obj,clr){
	if (arguments.length == 1) {
		clr = "#eeeeee";
	}
	obj.style.border = "1px solid " + clr;
}

//-->
</script>
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>
<%
dim onload
onload = "preloadImages();"
if request("printview") = "yes" then
	onload = onload & "print();"
end if
%>
<body bgcolor="#eeeeee" onload="<%=onLoad%>">
<form>
<%
if request("caller")="joblog" then
	dim rst2
	Set rst2 = Server.CreateObject("ADODB.recordset")
	Dim tcolor,Desc,company,job,jtype,cstatus
	rst2.Open "SELECT description, type,company, job, status FROM MASTER_JOB WHERE id='"&jid&"'", cnn1
	if not rst2.EOF then
		Desc = rst2("description")
		company = rst2("company")
		job = rst2("job")
		jtype=left(rst2("type"),6)
		cStatus = rst2("status")
		Select Case cStatus
			case "In progress"
				tcolor = "#66ff66"
			case "Unstarted"
				tcolor = "#ffcc00"
			case "Closed"
				tcolor = "#cc0033"
		end select 
	end if 'end not rst2.eof
	rst2.close
	%>
	<table border=0 cellpadding="3" cellspacing="0" width="100%" id="genjobtable" bgcolor="#eeeeee" style="border-top:1px solid #cccccc;">
		<tr> 
			<td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
				<table border="0" cellspacing="0" cellpadding="3" width="100%">
					<tr>
						
            <td> <a href="<%="/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=" & jid %>">General 
              Info</a> |&nbsp;<a href="/gEnergy2_Intranet/opsmanager/joblog/jobestimates.asp?jid=<%=jid%>&job=<%=job%>&c=<%=company%>&mode=view">Job 
              Estimates</a>&nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobtime.asp?jid=" &jid%>">Job 
              Time</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job 
              Folder</a> &nbsp;|&nbsp; <b>Requisition Forms</b> &nbsp;|&nbsp; 
              <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=job%>&jid=<%=jid%>&caller=<%=request("caller")%>"> 
              Job Cost </a> &nbsp;|&nbsp; <a href="/genergy2_intranet/opsmanager/joblog/viewchange.asp?jid=<%=jid%>">Change 
              Orders</a> &nbsp;|&nbsp; <a href="/genergy2_Intranet/opsmanager/joblog/jobtasks.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Job
              Tasks</a></td>
					</tr>
				</table>
			</td>
			<td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;" bgcolor="#eeeeee">
				<img src="images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" 
					onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">
			</td>
		</tr>
		<tr> 
			<td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
				<table border="0" cellspacing="0" cellpadding="3" width="100%">
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
					</tr><%
					if not rst1.EOF then %>
						<tr>
							<td>Total requisition forms:</td>
							<td>
								<b>
								<%
								dim rst3, str
								Set rst3 = Server.CreateObject("ADODB.recordset")
								str="select sum(po_total) as sum1 from po where jobnum=" & jid & " and accepted = 1"
								rst3.Open str, cnn1, 0, 1, 1
								if rst3("sum1") > 0 then %>
									<%=Formatcurrency(rst3("sum1"))%>
									<% else%>
									$0 
								<% end if 'sum1 > 0
								rst3.close
								%>
								</b>
							</td>
						</tr>
						<tr>
							<td>Total hours invoiced:</td>
							<td>
								<b>
								<%
								dim rst4
								Set rst4 = Server.CreateObject("ADODB.recordset")
								str="select sum(hours) as sum1 from invoice_submission where jobno=" & jid & " and submitted=1"
								rst4.Open str, cnn1, 0, 1, 1
								if trim(rst4("sum1")) > 0 then %>
									<%=rst4("sum1")%><%
								else %>
									0<%
								end if 'sum1 > 0
								rst4.close 
								%>
								</b>
							</td>
						</tr><%
					end if 'not rst1.eof %>
				</table>
			</td>
		</tr><%
		if rst1.EOF then%>
			<tr>
				<td colspan="2" bgcolor="#ffffff" style="border-top:2px outset #ffffff;padding:6px;">
					No requisition forms were found under the requested job id<br>&nbsp;
				</td>
			</tr>
			<%
		end if%>
	</table><%
end if 'caller=joblog 

if rst1.EOF then 

	if request("caller")<>"joblog" then%>
		<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#ffffff">
			<tr>
				<td style="border-top:2px outset #ffffff;border-bottom:1px solid #cccccc;">&nbsp;No results found. Please try another search.<br>&nbsp;</td>
			</tr>
		</table><%
	end if 'caller=joblog

Else 'if not rst1.eof
	x=0
	%>
	<table border=0 cellpadding="3" cellspacing="1" width="100%">
		<tr bgcolor="#dddddd" style="font-weight:bold;"> 
			<td width="10%">RF Date</td>
			<td width="10%">RF Number</td>  
			<td width="10%">RF Amount</td>  
			<td width="25%">Vendor</td>
			<td width="20%">Requisitioner</td>
			<td width="10%">Approved By</td>
			<td width="10%">Status</td>
		</tr>
	</table>
	<%
	dim height
	height = 525
	if request("caller") = "joblog" then
		height = 200
	end if
	if request("printview") <> "yes" then%>
	<div style="width:100%; overflow:auto; height:<%=height%>px;"> 
	<%end if%> 
		<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#dddddd"><%
			While not rst1.EOF 
				dim statusOfPO			
				if (rst1("submitted")) and not (rst1("accepted")) then
					statusOfPO= "Submitted"
				elseif rst1("accepted") and (not rst1("approved")) then
					statusOfPO = "Accepted"
				elseif rst1("approved") and (not rst1("closed")) then
					statusOfPO = "Approved"
				elseif rst1("closed") then
					statusOfPO = "Closed"
				else
					statusOfPO = "Not yet submitted"
				end if		%>
				<tr bgcolor="#ffffff" onMouseOver="highlight(this);" onMouseOut="unlight(this);" 
					onClick="document.location='<%="poview.asp?po=" & rst1("ponumber") & "&jid=" & jid & "&caller=" & request("caller") %>';" style="cursor:hand"> 
					
					<td width="10%"><%=rst1("podate")%></td>
					<td width="10%"><%=rst1("ponumber")%></td>
					<td width="10%" align="right"><%=formatcurrency(rst1("po_total"))%></td>
					<td width="25%"><%=rst1("vendorname")%></td>
					<td width="20%"><%=rst1("req")%><input type="hidden" name="job" value="<%=rst1("requistioner")%>"></td>
					<td width="10%"><%=rst1("appuser")%></td>
					<td width="10%"><%=statusOfPO%></td>
				</tr>
				<%
				x=x+1
				rst1.movenext
			Wend
			%>
		</table>
	</div>
	<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#dddddd">
		<tr>
			<td colspan="5" bgcolor="#eeeeee"><i><%=x%> requisition forms found</i></td>
		</tr>
	</table>
	<%
end if
rst1.close
%>

<%if request("caller")="joblog" then%>
	<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #cccccc;">
		<tr>
		<td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;"> 
			<!-- #include virtual="/includes/jobfooterbuttons.asp" -->
		</td>
			<td align="right" style="border-top:1px solid #ffffff;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
		</tr>
	</table>
<% end if %>
</form>
</body>
</html>