<%@Language="VBScript"%>
<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim item, var
item= Request("select")
var= Request("findvar")
dim jid,permissionflag
jid = trim(Request("findvar"))

'if session("corp")=5 then
permissionflag=""

Dim cnn1, rst1, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

sqlstr = "select * from master_job_Tasks where jobid = " &jid

rst1.Open sqlstr, cnn1
%>
<html>
<head>
<title>View Job Tasks</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">

function highlight(tRow){
	tRow.style.backgroundColor = "lightgreen";
}

function unlight(tRow){
	tRow.style.backgroundColor = "white";
}

function newPO(){
	<% if request("caller")="joblog" then %>
		document.location="/um/opslog/newpo.asp?jid=<%=jid%>&caller=joblog"
	<% else %>
		document.location="newpo.asp?jid=<%=jid%>"
	<% end if %>
}

function newTask(jid){	
	theURL=	"/genergy2_intranet/opsmanager/joblog/edittasks.asp?mode=new&jobnum=" + jid	
	//window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,400,230)

}

function viewinvoice(jid) {
	theURL="/um/opslog/timesheetmain.asp?job=" + jid
	//window.document.all['genjobtable'].Border ="1"
	window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,750,400)
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
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
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
				<table border="0" cellspacing="0" cellpadding="3" width="100%" >
					<tr>
						
            <td> <a href="<%="/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=" & jid %>">General 
              Info</a> |&nbsp;<a href="/gEnergy2_Intranet/opsmanager/joblog/jobestimates.asp?jid=<%=jid%>&job=<%=job%>&c=<%=company%>&mode=view">Job 
              Estimates</a>&nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobtime.asp?jid=" &jid%>">Job 
              Time</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job 
              Folder</a> &nbsp;|&nbsp;<a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requsition
              Forms</a> &nbsp;|&nbsp; <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=job%>&jid=<%=jid%>&caller=<%=request("caller")%>"> 
              Job Cost </a> &nbsp;|&nbsp; <a href="/genergy2_intranet/opsmanager/joblog/viewchange.asp?jid=<%=jid%>">Change 
              Orders</a>  &nbsp;|&nbsp; <b>Job Tasks</b></td>
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
				<table border="0" cellspacing="0" cellpadding="3" width="100%" ID="Table2">
					<tr>
						<td width="130">Job:</td>
						<td><%=Desc%>&nbsp; (<%=company%>: &nbsp;<%=job%>) &nbsp;<%=jtype%></td>
					</tr>
					<tr>
						<td>Status:</td>
						<td>
							<table border=0 cellpadding="0" cellspacing="0" ID="Table3">
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
							<td>Total open tasks:</td>
							<td>
								<b>
								<%
								dim rst4,str
								Set rst4 = Server.CreateObject("ADODB.recordset")
								str="select count(*) as count1 from master_job_Tasks where jobid = " &jid& " and status <> 'complete' "
								rst4.Open str,cnn1
								'rst4.Open str, cnn1  , 0, 1, 1
								if trim(rst4("count1")) > 0 then %>
									<%=rst4("count1")%><%
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
					No tasks were found under the requested job id<br>&nbsp;
				</td>
			</tr>
			<%
		end if%>
	</table><%
end if 'caller=joblog 

if rst1.EOF then 

	if request("caller")<>"joblog" then%>
		<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#ffffff" ID="Table4">
			<tr>
				<td style="border-top:2px outset #ffffff;border-bottom:1px solid #cccccc;">&nbsp;No results found. Please try another search.<br>&nbsp;</td>
			</tr>
		</table><%
	end if 'caller=joblog

Else 'if not rst1.eof
	x=0
	%>
	<table border=0 cellpadding="3" cellspacing="1" width="100%" ID="Table5">
		<tr bgcolor="#dddddd" style="font-weight:bold;"> 
			<td width="10%">Due Date</td>	
			<td width="40%">Description</td>		
			<td width="10%" align="right">%Complete</td>  
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

	<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd" ><%
			While not rst1.EOF 
				%>
				<tr bgcolor="#ffffff" onMouseOver="highlight(this);" onMouseOut="unlight(this);"					
						onClick="javascript:openwin('edittasks.asp?mode=view&taskid=<%=rst1("taskid")%>',400,260)" style="cursor:hand"> 
					<td width="10%"><%=rst1("Duedate")%></td>	
					<td width="40%"><%=rst1("Description")%></td>					
					<td width="10%" align="right"><%=rst1("percentcomplete")*100%></td>					
					<td width="10%"><%=rst1("status")%></td>
					<input type="hidden" name="tid" value="<%=rst1("taskid")%>">
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
			<td colspan="5" bgcolor="#eeeeee"><i><%=x%> job tasks found</i></td>
		</tr>
	</table>
	<%
end if
rst1.close
%>

<%if request("caller")="joblog" then%>
	<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #cccccc;" >
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
