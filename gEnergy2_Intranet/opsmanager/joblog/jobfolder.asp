<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn, rst, strsql
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

dim jid,Desc , company ,job ,cStatus ,address,jtype,site_phone,fax_phone,projmanager,comppercent ,jobnotes,primarybilling,secondarybilling,primary_amt,secondary_amt, tcolor, customer, address_street, citystatezip, custid,total,tot1,permissionflag

jid = secureRequest("jid")

if session("corp")=5 then
  permissionflag=""
else
  permissionflag="1"
end if

	rst.Open "SELECT * FROM MASTER_JOB WHERE id='"&jid&"'", cnn
	if not rst.EOF then
		Desc = rst("description")
		company = rst("company")
		job = rst("job")
    jtype=left(rst("type"),6)
		cStatus = rst("status")
		Select Case cStatus
      case "In progress"
        tcolor = "#66ff66"
      case "Unstarted"
        tcolor = "#ffcc00"
      case "Closed"
        tcolor = "#cc0033"
		end select 
		address = rst("address_1")
		address_street = rst("address_2")
		citystatezip = rst("city") & ", " & rst("state") & " " & rst("zip_code")
		site_phone = rst("site_phone")
		fax_phone = rst("fax_phone")
		projmanager = rst("pm_first") & " " & rst("pm_last")
		comppercent = rst("percent_complete")
		jobnotes = rst("job_notes")
		primarybilling = rst("billing_method_1")
		secondarybilling = rst("billing_method_2")
		primary_amt = rst("amt_1")
		secondary_amt = rst("amt_2")	
		customer= rst("customer_name")
		custid = rst("customer")
	end if
	rst.close
	if trim(company)="IL" or trim(company)="03" then
	strsql = "Select * from ilite.dbo.times where jobno='" & jid & "' order by date desc"
	else
	strsql = "Select * from times where jobno='" & jid & "' order by date desc"
	end if

	'rst.open strsql,cnn    * Not sure what this does- rst not referenced below

%>
<html>
<head>
<title>View Job</title>
<script language="JavaScript1.2">
function customerdetail(cid,company) {
	theURL="cis_detail.asp?cid=" + cid + "&company=" + company
	openwin(theURL,600,475)
}
function edit_job(jid) {
	theURL="updatejob.asp?jid=" + jid
	//window.document.all['genjobtable'].Border ="1"
	window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,750,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function newTask(jid){	
	theURL=	"/genergy2_intranet/opsmanager/joblog/edittasks.asp?mode=new&jobnum=" + jid	
	//window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,400,230)
	}
	
function jobfolder(job){
	var jobid = new String(job)
	var dir = "data" + jobid.substr(0,1)
	var temp = "/um/opslog/gfile.asp?clientname=Job "+job+"&jobno=" + job
	document.all.jobfolder.src = temp;
}
function uploadfile(job){
	var jobid = new String(job)
	var dir = "data" + jobid.substr(0,1)
	var temp = "/um/opslog/gfile.aspx?clientname=Job "+job+"&jobno=" + job
	document.all.jobfolder.src = temp;
}
function newPO(){
  document.location="/um/opslog/newpo.asp?jid=<%=jid%>&caller=joblog"
}

function viewinvoice(jid) {
  theURL="/um/opslog/timesheetmain.asp?flag=0&job=" + jid
  //window.document.all['genjobtable'].Border ="1"
  window.document.all['genjobtable'].bgColor ="#999999"
  openwin(theURL,750,550)
}

var loaded = 0;
function preloadImages(){
  path = "/um/opslog/images/";
  edit_jobOn = new Image(); edit_jobOn.src = path + "btn-edit_job-1.gif";
  edit_jobOff = new Image(); edit_jobOff.src = path + "btn-edit_job.gif";
  new_poOn = new Image(); new_poOn.src = path + "btn-new_po-1.gif";
  new_poOff = new Image(); new_poOff.src = path + "btn-new_po.gif";
  invoiceOn = new Image(); invoiceOn.src = path + "btn-invoice-1.gif";
  invoiceOff = new Image(); invoiceOff.src = path + "btn-invoice.gif";
  
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
<link rel="Stylesheet" href="../../styles.css" type="text/css">   
</head>

<body bgcolor="#eeeeee" onload="jobfolder('<%=cint(jid)%>')">
<form name="form2" method="post" action="">
<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee" id="genjobtable" style="border-top:1px solid #cccccc;">
<tr> 
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
  <table border="0" cellspacing="0" cellpadding="3" width="100%">
  <tr>
            <td> <a href="<%="viewjob.asp?jid=" & jid %>">General Info</a> &nbsp;|&nbsp;<a href="/gEnergy2_Intranet/opsmanager/joblog/jobestimates.asp?jid=<%=jid%>&job=<%=job%>&c=<%=company%>&mode=view">Job Estimates</a> &nbsp;|&nbsp; <a href="<%="jobtime.asp?jid="&jid%>">Job 
              Time</a> &nbsp;|&nbsp; <b>Job Folder</b>&nbsp;|&nbsp; <a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requisition 
              Forms</a>&nbsp;|&nbsp; <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=job%>&jid=<%=jid%>&caller=joblog">Job 
              Cost</a>&nbsp;|&nbsp; <a href=viewchange.asp?jid=<%=jid%>>Change 
              Orders</a> &nbsp;|&nbsp; <a href="/genergy2_Intranet/opsmanager/joblog/jobtasks.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Job
              Tasks</a> &nbsp;|&nbsp; <a href="/um/opslog/gfile.aspx?caller=joblog&select=jobnum&findvar=<%=jid%>">Upload File</a></td>
  </tr>
  </table>
  </td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;" bgcolor="#eeeeee"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
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
</table>
<iframe src="null.htm" name="jobfolder" width="100%" height="70%" marginwidth="0" marginheight="0" border=0 frameborder=0></iframe> 
<% set cnn = nothing %>

<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #cccccc;">
<tr>
		<td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;"> 
			<!-- #include virtual="/includes/jobfooterbuttons.asp" -->
		</td>
  <td align="right" style="border-top:1px solid #ffffff;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>

</form>
</body>
</html>