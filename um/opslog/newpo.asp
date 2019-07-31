<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

Dim cnn1, rst
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

dim jid, fullname, sessionname, company, companyname
company = request("company")
jid = request("jid")
if Session("name")<>"" then
  sessionname = (Session("name"))
end if
rst.open "select name from companycodes where code='"&company&"' order by name", cnn1
if not rst.eof then
	companyname = rst("name")
end if
rst.close
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
//<!--
function newPO(){
  document.location="/um/opslog/newpo.asp?jid=<%=jid%>&caller=joblog"
}
function openwin(url,mwidth,mheight){
  window.name="opener";
  popwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth);
  popwin.focus();
}

	
function setDesc(job, id, name){
   	//doesnt do anything, but the job picker gets mad if its not here.
}

function jobPicked(job){
	document.form1.jobnum.value = job;
}

function newTask(jid){	
	theURL=	"/genergy2_intranet/opsmanager/joblog/edittasks.asp?mode=new&jobnum=" + jid	
	//window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,400,230)
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
function checkSubmit(){

 if (document.form1.jobnum.value == ""){
 	alert("Job Number is Required.")		
 }
 else {
 	document.form1.submit()
 }
}
//-->
</script>
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#eeeeee" text="#000000">
<form name="form1" method="post" action="savepo.asp">
<table border=0 cellpadding="3" cellspacing="0" width="100%" id="genjobtable" bgcolor="#eeeeee" style="border-top:1px solid #cccccc;">
<%if request("caller")="joblog" then
caller = "joblog"
Set rst2 = Server.CreateObject("ADODB.recordset")
Dim tcolor,Desc,job,jtype,cstatus
rst2.Open "SELECT description, company, job, type, status FROM MASTER_JOB WHERE id='"&jid&"'", cnn1
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
end if
rst2.close

%>
<tr bgcolor="#eeeeee"> 
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
  <table border="0" cellspacing="0" cellpadding="3" width="100%">
  <tr>
    <td>
    <a href="<%="/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=" & jid %>">General Info</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobtime.asp?jid=" &jid%>">Job Time</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job Folder</a> &nbsp;|&nbsp; <b><a href="/um/opslog/posearch.asp?caller=<%=caller%>&select=jobnum&findvar=<%=jid%>">Requisition Forms</a></b>&nbsp;|&nbsp; <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=job%>&jid=<%=jid%>">Job Cost</a>&nbsp;|&nbsp; <a href="/genergy2_Intranet/opsmanager/joblog/jobtasks.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Job
              Tasks</a>
    </td>
  </tr>
  </table>
  </td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;" bgcolor="#eeeeee"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
<tr bgcolor="#eeeeee"> 
  <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <table border="0" cellspacing="0" cellpadding="3">
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
  <tr>
    <td>Total purchase orders:</td>
    <td>
    <b><!--[[span id="totalPOs"]][[/span]]--> 
    <% Set rst3 = Server.CreateObject("ADODB.recordset")
    str="select sum(po_total) as sum1 from po where jobnum=" & jid & " and accepted = 1"
    rst3.Open str, cnn1, 0, 1, 1
    if rst3("sum1") > 0 then %>
      <%=Formatcurrency(rst3("sum1"))%>
      <% else%>
      $0 
    <% end if 
    rst3.close
    %>
    </b>
    </td>
  </tr>
  <tr>
    <td>Total hours invoiced:</td>
    <td>
    <b>
    <% Set rst4 = Server.CreateObject("ADODB.recordset")
    str="select sum(hours) as sum1 from invoice_submission where jobno=" & jid & " and submitted=1"
    rst4.Open str, cnn1, 0, 1, 1
    if trim(rst4("sum1")) > 0 then %>
      <%=rst4("sum1")%>
      <% else %>
      0    
    <% end if
    rst4.close 
    %>
    </b>
    </td>
  </tr>
  </table>
  </td>
</tr>
<%  end if %>
<tr> 
  <td colspan="2" bgcolor="#dddddd" style="border-bottom:1px solid #ffffff;"> 
    <b>New Requisition Form <%if trim(companyname)<>"" then%>(<%=companyname%>)<%end if%></b>
    <input type="hidden" name="job" value="<%=job%>">
  </td>
</tr>
<tr> 
  <td colspan="2" bgcolor="#eeeeee"> 
  <%
if company="" then
	%>Specify Company  
					<select name="company" onchange="document.location='newpo.asp?company='+this.value"><%
			        rst.Open "select * from companycodes where active = 1 order by name", getConnect(0,0,"intranet")
			        if not rst.eof then
			        do until rst.eof
				        %><option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><%=rst("name")%></option><%
				        rst.movenext
			        loop
			        end if
			        rst.close%>
				</select>

	</td></tr></table></body></html><%
	response.end
end if
%>

  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr valign="top">
    <td>
    <!-- begin first column -->
    <table border=0 cellpadding="3" cellspacing="0">
    <tr> 
      <td>Date (mm/dd/yyyy)</td>
      <td><input type="hidden" name="podate" value="<%=date()%>"><%=date()%></td>
    </tr>
    <tr>
      <td>Vendor</td>
      <td>
			<select name="vid"><%
			rst.Open "SELECT distinct vendor, name FROM " & company & "_MASTER_APM_VENDOR order by name", getConnect(0,0,"intranet")
			do until rst.eof
				%><option value="<%=rst("vendor")%>"><%=rst("name")%></option><%
				rst.movenext
			loop
			rst.close%>
			</select>
	  </td>
    </tr>
    <tr>
      <td>Requisitioner</td>
      <td>  
      <select name="req">
      <%Set rst3 = Server.CreateObject("ADODB.recordset")
      sqlstr = "select department, username as user1,FullName from ADusers_GenergyUsers order by department, FullName"
      rst3.Open sqlstr, getconnect(0,0,"dbCore"), 0, 1, 1
      if not rst3.eof then
		dim tracktype
		tracktype = ""
        do until rst3.eof 
        fullname = trim(rst3("FullName"))
		if tracktype = "" then
				%>
				<OPTGROUP Label="<%=rst3("Department")%>">
				<%
				elseif trim(tracktype) <> trim(rst3("department")) then 
				%>
				</OPTGROUP><OPTGROUP Label="<%=rst3("department")%> ">
				<% 
		end if 
		  tracktype = trim(rst3("department"))
		%> 
        <option value="<%=rst3("user1")%>" <% if trim(fullname) = trim(sessionname) then %> selected<% end if %>><%=rst3("FullName")%></option>
        <%
      rst3.movenext
      loop
      end if
      rst3.close
      %>
      </select>
      </td>
    </tr>
    <tr>
      <td>Job Address</td>
      <td><input type="text" name="jobaddr"></td>
    </tr>
    <tr>
      <td>Shipping Address</td>
      <td><input type="text" name="shipaddr"></td>
    </tr>
    <tr>
      <td>Job Number</td>
      <td>
      <% if trim(jid)<>"" then %>
      <%=jid%>
      <input type="hidden" name="jobnum" value="<%=jid%>">
      <% else %>
      <input name="jobnum" type = "text" size = "4">
	  <img src="/genergy2/setup/images/aro-rt.gif" border="0">&nbsp;<a href="javascript:openwin('/um/opslog/timesheet-beta/joblist.asp?name=<%=request("name")%>',260,320);">Quick job search</a>
	  
      <% end if %>
      </td>
    </tr>
    </table>
    <!-- end first column -->
    </td>
    <td>
    <!-- begin second column -->
    <p>RF Description<br><textarea name="description" cols="50" rows="3"></textarea></p>
    <!-- end second column -->
    </td>
  </tr>
  </table>
  </td>
</tr>
<tr> 
  <td colspan="2" bgcolor="#dddddd" style="border-top:1px solid #cccccc;"> 
  <input type="button" name="choice" value="Save" onclick="checkSubmit()">
  <input type="button" name="Button22" value="Cancel" onClick="Javascript:history.back()">
  <input type="hidden" name="caller" value="<%=caller%>">
  </td>
</tr>
</table>
<% if caller = "joblog" then %>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #cccccc;" bgcolor="#eeeeee">
<tr>
  <td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;">
  <table border=0 cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="/um/opslog/images/btn-edit_job.gif" value="Edit Job" align="middle" onclick="editjob('<%=jid%>');return false;" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">&nbsp;</td>
    <td><img src="/um/opslog/images/btn-invoice.gif" value="Invoice Job" id="invoice" align="middle" onclick="viewinvoice('<%=jid%>');" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">&nbsp;</td>
    <td><img src="/um/opslog/images/btn-new_po.gif" value="New Requisition Form" id="new_po" align="middle" onclick="newPO();" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">&nbsp;</td>
  </tr>
  </table> 
  </td>
  <td align="right" style="border-top:1px solid #ffffff;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
<% end if %>
</form>

</body>
</html>
