<%option explicit%>
<% 'N.Ambo 5/30/2008 modified to add date range values to url string when 'print frame' is selected%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim cnn, rst, strsql, fromTime, toTime, andWhere

set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

dim jid,Desc , company ,job ,cStatus ,address,jtype,site_phone,fax_phone,projmanager,comppercent ,jobnotes,primarybilling,secondarybilling,primary_amt,secondary_amt, tcolor, customer, address_street, citystatezip, custid,total,tot1,permissionflag

'**** 10/07/08
dim rstUser, Uname, flagUser, sUser
						
if not secureRequest("UsrName")="" then
	Uname=secureRequest("UsrName")
	if not Trim(Uname)="All" then
		flagUser=1
	end if
end if
'****

dim sbu  'sort by user - a boolean value
sbu = 0

dim printFlag	'printable view, a boolean value


jid 		= secureRequest("jid")
fromTime 	= secureRequest("fromTime")
toTime 		= secureRequest("toTime")
sbu			= secureRequest("sbu")
printFlag	= secureRequest("printview")
if secureRequest("printview") = "" then
	printFlag = "false"
end if
	
if isnumeric(sbu) then sbu = cint(sbu) else	sbu = 0

'N.Ambo added 5/30/2008 to retrieve date range values when button 'Print Frame' is selected
if printFlag = "true" then
	fromTime = request.QueryString("fromTime")
	toTime= request.QueryString("toTime")
end if

'if session("corp")=5 then
  permissionflag=""
'else
 ' permissionflag="1"
'end if


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
  
if trim(fromTime) <> "" then 
	andWhere = " and date between '" & fromTime & "' and '" & toTime & "'"
else
	andWhere = ""
end if
  
if trim(company)="IL" or trim(company)="NY" then
    '10/01/2008 by Kamto Cheng - updated query to convert matricola into lower case
	strsql = "Select lower(matricola) AS matricola, [date], description, hours, overT from times where right(jobno,4)='" & jid & "'" & andWhere 
else
	strsql = "Select lower(matricola) AS matricola, [date], description, hours, overT from times where jobno='" & jid & "'" & andWhere
end if
  
'**** 10/07/08
if flagUser=1 then
	strsql = strsql & " And RTrim(LTrim(matricola))='" & Trim(Uname) & "' "
end if
'****
  
  
if sbu = 1 then
	strsql = strsql & " order by matricola asc, date desc"
else
	strsql = strsql & " order by date desc"
end if

    
'response.write strsql
'response.end
rst.open strsql,cnn



%>
<html>
<head>
<title>View Job<%
if printFlag = "true" then
	%> ID - <%=jid%> - <%=Desc%><%end if%></title>
<script language="JavaScript1.2">
function displayTimeSheet(jid, sbu, from, toT, uname) {
	document.form2.sbu.value = 0
	document.location.href="jobtime.asp?jid=" + jid + "&sbu=" + sbu + "&fromTime=" + from + "&toTime=" + toT + "&UsrName=" + uname
}

function displayUserTimeSheet(jid, from, toT, uname) {
	document.form2.sbu.value = 1
	document.location.href="jobtime.asp?jid=" + jid + "&sbu=1&fromTime=" + from + "&toTime=" + toT + "&UsrName=" + uname
}

function displayDateTimeSheet(jid, from, toT) {
	document.form2.sbu.value = 0
	document.location.href="jobtime.asp?jid=" + jid + "&sbu=0&fromTime=" + from + "&toTime=" + toT + "&UsrName=" + uname
}

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
function viewinvoice(jid) {
  theURL="/um/opslog/timesheetmain.asp?flag=0&job=" + jid
  //window.document.all['genjobtable'].Border ="1"
  window.document.all['genjobtable'].bgColor ="#999999"
  openwin(theURL,750,550)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function newTask(jid){	
	theURL=	"/genergy2_intranet/opsmanager/joblog/edittasks.asp?mode=new&jobnum=" + jid	
	//window.document.all['genjobtable'].bgColor ="#999999"
	openwin(theURL,400,230)
	}
	
function newPO(){
  document.location="/um/opslog/newpo.asp?jid=<%=jid%>&caller=joblog"
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
function sortByUser(){
	var urllink="jobtime.asp";
	window.navigate(urllink);
}
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">   
</head>

<body bgcolor="#eeeeee" <%if printFlag = "true" then%> onLoad="window.print()" <%end if%>>

<form name="form2" method="post" action="jobtime.asp">
<% if printFlag = "false" then %>
<table border=0 cellpadding="3" cellspacing="0" id="genjobtable" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #cccccc;border-right:1px solid #cccccc;">
<tr> 
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
  <table border="0" cellspacing="0" cellpadding="3" width="100%">
  <tr>
            <td> <a href="<%="viewjob.asp?jid=" & jid %>">General Info</a> |&nbsp;<a href="/gEnergy2_Intranet/opsmanager/joblog/jobestimates.asp?jid=<%=jid%>&job=<%=job%>&c=<%=company%>&mode=view">Job 
              Estimates</a>&nbsp;|&nbsp; <b>Job Time</b> &nbsp;|&nbsp; <a href="<%="jobfolder.asp?jid="&jid%>">Job 
              Folder</a>&nbsp;|&nbsp; <a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requisition 
              Forms</a>&nbsp;|&nbsp; <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=job%>&jid=<%=jid%>&caller=joblog">Job 
              Cost</a>&nbsp;|&nbsp; <a href=viewchange.asp?jid=<%=jid%>>Change 
              Orders</a>&nbsp;|&nbsp;<a href="/genergy2_Intranet/opsmanager/joblog/jobtasks.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Job
              Tasks</a></td>
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
            <td> <table border=0 cellpadding="0" cellspacing="0">
                <tr> 
                  <td><div style="position:inline;width:18px;height:12px;background:<%=tcolor%>;border:1px solid #999999;">&nbsp;</div></td>
                  <td width="6">&nbsp;</td>
                  <td><%=cStatus%></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td valign="bottom">Time <%if trim(fromTime) <> "" then %><br>(<%=fromTime%> - <%=toTime%>)<%end if%>:</td>
            <td valign="bottom"><b><span id="totalhours"></span></b> &nbsp;hours total, <b><span id="totalover"></span></b> 
              &nbsp;hours overtime</td>
          </tr>
          <!--**** 10/07/08-->
          
          <tr> 
            <td valign="bottom">Select User:</td>
            <td valign="bottom">
				<select name="ddlUser" onChange="document.location='jobtime.asp?jid=<%=jid%>&UsrName='+this.value">
					<option value="All">All Users</option>
					<%Set rstUser = Server.CreateObject("ADODB.recordset")
						rstUser.Open "SELECT Distinct(RTrim(LTrim(matricola))) as [User] FROM times WHERE jobno='"& jid &"'", cnn
					do until rstUser.eof
					sUser=Split(rstUser("User"),"\")
					%>
					<option value="<%=rstUser("User")%>" <%if trim(Uname)=trim(rstUser("User")) then response.write " SELECTED"%>>
						<%=sUser(1)%>
					</option>
					<%
					rstUser.movenext
					loop
					rstUser.close%>
				</select>
		    </td>
          </tr>
          <!--****-->
          <tr>
            <td>&nbsp;</td>
            <td><p>
		<input name="jid" type="hidden" value="<%=jid%>">
		Adjust Timeframe to View: From Date - 
		<input type="text" name="fromTime" value="<%=fromTime%>">
		To Date - 
		<input type="text" name="toTime" value="<%=toTime%>">
		<input type="button" name="Submit" value="View" onClick="displayTimeSheet('<%=jid%>', '<%=sbu%>', document.form2.fromTime.value, document.form2.toTime.value, document.form2.ddlUser.value)">
		<!--input type="submit" name="Submit" value="View"-->
		<input type = "hidden" name= "sbu" value=<%=sbu%>>
		<input type="hidden" name="UsrName" value="<%=Uname%>">
		
		<input type = "hidden" name = "printview" value = "false">
		<!--**** 10/07/2008-->
		<%if flagUser = 1 then %>
			<%if sbu = 0 then %>
				<input type="submit" name="ByUser1" value="Sort By User" disabled=true>
			<%else%>
				<input type="submit" name="ByDate1" value="Sort By Date" disabled=true>
			<% end if %>	
		<%else %>
			<%if sbu = 0 then %>
				<input type="button" name="ByUser" value="Sort By User" onClick="displayUserTimeSheet('<%=jid%>', document.form2.fromTime.value, document.form2.toTime.value, document.form2.ddlUser.value)">
			<%else%>
				<input type="button" name="ByDate" value="Sort By Date" onClick="displayDateTimeSheet('<%=jid%>', document.form2.fromTime.value, document.form2.toTime.value, document.form2.ddlUser.value)">
			<%end if%>		
		<%end if%>
		
		<%if flagUser = 1 then %>
			<input type="button" name="PrintViewUser" value="Print Frame" onClick="openwin('jobtime.asp?jid=<%=jid%>&UsrName='+document.form2.ddlUser.value+'&sbu=<%=sbu%>&fromTime='+document.form2.fromTime.value+'&toTime='+document.form2.toTime.value+'&printview=true', 600, 600)">
		<%else %>
			<input type="button" name="PrintView1" value="Print Frame" onClick="openwin('jobtime.asp?jid=<%=jid%>&sbu=<%=sbu%>&fromTime='+document.form2.fromTime.value+'&toTime='+document.form2.toTime.value+'&printview=true', 600, 600)">
		<%end if%>
		<!--****-->
	   </p>
	   <p>&nbsp; </p></td>
          </tr>
        </table>
  </td>
</tr>
</table>
<% end if 		'this will be in every view %>
<table border=0 cellpadding="0" cellspacing="0" width="100%" bgcolor="#eeeeee" style="border-right:1px solid #cccccc;">
<tr>
 <td> 
 <table border=0 cellpadding="3" cellspacing="1" width="100%">
 <tr bgcolor="#dddddd"> 
 	<%if sbu = 1 then%>
		<td width="10%">User</td>
	<%end if%>
	<td width="10%">Date</td>
	<td width="60%">Description</td>
	<td width="10%">Hours</td>
	<td width="10%">Overtime</td>
	<%if sbu = 0 then%>
		<td width="12%">User</td>
	<%end if%>
 </tr>
 </table>
 
 <%if trim(session("login"))<>"luna" and printFlag = "false" then %>
 	<div style="width:100%; overflow:auto; height:300px;border-right:1px solid #cccccc;">
 <%end if%>  
 
 <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
 <%
 total = 0
  
 dim lastCola, thisCola, totalColaHours, totalColaOTHours
 if not rst.EOF then
  	thisCola = trim(rst("matricola"))
	lastCola = trim(rst("matricola"))
  end if
  
  while not rst.EOF 
  	thisCola = trim(rst("matricola"))
  	total = total + cdbl(rst("hours"))
  	tot1 = tot1 + cdbl(rst("overt"))
  	
	if ((sbu = 1) AND (thisCola <> lastCola)) then 'totals for user name%>
		<tr bgcolor="#eeeeee"> 
  		<td width="10%">&nbsp;</td>
		<td width="10%">&nbsp;</td>
    	<td width="60%" align="right"><b>Total hours for <%=lastCola%></b></td>
    	<td width="10%" align="right"><b><%=formatnumber(totalColaHours,2)%></b></td>
    	<td width="10%" align="right"><b><%=formatnumber(totalColaOTHours,2)%></b></td>
		</tr>
		</table>
		
		<%
		totalColaHours = 0
		totalColaOTHours = 0
		%>
		<br>
		<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
		<%
	end if
  %>
  	
  	<tr bgcolor="#ffffff"> 
  	<%if sbu = 1 then%>
  		<td width = "10%"><%=rst("matricola")%></td>
	<%end if%>	
    <td width="10%"><%=rst("date")%> </td>
    <td width="60%"><%=rst("description")%></td>
    <td width="10%" align="right"><%=formatnumber(rst("hours"),2)%></td>
    <td width="10%" align="right"><%=formatnumber(rst("overt"),2)%></td>
  	<%if sbu = 0 then%>
    	<td width = "10%"><%=rst("matricola")%></td>
  	<%end if%>
  	</tr>
	<%
	totalColaHours = cdbl(rst("hours")) + cdbl(totalColaHours)
	totalColaOTHours = cdbl(rst("overt")) + cdbl(totalColaOTHours)

 
  	lastCola = trim(rst("matricola"))
   	rst.movenext
  wend				' end of the loop, we have stepped through all time entries
  
  if (sbu = 1) then ' this is here so that the last user gets totals too
  					' since the colas never change for him
	%>
	<tr bgcolor="#eeeeee"> 
  	<td width="10%">&nbsp;</td>
	<td width="10%">&nbsp;</td>
    <td width="60%" align="right"><b>Total hours for <%=lastCola%></b></td>
    <td width="10%" align="right"><b><%=formatnumber(totalColaHours,2)%></b></td>
    <td width="10%" align="right"><b><%=formatnumber(totalColaOTHours,2)%></b></td>
	</tr>
	<%
	totalColaHours = 0
	totalColaOTHours = 0
  end if
  %>
  
  <tr bgcolor="#eeeeee">
    <td width="10%">&nbsp;</td>
	<%if sbu = 1 then%>
	<td width="10%">&nbsp;</td>
	<%end if%>
    <td width="60%" align="right"><b>Total hours for all users:</b></td>
    <td width="10%" align="right"><b><%=formatnumber(total,2)%></b></td>
    <td width="10%" align="right"><b><%=formatnumber(tot1,2)%></b></td>
    
	<%if sbu = 0 then %>
	<td width="10%">&nbsp;</td>
	<%end if%>
  </tr>
  </table>
  <%
  if printFlag = "false" then
  	%>
  	<script>
  	//<!--
  	document.all.totalhours.innerHTML = "<%=formatnumber(total,2)%>";
  	document.all.totalover.innerHTML = "<%=formatnumber(tot1,2)%>";
  	//-->
  	</script>
	<%
  end if
  %>
  </div>

  <%
  set cnn = nothing %>
  </td>
</tr>
</table>
<% if printFlag = "false" then %>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border-right:1px solid #cccccc;">
<tr>
		<td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;"> 
			<!-- #include virtual="/includes/jobfooterbuttons.asp" -->
		</td>
  <td align="right" style="border-top:2px outset #ffffff;border-bottom:1px solid #cccccc;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
<% end if %>
</form>
</body>
</html>






