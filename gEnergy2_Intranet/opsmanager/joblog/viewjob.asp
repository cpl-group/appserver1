<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<%
dim cnn, rst, strsql, rstSuperChecker, nojob,rst666
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

dim jid,Desc , company ,job ,cStatus ,address1,address2,jtype,site_phone,fax_phone,projmanager,comppercent ,jobnotes,primarybilling,secondarybilling,primary_amt,secondary_amt, tcolor, customer, address_street,floor, citystatezip, custid,permissionflag, timb_ready, rfp, refjob, bldgnum, ticketcount, masterticketid,opentickets, totaltickets,taxstatusid,taxstatusdesc,taxcert,createdon,updatedby,updatetime,percenUpdatedby,nycompany,NYjid
dim contract , contract_xp, jobready 'added 4/1/2008 N.Ambo

jid = int(request("jid"))

'if allowGroups("Genergy_Corp,Joblog_Admin") then
	permissionflag=""
'else
	'permissionflag="1"
'end if

if trim(jid)<>"" then
	dim sqlopen
	sqlopen = "SELECT * FROM MASTER_JOB mj inner join taxstatuslist t on  t.id = mj.taxstatusid WHERE mj.id='"&jid&"'"
	'response.Write(sqlopen)

	rst.Open sqlopen , cnn 
	if not rst.EOF then
		bldgnum = rst("bldgnum")	
		jobnotes = rst("job_notes")		
		comppercent = rst("percent_complete")		
		Desc = rst("description")
		company = rst("company")
		nycompany = split(rst("job"),"-")(0) 'due to that one instance of a different int type then 00 for NY jobs

		rfp = rst("rfp")
		refjob = rst("referencejob")
		job = rst("job")
		jtype= rst("type")
		cStatus = rst("status")
		Select Case cStatus
			case "In progress"
				tcolor = "#66ff66"
			case "Unstarted"
				tcolor = "#ffcc00"
			case "Closed"
				tcolor = "#cc0033"
		end select 
		address1 = rst("address_1")
		address2 = rst("address_2")
		floor = rst("floor")
		if floor="Floor" then   ' Screen for default value
			floor=""
		else
			if trim(address2)<>"" and trim(floor)<>"" then
				floor="<br>"&floor
      		end if
		end if
		citystatezip = rst("city") & ", " & rst("state") & " " & rst("zip_code")
		site_phone = rst("site_phone")
		fax_phone = rst("fax_phone")
		
		set rst666 = server.createobject("ADODB.recordset")
		rst666.Open "select Firstname,lastname from managers where mid='"&rst("project_manager")& "'" , cnn
		if not rst666.EOF then
		projmanager = rst666("Firstname") & " " & rst666("lastname")
		end if
		rst666.close
		
		
		primarybilling = rst("billing_method_1")
		secondarybilling = rst("billing_method_2")
		primary_amt = rst("amt_1")
		secondary_amt = rst("amt_2")  
		customer= rst("customer_name")
		custid = rst("customer")
		taxstatusid = rst("taxstatusid")
		taxstatusdesc = rst("taxdesc")
		taxcert = rst("taxcert")
	    updatetime= rst("updatetime")
		percenUpdatedby=rst("percentupdatedby")
		contract = rst("contract")
		contract_xp = rst("contract_expdate")
		jobready = rst("readytoclose")
		
		
		
	    dim cnnando, rstando,sqlando,fname,sqlando2,stmt
		set cnnando = server.createobject("ADODB.connection")
		set rstando = server.createobject("ADODB.recordset")
		
	    sqlando = "select distinct fullname  from  master_job mjob inner join ["& application("Coreip")&"].dbcore.dbo.ADusers_GenergyUsers aduser on mjob.[user] = aduser.username  where username ='" & rst("user") &"'"
		cnnando = getConnect(0,0,"intranet")
		rstando.Open sqlando ,cnnando
		if not rstando.eof  then fname = rstando("fullname")
	
		if rst("user")<> "" and fname <>"" then stmt = " By: " & fname else stmt = ""
		
		CreatedOn = rst("Actual_Start_Date")  & stmt
		rstando.close
		if updatetime <> "" then updatetime = " On "  & rst("updatetime") else updatetime  = "" 
		
		percenUpdatedby = userFullname(percenUpdatedby)
		if percenUpdatedby = "N/A"  then  percenUpdatedby = "N/A"
		 
		 updatedby = userFullname(rst("updatedby"))
		 if updatedby <> "N/A" then updatedby = updatedby else updatedby="N/A" 
		
  dim cmd,AmountBilled,prm,rsss,jtdAmount,openpo,sql,hours
  set rsss = server.createobject("ADODB.Recordset")
  set cmd = server.createobject("ADODB.Command")
  cmd.CommandText = "sp_job_cost"
  cmd.CommandType = adCmdStoredProc
  Set prm = cmd.CreateParameter("c", adchar, adParamInput,2)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("j", advarchar, adParamInput,9)
  cmd.Parameters.Append prm
  cmd.Name = "test"
  Set cmd.ActiveConnection = cnn
	
  cnn.test company,job,rsss
 if company = "NY" then
   if len(jid)=4 then
	NYjid="00" & jid

	else
		NYjid="0" & jid
	end if
end if
  if not rsss.eof then
  AmountBilled=formatcurrency(rsss("JTD_work_billed"),2)
  jtdAmount=formatcurrency(rsss("jtd_cost"),2)
  rsss.close
  if company="GY" or company="GE" then
  'get sum of hours from time sheet
  rsss.open "SELECT isnull(sum(hours)+sum(overt),0) FROM times WHERE jobno = '"&jid &"'",cnn
  hours=rsss(0)
  rsss.close
  sql="select distinct isnull(a.amount-a.amount_invoiced,0),a.commitment from "&company&"_master_po_item a,"&company&"_master_po b where b.closed=0 and a.job='"&company&"-"&jid&"' and a.commitment=b.commitment order by a.commitment"
  else
  
  'sql="select isnull(sum(a.amount)-sum(a.amount_invoiced),0) from "&c&"_master_po_item a,"&c&"_master_po b where b.closed=0 and a.job='"&j&"' and  a.commitment=b.commitment"
  sql="select distinct isnull(a.amount-a.amount_invoiced,0),a.commitment from "&company&"_master_po_item a,"&company&"_master_po b where b.closed=0 and a.job='"&nycompany&"-"&NYjid&"' and  a.commitment=b.commitment order by a.commitment"

  rsss.open "SELECT isnull(sum(hours)+sum(overt),0) FROM times WHERE jobno = '"&jid&"'",cnn
  
  hours=rsss(0)
  rsss.close
  
  end if
  

  rsss.open sql,cnn



  while not rsss.EOF 
openpo = openpo + rsss(0)
'response.write "----------------)" &openpo &"<BR>"
rsss.movenext
wend


 
  rsss.close
  end if
		 
	else 
		nojob=true
	end if	'end if not rst.EOF then
	rst.close
end if

dim ticket
set ticket = New tickets
ticket.Label="Job"
ticket.Note = "Job Log Master Ticket for Job "&jid
ticket.ccuid  = ""
ticket.client = 1
ticket.requester = "JOBLOGADMIN"
ticket.department = "OPERATIONS"
ticket.userid = "JOBLOGADMIN" '750
if jid<>"0" and jid<>"" and isnumeric(jid) then ticket.findtickets "joblog", clng(jid)
'findtickets cint(jid), ticketcount, masterticketid,opentickets, totaltickets
 

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
  window.document.all['genjobtable'].bgColor ="#cccccc"
	<%
	if (allowgroups("Genergy_Corp,Joblog_Admin,gAccounting")) then 
		%>openwin(theURL,780,525)
		//alert("testing!")
		<%	
		else 
		%>openwin(theURL, 780, 400)<%
		end if
	%>
		
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function newPO(){
  document.location="/um/opslog/newpo.asp?jid=<%=jid%>&caller=joblog"
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

</script>
<link rel="Stylesheet" href="/styles.css" type="text/css">    
</head>

<body bgcolor="#eeeeee">
<%if nojob then
	response.write "<center>Job Number Not Found</center>"
else
%>
<form name="form2" method="post" action="">
<table border=0 cellpadding="3" cellspacing="0" width="100%" id="genjobtable" bgcolor="#eeeeee" style="border-top:1px solid #cccccc;border-right:1px solid #cccccc;">
	<tr> 
		<td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee">
			<table border="0" cellspacing="0" cellpadding="3" width="100%">
				<tr>
					
            <td> <b>General Info</b> |&nbsp;<a href="/gEnergy2_Intranet/opsmanager/joblog/jobestimates.asp?jid=<%=jid%>&job=<%=job%>&c=<%=company%>&mode=view">Job 
              Estimates</a>&nbsp;|&nbsp; <a href="<%="/gEnergy2_Intranet/opsmanager/joblog/jobtime.asp?jid="&jid&"&sbu=0"%>">Job 
              Time</a> &nbsp;|&nbsp; <a href="<%="/gEnergy2_Intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job 
              Folder</a>&nbsp;|&nbsp; <a href="/um/opslog/posearch.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Requisition 
              Forms</a>&nbsp;|&nbsp; <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=job%>&jid=<%=jid%>&caller=joblog">Job 
              Cost</a>&nbsp;|&nbsp; <a href="viewchange.asp?jid=<%=jid%>">Change 
              Orders</a>&nbsp;|&nbsp; <a href="/genergy2_Intranet/opsmanager/joblog/jobtasks.asp?caller=joblog&select=jobnum&findvar=<%=jid%>">Job
              Tasks</a></td>
  </tr>
  </table>
  </td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;" bgcolor="#eeeeee"><div id="backbutton2"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></div></td>
</tr>
<tr> 
  <td colspan="2" style="border-top:1px solid #ffffff;">  <table border=0 cellspacing="0" cellpadding="3" width="100%">
    <tr>
      <td width="170">Job:</td>
      <td colspan="4"><b><%=Desc%>&nbsp; (<%=company%>: &nbsp;<%=job%>) &nbsp;</b></td>
    </tr>
    <tr>
      <td>Job Type:</td>
      <td colspan="4"><%=jtype%></td>
    </tr>
    <tr>
      <td>Status:</td>
      <td colspan="4">
        <table border=0 cellpadding="0" cellspacing="0">
          <tr>
            <td><div style="position:inline;width:18px;height:12px;background:<%=tcolor%>;border:1px solid #999999;">&nbsp;</div></td>
            <td width="6">&nbsp;</td>
            <td><%=cStatus%> </td>
            <td width=18 align="center" valign="middle">&nbsp;</td>
            <td width=18 align="center" valign="middle"><%ticket.Display 0,true, true, false%></td>
          </tr>
      </table></td>
    </tr>
    <tr valign="top">
      <td width="170">Customer:</td>
      <td colspan="4"><a href="<%="javascript:customerdetail('" & custid & "','" & company &"')"%>"><%=customer%></a></td>
    </tr>
    <tr valign="top">
      <td rowspan="10">Address:</td>
      <td width="409"><%=address1%> </td>
	 <td width="226" style="background-Color:white;border-left:1px solid black;border-top:1px solid black;">Project Manager:</td>
      <td width="234" style="background-Color:white;border-right:1px solid black;border-top:1px solid black;">&nbsp;<%=projmanager%></td>
      <td width="145">&nbsp;</td>
    </tr>
    <tr valign="top">
      <td><%=address2&floor%></td>
      <td style="background-Color:white;border-left:1px solid black;">% Complete:</td>
      <td style="background-Color:white;border-right:1px solid black;">
      <table width="200" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="<%=comppercent%>%" bgcolor="<%if cint(comppercent) > "0" then%>#00FF00<%else%>#999999<%end if%>"><div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=comppercent%>%</font></div></td>
            <td width="<%=100 - comppercent%>%" bgcolor="<%if cint(comppercent) = "100" then%>#00FF00<%else%>#999999<%end if%>"></td>
          </tr>
      </table></td>     
      <td>&nbsp;</td>
    </tr>
    <tr valign="top">
      <td>&nbsp;</td>
      <td style="background-Color:white;border-left:1px solid black;">Job ready to be closed?:</td>
      <td style="background-Color:white;border-right:1px solid black;"><%if jobready then%>Yes<%else%>No<%end if%></td>
      <td>&nbsp;</td>
    </tr>
    <tr valign="top">
      <td>&nbsp;</td>
      <td style="background-Color:white;border-left:1px solid black;">% Updated Last By:</td>
      <td style="background-Color:white;border-right:1px solid black;"><%=percenUpdatedby &" " &updatetime%></td>
      <td>&nbsp;</td>
    </tr>
    <tr valign="top">
      <td><%=citystatezip%></td>
      <td style="background-Color:white;border-left:1px solid black;">Primary Billing:</td>
      <td style="background-Color:white;border-right:1px solid black;"><%=primarybilling%> (<%=formatcurrency(primary_amt,2)%>)</td>
      <td>&nbsp;</td>
    </tr>
    <tr valign="top">
      <td>&nbsp;</td>
     
    <td style="background-Color:white;border-left:1px solid black;">Secondary Billing:</td>
      <td style="background-Color:white;border-right:1px solid black;"><%=secondarybilling%> (<%=formatcurrency(secondary_amt,2)%>)</td>
    
	<tr valign="top">
      <td>&nbsp;</td>
       <td style="background-Color:white;border-left:1px solid black;">Amount Billed:</td>
      <td style="background-Color:white;border-right:1px solid black;"><%=AmountBilled%></td>
	<tr valign="top">
	  <td>&nbsp;</td>
	  <td style="background-Color:white;border-left:1px solid black;">Total Job Cost : </td>
	  <td style="background-Color:white;border-right:1px solid black;border-right:1px solid black;"><%= formatcurrency(jtdAmount+openpo,2)%></td>
	<%if contract then%>
	<tr valign="top">
	  <td>&nbsp;</td>
	  <td style="background-Color:white;border-left:1px solid black;">Contract Expiration Date: </td>
	  <td style="background-Color:white;border-right:1px solid black;border-right:1px solid black;"><%=contract_xp%></td>
	  <%end if%>
	<tr valign="top">
      <td>&nbsp;</td>
      <td style="background-Color:white;border-left:1px solid black;border-bottom:1px solid black;">Created On: </td>
      <td style="background-Color:white;border-right:1px solid black;border-right:1px solid black;border-bottom:1px solid black;"><%=CreatedOn%></td>
    <tr valign="top">
      <td>&nbsp;</td>
    <tr valign="top">
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td align="right"><a href="#" onClick="window.open('./statusrequest.asp?pm=<%=projmanager%>&jid=<%=jid%>&desc=<%=Desc%>','RequestUpdate','width=300,height=50, scrollbars=no')">Request Status Update</a></td>
      <td width="1">&nbsp;</td>
    </tr>
    <tr valign="top">
      <td> Tax Status </td>
      <td colspan="4"><%=taxstatusdesc%>
          <% if cint(taxstatusid) = 1 or cint(taxstatusid) = 3 then %>
          ;
          <%if taxcert then%>
      Certificate Received
      <%else%>
      Certificate Not Yet Received
      <%end if%>
      <%end if%>
      </td>
    </tr>
    <%
  if rfp = "True" then
  	%>
    <tr valign = "top">
      <td colspan=5>This job is marked as an RFP.</td>
    </tr>
    <%
  end if
  %>
    <%if trim(refjob) <> "" then%>
    <tr valign = "top">
      <td>Reference job number:</td>
      <td colspan="4"><%=refjob%></td>
    </tr>
    <%end if%>
    <%if trim(bldgnum) <> "" and trim(bldgnum) <> "0" then%>
    <tr>
      <td> This job is for: </td>
      <td colspan="4">
        <%
				dim rstOneBldg
				set rstOneBldg = server.createobject("ADODB.recordset")
				rstOneBldg.open "select b.strt as name from buildings b where bldgnum = '" _
					& bldgnum & "' order by b.strt", getConnect(0,0,"dbCore")
				if not rstOneBldg.eof then
					response.write rstOneBldg("name") & "  (" & bldgnum & ")"
				end if
				%>
        <input type="hidden" name="bldgnum" value="<%=bldgnum%>">
        <%
				rstOneBldg.close
				set rstOneBldg = nothing
				%>
      </td>
    </tr>
    <%end if%>
  </table></td>
</tr>
<tr>
		<td style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;" colspan="2"> 
			<!-- #include virtual="/includes/jobfooterbuttons.asp" -->
		</td>
</tr>
</table>
</form>
<% end if %>
</body>
</html>
<script>
	if (history.length == 0) {
		document.getElementById("backbutton1").innerHTML = "&nbsp;"
		document.getElementById("backbutton2").innerHTML = "&nbsp;"
	}
</script>