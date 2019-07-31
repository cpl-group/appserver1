<%  '12/20/2007 N.Ambo made changes to allow only members of group 'Job Status Admins' to change the status of 'an 
'“in progress” or “unstarted” job to“closed”, or to change the status of a “closed” job to “in progress” or “unstarted'
'Addded a check box for 'contract on file', if checked then user must enter an expiration date
'1/7/2007 N.Ambo made change to show only project managers who are active
 %>
<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<% dim super,jobadmin

super = allowgroups("Genergy_Corp")
jobadmin = allowgroups("Joblog_Admin")
'super = true
%>
<html>
<head>
<title>Update Job</title>
<script>
function openwin(url,mwidth,mheight){
  window.name="opener";
  popwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth);
  popwin.focus();
}
function jobPicked(job){
	//N.ambo 4/4/2008 added this if statement to accomodate multiple job numbers
		if (document.form2.refjob.value == "") {
			document.form2.refjob.value = job;
		}
		else {	
			document.form2.refjob.value = document.form2.refjob.value+"; "+job;
		}
}
function updateCustomer(customerid, custname){
  document.form2.custid.value = customerid
  document.all.customername.innerHTML = custname
  document.form2.customer.value= custname	
}
	
function setDesc(job, id, name){
 	// the only reason this method is here is for the "open job list" window, which likes to call it sometimes.
	// it doesnt do anything, its just here so that an error is not generated.
	// dont want to scare the locals.
}

function closepage(action)
{
  switch (action)
  {
  case "close":
    opener.window.document.all['genjobtable'].bgColor ="#eeeeee"
    break;
  case "cancel":
    if (confirm("Cancel changes?")){
      opener.window.document.all['genjobtable'].bgColor ="#eeeeee"
      window.close()
    }
    break;
  case "save":
    alert("Update saved")
    opener.window.document.all['genjobtable'].bgColor ="#eeeeee"
    opener.document.location.reload()
    window.close()
    break;
  default:
    break;
  }
}
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">   
</head>
<%
dim cnn, rst, strsql,rst1,flagSave,flagReload, cmd, jstatus
' determine action
if cint(request("formsubmitted"))=2 then
  flagSave=True
  flagReload=True
else
  flagSave=False
  if request("formsubmitted")="" then
    flagReload=False
  else
    flagReload=True
  end if
end if

set cnn = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
if not flagSave then ' on save rst is not needed 
	set rst = server.createobject("ADODB.recordset")
	set rst1 = server.createobject("ADODB.recordset")
	%>
	<script>
	function getaddress(dd) {
	  document.form2.formsubmitted.value="1"
	  document.form2.submit()
	}
	function checkform(supervisor, jstatus) {
		if (jstatus != "closed"){
			if (supervisor=="True") {
				if (document.form2.jtypeid.selectedIndex==0 || document.form2.custid.value=='-1' || document.form2.taxexempt.selectedIndex==0 )
					alert("The Job Type, the Primary Customer, and the Tax Exempt Status are all required.")
				//7/3/2008 N.Ambo added to force expiration date entry if contract is on file
				else if (document.form2.contract.checked==true && document.form2.contract_xp.value=="")	
						alert("Expiration date is required if there is a contract on file.	")		
				//7/3/2008 N.Ambo added to force amount greate than 100 to be entered if billing method is not T&M
				//else if (document.form2.primarybilling.value.substring(0,7)!="T and M" && document.form2.primary_amt.value < 100)
					//alert("Please enter estimate if no current amount (at least $100)")			
				else
					document.form2.submit();		
			}else{
				alert("You do not have permission to edit this job.\nPlease request permision from your network administrator for access to one of the following:\nGenergy_Corp,Joblog_Admin,gAccounting.");
			}
		} else if (supervisor=="True") {
			document.form2.submit();
		}
	}
	function BuildingPicked(bldg){
	document.form2.bldgnum.value = bldg
	}

	//4/1/2008 N.Ambo added to show/remove date field based on whether contract is checked. If checked then show date field.
	function showdate() {	
		document.form2.formsubmitted.value = "1"
		document.form2.submit()			
	}
	</script>
	<%
end if
cnn.open getConnect(0,0,"intranet")

dim jid,Desc , company ,job ,cStatus ,address,site_phone,fax_phone,projmanager,comppercent ,jobnotes,primarybilling,secondarybilling,primary_amt,secondary_amt, tcolor, customer, address_street,address_2, jcity,jstate,jzip,jfloor,custid,projid,jtype,jtypeid,probability,timb_ready,pm_first, pm_last, rfp,refjob, bldgnum,taxstatusid,taxstatusdesc,taxcert,sqlopen,tm
dim contract, contract_xp, jobready '4/1/2008 N.Ambo added

'Response.write("allowgroups: " & allowgroups("Genergy_Corp,Joblog_Admin"))

jid = request("jid")

if flagReload then ' get fields from POST
    desc = left(trim(request("desc")),30)
    rfp = request("rfp")
    tm = request("tm")
	bldgnum = request("bldgnum")
	if rfp then
		rfp = 1
	else
		rfp = 0
	end if
    if tm then
		tm = 1
	else
		tm = 0
	end if
		
	cStatus = request("cstatus")
	company = request("company")
    customer= request("customer")
    custid = request("custid") 
    job = request("job")
	jtype = request("jtype")
    jtypeid = request("jtypeid")
	
	if jtypeid <> "36" then 
		bldgnum = 0
	end if 
	
	refjob = request("refjob")
	
    Select Case cStatus
      case "In progress"
        tcolor = "#33FF00"
      case "Unstarted"
        tcolor = "#FFFF00"
      case "Closed"
        tcolor = "#990000"
    end select 
	if not flagSave then
	  rst.open "select Address_1,Address_2,City,State,ZIP_Code from "&company&"_master_arm_customer where customer='" & custid & "'",cnn
	  address_street = rst("Address_1")
	  address_2 = rst("Address_2")
	  jcity = rst("City") 
	  jstate = rst("State") 
	  jzip  = rst("ZIP_Code")
	  rst.close
	else
      address_street = request("address_street")
	  address_2 = request("Address_2")
      jcity = request("jcity") 
      jstate = request("jstate") 
      jzip  = request("jzip")
	end if
	jfloor=request("jfloor")
    site_phone = request("site_phone")
    fax_phone = request("fax_phone")
    projid = request("projid")
    comppercent = request("comppercent")
    jobnotes = request("jobnotes")   
	primarybilling = request("primarybilling")	
	

	'4/1/2008 N.Ambo added if condition
	if request("primary_amt")="" then
		primary_amt=0
	else
		primary_amt = request("primary_amt")
	end if
	
	probability=request("probability")
	taxstatusid = request("taxexempt")
	taxcert	= request("taxcert")
		dim oldpercen,newpercen
		newpercen = request("comppercent")
		oldpercen=request("oldcomppercent")	
		
		
		
		
 if taxcert = "" then taxcert = 0 end if 
	if probability="" then
	  probability=0
	end if
	
	'4/1/2008 N.Ambo added if condition
	
	
	contract = request("contract")
	contract_xp = request("contract_xp")
	jobready = request("jobready")
	
	if flagSave then
'4/1/2008 N.Ambo added contract and contract_expdate to statements
'if contract is checked then expiration date is required and saved
      if newpercen <> oldpercen then 
		if contract then
			strsql = "UPDATE MASTER_JOB set description='"&desc&"', address_1='"&address_street&"', address_2='"&address_2&"', floor='"&jfloor&"', city='"&jcity&"', state='"&jstate&"', zip_code='"&jzip&"', site_phone='"&site_phone&"', fax_phone='"&fax_phone&"', project_manager='"&projid&"', status='" & cstatus & "', percent_complete='"&comppercent&"', job_notes='"&jobnotes&"', billing_method_1='"&primarybilling&"', customer='"&custid&"', customer_name='"&customer&"', amt_1=" & primary_amt & ", probability=" & probability & ", billing_method_2='', amt_2=0, type='"& jtype &"', type_id="& cint(trim(jtypeid)) & ", rfp="& trim(rfp) &", referencejob='" & trim (refjob) & "', bldgnum = '" & bldgnum & "', taxstatusid= "& cint(trim(taxstatusid)) & ", taxcert= "& cint(trim(taxcert))&", updatedby= '"& getxmlusername() & "', percentupdatedby= '"& getxmlusername() & "', updatetime= '"& date &" "& time &  "', contract= '"&contract& "', contract_expdate= '" &contract_xp& "', readytoclose= '" &jobready& "', tm= '" &tm& "' WHERE id=" &  jid
		else
			strsql = "UPDATE MASTER_JOB set description='"&desc&"', address_1='"&address_street&"', address_2='"&address_2&"', floor='"&jfloor&"', city='"&jcity&"', state='"&jstate&"', zip_code='"&jzip&"', site_phone='"&site_phone&"', fax_phone='"&fax_phone&"', project_manager='"&projid&"', status='" & cstatus & "', percent_complete='"&comppercent&"', job_notes='"&jobnotes&"', billing_method_1='"&primarybilling&"', customer='"&custid&"', customer_name='"&customer&"', amt_1=" & primary_amt & ", probability=" & probability & ", billing_method_2='', amt_2=0, type='"& jtype &"', type_id="& cint(trim(jtypeid)) & ", rfp="& trim(rfp) &", referencejob='" & trim (refjob) & "', bldgnum = '" & bldgnum & "', taxstatusid= "& cint(trim(taxstatusid)) & ", taxcert= "& cint(trim(taxcert))&", updatedby= '"& getxmlusername() & "', percentupdatedby= '"& getxmlusername() & "', updatetime= '"& date &" "& time &  "', contract= '"&contract& "', readytoclose= '" &jobready& "', tm= '" &tm& "' WHERE id=" &  jid
		end if
	  else 
	   if contract then
			strsql = "UPDATE MASTER_JOB set description='"&desc&"', address_1='"&address_street&"', address_2='"&address_2&"', floor='"&jfloor&"', city='"&jcity&"', state='"&jstate&"', zip_code='"&jzip&"', site_phone='"&site_phone&"', fax_phone='"&fax_phone&"', project_manager='"&projid&"', status='" & cstatus & "', percent_complete='"&comppercent&"', job_notes='"&jobnotes&"', billing_method_1='"&primarybilling&"', customer='"&custid&"', customer_name='"&customer&"', amt_1=" & primary_amt & ", probability=" & probability & ", billing_method_2='', amt_2=0, type='"& jtype &"', type_id="& cint(trim(jtypeid)) & ", rfp="& trim(rfp) &", referencejob='" & trim (refjob) & "', bldgnum = '" & bldgnum & "', taxstatusid= "& cint(trim(taxstatusid)) & ", taxcert= "& cint(trim(taxcert))&", updatedby= '"& getxmlusername() & "', contract= '"&contract& "', contract_expdate= '" &contract_xp& "', readytoclose= '" &jobready& "', tm= '" &tm& "' WHERE id=" &  jid
	   else
			strsql = "UPDATE MASTER_JOB set description='"&desc&"', address_1='"&address_street&"', address_2='"&address_2&"', floor='"&jfloor&"', city='"&jcity&"', state='"&jstate&"', zip_code='"&jzip&"', site_phone='"&site_phone&"', fax_phone='"&fax_phone&"', project_manager='"&projid&"', status='" & cstatus & "', percent_complete='"&comppercent&"', job_notes='"&jobnotes&"', billing_method_1='"&primarybilling&"', customer='"&custid&"', customer_name='"&customer&"', amt_1=" & primary_amt & ", probability=" & probability & ", billing_method_2='', amt_2=0, type='"& jtype &"', type_id="& cint(trim(jtypeid)) & ", rfp="& trim(rfp) &", referencejob='" & trim (refjob) & "', bldgnum = '" & bldgnum & "', taxstatusid= "& cint(trim(taxstatusid)) & ", taxcert= "& cint(trim(taxcert))&", updatedby= '"& getxmlusername() & "', contract= '"&contract& "', readytoclose= '" &jobready& "', tm= '" &tm& "' WHERE id=" &  jid
	   end if
	 end if
	
	 cnn.Execute strsql
	 '10/24/2007 N.Ambo remove automatic emails sent to GESCO managers for job types 'GE'
		'if trim(request("statusupdate")) = "1" then
			'dim prm
			'cmd.ActiveConnection = cnn
			'cmd.CommandText = "GE_STATUS_EMAIL"
			'cmd.CommandType = adCmdStoredProc
			'Set prm = cmd.CreateParameter("id", adInteger, adParamInput)
			'cmd.Parameters.Append prm
			'cmd.Parameters("id") = jid
			'cmd.execute
		'end if
      %>
  <script> closepage('save')</script>
      <%
	end if
	
elseif trim(jid)<>"" then  ' get fields from rst
 	sqlopen = "SELECT * FROM MASTER_JOB mj inner join taxstatuslist t on  t.id = mj.taxstatusid WHERE mj.id='"&jid&"'"
  	rst.Open sqlopen,cnn
 
  if not rst.EOF then
  	refjob = rst("referencejob") 
  	jobnotes = rst("job_notes")
  	probability=rst("probability")

  	bldgnum = rst("bldgnum")
'	timb_ready = rst("timb_ready")
	'response.write("TIMB"& timb_ready)
  	
    'response.write("JOB_NOTES:" & rst("job_notes") )
	
	Desc = rst("description")
    company = rst("company")
    job = rst("job")
	rfp = rst("rfp")
    tm = rst("tm")
	
	
    cStatus = rst("status")
    jstatus=lcase(cstatus)
	Select Case cStatus
		case "In progress"
			tcolor = "#33FF00"
		case "Unstarted"
			tcolor = "#FFFF00"
		case "Closed"
			tcolor = "#990000"
	end select 
    address_street = rst("address_1")
	address_2 = rst("address_2")
	jfloor=rst("floor")
    jcity = rst("city") 
    jstate = rst("state") 
    jzip  = rst("zip_code")
    site_phone = rst("site_phone")
    fax_phone = rst("fax_phone")
    projmanager = rst("pm_first") & " " & rst("pm_last")
    projid = rst("project_manager")
    comppercent = rst("percent_complete")
    
	'response.write("<br>JOBNOTES:" & jobnotes )
    primarybilling = rst("billing_method_1")
    
    primary_amt = rst("amt_1")
    
    customer= rst("customer_name")
    custid = rst("customer")
	jtype = rst("type")
	jtypeid = rst("type_id")
	taxstatusid = rst("taxstatusid")
	taxstatusdesc = rst("taxdesc")
	taxcert = rst("taxcert")
	contract = rst("contract")
	contract_xp = rst("contract_expdate")
	jobready = rst("readytoclose")
   dim stmt,all,CreatedOn ,updatedby
		
		dim cnnando, rstando,sqlando,fname
		set cnnando = server.createobject("ADODB.connection")
		set rstando = server.createobject("ADODB.recordset")
		
	sqlando = "select distinct fullname  from  master_job mjob inner join ["& application("Coreip")&"].dbcore.dbo.ADusers_GenergyUsers aduser on mjob.[user] = aduser.username  where username ='" & rst("user") &"'"
		cnnando = getConnect(0,0,"intranet")
		rstando.Open sqlando ,cnnando
		if not rstando.eof  then fname = rstando("fullname")
		if rst("user")="" then stmt = "" else stmt = " By: " & fname
		CreatedOn = rst("Actual_Start_Date")  & stmt
		all= " ("&company & " : " & job &")"&"&nbsp" & "Created On: " & CreatedOn 
         rstando.close
  end if
  rst.close
end if  ' end field loading
if not flagSave then ' display form
%>

<body bgcolor="#dddddd" onunload="closepage('close')">
<form name="form2" method="post" action="updatejob.asp">
	<input name="formsubmitted" type="hidden" value="2">
	<input name="company" type="hidden" value="<%=company%>">
	<input name="job" type="hidden" value="<%=job%>">
	
  <table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#6699cc"> 
      <td colspan="4"><span class="standardheader">Update Job<%=all%> </span></td>
    </tr>
    <tr> 
      <td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">All 
        fields are required except as noted<br></td>
    </tr>
    <tr> 
      <td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table width="100%" border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td width="96">Description</td>
            <td width="490"> <input name="jid" type="hidden" value="<%=jid%>"> 
              <%if (super or jobadmin) and jstatus <>"closed" then %>
              <input name="desc" type="text" value="<%=Desc%> " size="30" maxlength="30"> 
              &nbsp;<!--(<%'=company%> : <%'=job%>) -->
              <%else
								%>
              <input name="desc" type="hidden" value="<%=Desc%>"> 
              <%
								'Response.write(Desc & "   " & company & " : " & job)
							end if%>
            </td>
            <td width="111" style="border-top:2px solid black;border-left:2px solid black;">Tax Exempt Status</td>
            <td width="75"  style="border-top:2px solid black;border-right:2px solid black;" nowrap>
			<%if jstatus <> "closed" then%><select name="taxexempt">
                <option value=""><font face="Arial, Helvetica, sans-serif">Select 
                Tax Status</font></option>
                <%
									strsql = "SELECT * from taxstatuslist where companycode = 'all' or companycode='"&company&"'"
									rst.Open strsql, cnn
									if not rst.eof then
										do until rst.eof
												%>
                <option value="<%=rst("ID")%>" <%if trim(taxstatusid)=trim(rst("id")) then%>selected<%end if%>> 
                <font face="Arial, Helvetica, sans-serif"><%=left(rst("taxdesc"),30)%></font> 
                </option>
                <%
											rst.movenext
										loop
									end if
									rst.close
									%>
              </select>
              <%else%>
              <%=taxstatusdesc%>
              <input name="taxexempt" type="hidden" value="<%=taxstatusid%>">
              <%end if%></td>
          </tr>
          <tr> 
            <td>Status</td>
            <td> 
              <select name="cStatus" onchange="statusUpdate.value=1">
                <%
				
									rst1.Open "select distinct status from status where job=1 order by status desc", cnn
									if not rst1.eof then
										do until rst1.eof
											if rst1("status")=cStatus then%>
                <option value="<%=rst1("status")%>" selected ><font face="Arial, Helvetica, sans-serif"><%=rst1("status")%></font></option>
                <%elseif rst1("status") ="Closed" then
					'12/20/2007 N.Ambo added this if statement so that only members of 'Job Status Admins' can close a job which is 'unstarted' or 'in progress'
					if allowgroups("Job Status Admins") then%>
					<option value="<%=rst1("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("status")%></font></option>
					<%end if%>
                <%elseif rst1("status") <> "Closed" and cStatus <> "Closed" then %>
                <option value="<%=rst1("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("status")%></font></option>
                <%elseif cStatus = "Closed" then
					'12/20/2007 N.Ambo added this if statement so that only members of 'Job Status Admins' can change back the job status if it is closed
					if allowgroups("Job Status Admins") then%>                
					<option value="<%=rst1("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("status")%></font></option>
					<%end if               
				end if 
				rst1.movenext
				loop
				end if
				rst1.close%>
              </select> 
              &nbsp;Job Type 
              <input name="jtype" type="hidden" value="<%=jtype%>"> 
              <%
							
							if (super or jobadmin) and jstatus <>"closed"  then 
								%>
              <select name="jtypeid" onChange="javascript:document.form2.jtype.value=this.options[this.selectedIndex].text;checkBldgSelect();">
                <option value=""><font face="Arial, Helvetica, sans-serif">Select 
                Job Type</font></option>
                <%
									
									strsql = "SELECT Type,Type_ID FROM master_job_types " & _
									"where job=1 and company='" &company& "' and enable =1 " & _
									"ORDER BY Type "
									
									rst.Open strsql, cnn
									if not rst.eof then
										do until rst.eof
											response.Write("<!--"&jtypeid&"-->")
											if trim(jtypeid)=trim(rst("Type_ID")) then  
												%>
                <option value="<%=rst("Type_ID")%>" selected> <font face="Arial, Helvetica, sans-serif"><%=rst("Type")%></font> 
                </option>
                <%
											else
												%>
                <option value="<%=rst("Type_ID")%>"> <font face="Arial, Helvetica, sans-serif"><%=rst("Type")%></font> 
                </option>
                <%
											end if
											rst.movenext
										loop
									end if
									rst.close
									%>
              </select> 
              <%
							else		' user is not a supervisor
								%>
              <input name="jtypeid" type = "hidden" value="<%=jtypeid%>"> 
              <%
								Response.write(jtype)
							end if
							%>
            </td>
            <td  style="border-bottom:2px solid black;border-left:2px solid black;">Certification Received</td>
            <td  style="border-bottom:2px solid black;border-right:2px solid black;"><input type="checkbox" name="taxcert" value="1" <%if taxcert then %>checked<%end if%> <%if (not super and not jobadmin) or jstatus = "closed" then%>disabled<%end if%>></td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td width="48%" colspan="2" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td width="80">Customer</td>
            <td> <input name="customer" type="hidden" value="<%=customer%>"> 
              <input name="custid" type = "hidden" value="<%=custid%>"> 
              <span id="customername"><%Response.write(customer)%></span>
            </td>
          </tr>
          <tr>
            <td></td>
            <td><%if (super or jobadmin) then%><a href="#" onclick="openwin('multicustomerlink.asp?jid=<%=jid%>&company=<%=company%>&jstatus=<%=jStatus%>&customer=<%=custid%>',550,300)">link customer(s) to this job</a><%end if%></td>
          </tr>
          <tr> 
            <td></td>
            <td><br> <b>Job Address</b><br><%if (super or jobadmin) and jstatus <>"closed" then %>
<input name="address_street" type="text" value="<%=address_street%>" size="40" maxlength="30"><%else %>
              <input name="address_street" type = "hidden" value="<%=address_street%>">
              <%=address_street%><%end if%> 
              <br>
              <%if (super or jobadmin) and jstatus <>"closed" then %>
              <input name="address_2" type="text" value="<%=address_2%>" size="40" maxlength="30">
              <%else %><input name="address_2" type = "hidden" value="<%=address_2%>">
<%=address_2%>
              <%end if%>
            </td>
          </tr>
          <tr> 
            <td align="right">Floor:</td>
            <td><%if (super or jobadmin) and jstatus <>"closed" then %><input name="jfloor" type="text" value="<%=jfloor%>" size="30" maxlength="15">
              <%else %><input name="jfloor" type = "hidden" value="<%=jfloor%>"><%=jfloor%>
              <%end if%>
            </td>
          </tr>
          <tr> 
            <td align="right">City:</td>
            <td><%if (super or jobadmin) and jstatus <>"closed" then %><input name="jcity" type="text" value="<%=jCity%>" size="30" maxlength="15">
              <%else %><input name="jcity" type = "hidden" value="<%=jCity%>"><%=jCity%>
              <%end if%>
            </td>
          </tr>
          <tr> 
            <td align="right">State:</td>
            <td><%if (super or jobadmin) and jstatus <>"closed" then %><input name="jstate" type="text" value="<%=jState%>" size="2" maxlength="4">
              <%else %><input name="jstate" type = "hidden" value="<%=jState%>"><%=jState%>
              <%end if%>
              &nbsp;&nbsp;Zip:&nbsp;
              <%if (super or jobadmin) and jstatus <>"closed" then %>
              <input name="jzip" type="text" value="<%=jZip%>" size="8" maxlength="10">
              <%else %><input name="jzip" type = "hidden" value="<%=jZip%>"><%=jZip%>
              <%end if%>
            </td>
          </tr>
        </table></td>
      <td width="52%" bgcolor="#eeeeee" colspan ="2" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr bgcolor="#eeeeee"> 
            <td>Project Manager</td>
            <td colspan = 3> 
              <%if (super or jobadmin) and jstatus <>"closed" then%>
              <select name="projid">
                <%
                                '1/7/2007 N.Ambo amended
								rst1.Open "select * from Managers where Active is null and companycode = '"&trim(company)&"'order by lastname, firstname", cnn
								do until rst1.eof%>
                <option value="<%=rst1("mid")%>" <%If trim(projid)=trim(rst1("mid")) then%>selected<%end if%>><%=rst1("lastname")%>, 
                <%=rst1("firstname")%></option>
                <%
									rst1.movenext
								loop
								rst1.close
								%>
              </select> 
              <%
							else		'supervisor is fals
								%>
              <input name="projid" type = "hidden" value = "<%=projid%>"> 
              <%
								Response.write(projmanager)
							end if
							%>
            </td>
          </tr>
          <tr bgcolor="#eeeeee"> 
            <td>% Complete</td>
            <td colspan = 3> 
              <%
							if jstatus <>"closed"  then 
								%>
              <input name="comppercent" type="text" value="<%=comppercent%>" size="3" maxlength="40">
              %
                   <input name = "oldcomppercent" type = "hidden" value="<%=comppercent%>">
			  <%
							else		'Job Not Closed
								%>
              <input name = "comppercent" type = "hidden" value="<%=comppercent%>">
              <%
								response.write(comppercent & " %")
							end if
							%>
			<input type="checkbox" name="jobready" value="1" <%if jobready then %><%="CHECKED"%><%end if%> >
              Job ready to be closed  
            </td>
          </tr>
          <tr bgcolor="eeeeee"> 
            <td> % Probability </td>
            <td> 
              <%
							if (super or jobadmin) and jstatus <>"closed"  then 
								%>
              <input name="probability" type="text" value="<%=probability%>" size="3" maxlength="3">
              % 
              <%
							else		' supervisor is false
								%>
              <input name="probability" type="hidden" value="<%=probability%>"> 
              <%=probability%> % 
              <%
							end if
							%>
            </td>
          </tr>
          <tr bgcolor="#eeeeee"> 
            <td>Income Class</td>
            <td colspan = 3> 
              <%
							if (super or jobadmin) and jstatus <>"closed"  then 
								%>
              <select name="primarybilling">
                <%
									rst1.Open "select jobtype from jobtypes where companycode = '"&trim(company)&"'", cnn
									do until rst1.eof 
										if trim(rst1("jobtype")) = trim(primarybilling) then
											%>
                <option value="<%=rst1("jobtype")%>" selected><%=rst1("jobtype")%></option>
                <%
										else
											%>
                <option value="<%=rst1("jobtype")%>"><%=rst1("jobtype")%></option>
                <%
										end if
										rst1.movenext
									loop
									rst1.close
									%>
              </select> 
              <%
							else		'supervisor is false
								%>
              <input name="primarybilling" type="hidden" value="<%=trim(primarybilling)%>">
              <%
								Response.write( trim(primarybilling) )
							end if%>
            </td>
          </tr>
            <tr>
            <td>
              T&M Job            
            </td>
             <td> 
              <%
								if (super or jobadmin) and jstatus <>"closed"  then 
									%>
              <input type="checkbox" name = "tm" value = 1 <%if tm then%>checked<%end if%>>
              <%
								else
									%>
              <input name="tm" type = "hidden" value ="<%=tm%>">
              <%
								end if
								%>
            </td>

          </tr>
          <tr bgcolor="#eeeeee"> 
            <td>Amount ($)</td>
            <td> 
              <%if (super or jobadmin) and jstatus <>"closed"  then %>
              <input name="primary_amt" type="text" value="<%=primary_amt%>" size="17" maxlength="40"> 
              <%else%>
              <input name ="primary_amt" type = "hidden" value = "<%=primary_amt%>"> 
              <%=primary_amt%> 
              <%end if%>
            </td>
          </tr>
          <tr bgcolor="#eeeeee"> 
            
          <tr> 
            <td> 
              <%if (super or jobadmin) and jstatus <>"closed"  then %>
              RFP: 
              <%elseif rfp then%>
              Job is an RFP. 
              <%end if%>
            </td>
            <td> 
              <%
								if (super or jobadmin) and jstatus <>"closed"  then 
									%>
              <input type="checkbox" name = "rfp" value = 1 <%if rfp then%>checked<%end if%>>
              <%
								else
									%>
              <input name="rfp" type = "hidden" value ="<%=rfp%>">
              <%
								end if
								%>
            </td>
          </tr>
           <tr> 
           <td><input type="checkbox" name="contract" value="1" onclick="showdate();"<%if contract then %>checked<%end if%>>
              Contract on File</td>
           <% if contract then %>
           <td>&nbsp;&nbsp;Expiration Date: <input name="contract_xp" type="text" value="<%=contract_xp%>" ></td>
           <%end if%>
          </tr>
          <%if ucase(trim(company)) <> "GY"  then%>
          
		  <!------commenting the whole timb_ready part>
		  <%'if (super or jobadmin) and jstatus <>"closed"  then %>
          <tr> 
            <td>Ready For Accounting Transfer:</td>
            <td> <input name="timb_ready" type="checkbox" value="1" <%'if trim(timb_ready) = "True" then%>checked<%'end if%>> 
            </td>
          </tr>
          <%'else%>
          <tr> 
            <td colspan = 2> 
              <%
			'							if timb_ready = "True" then
			'								response.write("This job is ready for transfer to accounting.")
			'							else
			'								response.write("This job is NOT ready for transfer to accounting.")
			'							end if
			'							%>
            </td>
          </tr-->
          <%'end if%>
          <%end if%>
        </table>
        <br> </td>
    </tr>
    <%if (super or jobadmin) and jstatus <>"closed"  then  '%>
    <script>
					function checkBldgSelect(){
						var id = document.form2.jtypeid.value
						if (parseInt(id) == 36){
							document.getElementById('placeholder1').style.display='none';
							document.getElementById('placeholder2').style.display='none';
							document.getElementById('bldgSelect1').style.display='block';
							document.getElementById('bldgSelect2').style.display='block';
							
							
						}else{
							document.getElementById('placeholder1').style.display='block';
							document.getElementById('placeholder2').style.display='block';
							document.getElementById('bldgSelect1').style.display='none';				
							document.getElementById('bldgSelect2').style.display='none';
							document.form2.bldgnum.value = "N/A"
						}
					}
				</script>
    <%
				dim pholder, bs
				pholder="block"
				bs="none"
				if jtypeid="36" then
					pholder="none"
					bs="block"
				end if
				%>
    <tr> 
      <td id="placeholder1" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;display:<%=pholder%>;">&nbsp;</td>
      <td id="placeholder2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;display:<%=pholder%>;">&nbsp;</td>
      <td id="bldgselect1" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;display:<%=bs%>;"> 
        This job is for Building: </td>
      <td id="bldgselect2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;display:<%=bs%>;"> 
        <input name="bldgnum" type="text" value="<%=trim(lcase(bldgnum))%>" size="4"> 
        <img src="/genergy2/setup/images/aro-rt.gif" border="0">&nbsp;<a href="javascript:openwin('/genergy2/setup/quickbuildinglist.asp?name=<%=request("name")%>',260,330);">Quick 
        Building search</a> </td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">Reference 
        Job Number:</td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <input name = "refjob" type = "text" value = "<%=refjob%>" size = "30"> 
        <img src="/genergy2/setup/images/aro-rt.gif" border="0">&nbsp; <a href="javascript:openwin('/um/opslog/timesheet-beta/joblist.asp?name=<%=request("name")%>',260,320);">Quick 
        job search</a>      
        <table><span style="font-size:7pt;color:#999999;">(separate multiple jobs with a semicolon)</span></table> <%'5/8/2008 Nambo added this comment to be show with reference # field'%>
       </td>
    </tr>
    <%else%>
    <tr> 
      <%if trim(bldgnum) <> "" then%>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        This job is for: </td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
        <%
							dim rstOneBldg
							set rstOneBldg = server.createobject("ADODB.recordset")
							rstOneBldg.open "select b.strt as name from buildings b where bldgnum = '" _
								& bldgnum & "' order by b.strt", getConnect(0,bldgnum,"billing")
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
      <% else %>
      <td colspan="2"  style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">&nbsp;</td>
      <%end if		
					dim refLabel, refOutput
					if trim(refjob) <> "" then
						refLabel = "Reference Job Number:"
						refOutput = refJob
					end if%>
      <td width="20%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=refLabel%>&nbsp;</td>
      <td align ="left" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=refOutput%>&nbsp;</td>
      <input name = "refjob" type = "hidden" value = "<%=refjob%>">
    </tr>
    <%end if%>
    <td> </td>
    </tr>
    <tr bgcolor="#dddddd"> 
      <td colspan="4"> 
        <!--<div style="margin-left:89px;">-->
        <input type="button" value="Update" onClick="checkform('True','<%=jstatus%>')">
        &nbsp; <input type="button" value="Cancel" onclick="closepage('cancel');"><%if not super then%>&nbsp;&nbsp;For updating rights ask Net Admin for access to one of the following Genergy_Corp,Joblog_Admin, or gAccounting.<%end if%>
      </td>
    </tr>
  </table>
	<input type="hidden" name="statusUpdate" value="0">
</form>
</body>
</html>
<%end if ''
%>