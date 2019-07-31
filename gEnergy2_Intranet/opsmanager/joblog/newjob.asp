<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim cnn, rst, strsql
dim jid,Desc , company ,job ,cStatus ,projmanager,comppercent ,jobnotes,primarybilling,secondarybilling,primary_amt,secondary_amt, tcolor, customer, address_street
dim jcity,jstate,jzip,jfloor, cust_name,projid, address_floor, jtype,jtypeid,probability,bldgnum,taxstatusid,taxstatusdesc,taxcert, super
dim refjob
dim contract, contract_xp, jobready '4/3/2008 N.Ambo added

super = allowgroups("Genergy_Corp,Joblog_Admin")

set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

select case request("mode")
  case "new"
  
  if request("company")<>"" then
    company=request("company")
  '10/19/2007 N.Ambo added this line to defualt the company to Genergy and go straight to the fill in screen
  else
	company = "" 'rsm added
  end if
  %>
    <html>
    <head>
    <title>New Job</title>
    <script>
	function openwin(url,mwidth,mheight){
  		window.name="opener";
  		popwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth);
  		popwin.focus();
	}
	
	function setDesc(job, id, name){
   	//doesnt do anything, but the job picker gets mad if its not here.
	}
	function jobPicked(job){
		//N.ambo 4/4/2008 added this if statement to accomodate multiple job numbers
		if (document.form1.refjob.value == "") {
			document.form1.refjob.value = job;
		}
		else {	
			document.form1.refjob.value = document.form1.refjob.value+";"+job;
		}
	}
	function BuildingPicked(bldg){
	document.form1.bldgnum.value = bldg
	}
  function checkpercent() {
  if(document.form1.comppercent.value>100) {
    alert('Enter percent up to 100')
    }
  else {
    if(document.form1.cStatus.selectedIndex==1 && document.form1.comppercent.value==0) {
       alert('Percent must be greater than zero for started jobs')
      }
    }
  
  }
  
  //4/1/2008 N.Ambo added this new function for mandatory requirements for % complete and primary amount billed
  function updateprobability() 
  {
     
     if(document.form1.cStatus.selectedIndex>0)
     {
        document.form1.probability.value = '100'  
     }
  }
  function updateFields() {
  //bill amounts are not mandatory for 'T&M' jobs
  if(document.form1.primarybilling.value.substring(0,7)=="T and M" && (document.form1.primary_amt.value == "ENTER AMOUNT" || document.form1.primary_amt.value=="") ) {
		document.form1.primary_amt.value = 'N/A'  
   }
  
   if (document.form1.primarybilling.value.substring(0,7)!="T and M" && document.form1.primary_amt.value == "N/A") {
		document.form1.primary_amt.value = "ENTER AMOUNT"
	}
	//% complete mandatory only for 'Fixed Contract' jobs
  if (document.form1.primarybilling.value.substring(0,14)!= "Fixed Contract" && document.form1.comppercent.value=="" ) {
		document.form1.comppercent.value = 'N/A'		
   }
   
   if (document.form1.primarybilling.value.substring(0,14)== "Fixed Contract" && document.form1.comppercent.value=="N/A" ) {
		document.form1.comppercent.value = ''		
   }
  
  }
  
  //4/1/2008 N.Ambo amended this function for mandatory requirements for % complete and primary amount billed
    function checkform()
    {
      //if (document.form1.desc.value == "" || document.form1.address_street.value == "" || document.form1.jcity.value == "" || document.form1.jstate.value == "" || document.form1.projid.selectedIndex ==0 || document.form1.taxexempt.selectedIndex ==0 || document.form1.primarybilling.selectedIndex ==0 || document.form1.jtypeid.selectedIndex ==0 || document.form1.primary_amt.value == "ENTER AMOUNT" || document.form1.primary_amt.value == "" || (document.form1.primarybilling.value.substring(0,14)!="Monthly Charge"&&document.form1.primary_amt.value < <%if trim(company) <> "GE" then %>100<%else%>0<%end if%>) || document.form1.customer.selectedIndex ==0 || (document.form1.jtype.value.indexOf("RFP")!=-1 && document.form1.probability.value==0)||document.form1.comppercent.value==''||((parseInt(document.form1.jtypeid.value) == 36) && (document.form1.bldgnum.value == "0" && document.form1.RFP.checked==false) )){
      if (document.form1.desc.value == "" || document.form1.address_street.value == "" || document.form1.jcity.value == "" || document.form1.jstate.value == "" || document.form1.projid.selectedIndex ==0 || document.form1.taxexempt.selectedIndex ==0 || document.form1.primarybilling.selectedIndex ==0 || document.form1.jtypeid.selectedIndex ==0 || (document.form1.primarybilling.value.substring(0,7)!="T and M"&& (document.form1.primary_amt.value < <%if trim(company) <> "GE" then %>100<%else%>0<%end if%> || document.form1.primary_amt.value == "ENTER AMOUNT")) || document.form1.customer.selectedIndex ==0 || (document.form1.jtype.value.indexOf("RFP")!=-1 && document.form1.probability.value==0) || (document.form1.primarybilling.value.substring(0,14) == "Fixed Contract" && document.form1.comppercent.value=='') || ((parseInt(document.form1.jtypeid.value) == 36) && (document.form1.bldgnum.value == "0" && document.form1.RFP.checked==false) ) || (document.form1.contract.checked==true && document.form1.contract_xp.value=="") ){
		if(document.form1.primarybilling.selectedIndex ==0) 
		  alert("Job must have a Income Class")
		else
			if( (parseInt(document.form1.jtypeid.value) == 36) && (document.form1.bldgnum.value == "0") && (document.form1.RFP.checked == false))
				alert("R&B jobs must have a building selected.")
		else
		  if(document.form1.jtype.value.indexOf("RFP")!=-1 && document.form1.probability.value==0)
		    alert("RFPs must have a probability of acceptance")
	      else
		    //if((document.form1.primarybilling.value.substring(0,14)!="Monthly Charge"&&document.form1.primary_amt.value < 100) || document.form1.primary_amt.value == "ENTER AMOUNT")
		   if( document.form1.cStatus.selectedIndex>0 && (document.form1.primary_amt.value<1 || document.form1.primary_amt.value == "ENTER AMOUNT"))
		      alert("Amount($) required for NON Pending jobs")		
           else 
              if(document.form1.cStatus.selectedIndex==0 && document.form1.probability.value==0)
		        alert("Pending job requires probability")    
		      else
                if(document.form1.cStatus.selectedIndex>0 && document.form1.probability.value<100)
		          alert("All NON Pending jobs require probability be set to 100")    
		        else
                  alert("FORM NOT COMPLETE. PLEASE MAKE SURE TO COMPLETE ALL REQUIRED FIELDS")
      }
	  else{
        var yesno=confirm("Are you sure you want to create a job for "+document.form1.jtype.value+"?")
        if(yesno) {
          document.form1.mode.value="save"
          document.form1.submit()
        }
      }
       
    
    }
function getaddress(dd, company) {
	document.form1.cust_name.value=dd.options[dd.selectedIndex].text
	if (company != "GE") {
		document.form1.mode.value="new"
		document.form1.submit()
	}
}

function screencompany(company) {
  document.location.href="newjob.asp?mode=new&company="+company 
}
function closepage(action)
{
  switch (action)
  {
  case "cancel":
    if (confirm("Cancel new job?")){
      window.close()
    }
    break;
  default:
    break;
  }
}

//4/1/2008 N.Ambo added to show/remove date field based on whether contract is checked. If checked then show date field.
function showdate() {	
		document.form1.submit()			
}
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">
</head>
<body bgcolor="#dddddd" onunload="closepage('close')">
<form name="form1" method="get" action="newjob.asp">
<input type="hidden" name="mode" value="new">
    
    
  <table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee">
    <tr bgcolor="#6699cc"> 
      <td colspan="4"> 
        <% if company = "" then %>
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td nowrap><span class="standardheader">Select the company you wish 
              to add a job for</span></td>
            <td> 
              <!--<select name="company" onchange="screencompany(this.value)">-->
              <select name="company" onchange="submit();">
                <%if trim(company) = "" then %>
                <option value="">Select Company</option>
                <% end if %>
                <%
        rst.Open "select * from companycodes where active = 1 order by name", cnn 'rsm
        if not rst.eof then
        do until rst.eof
        %>
                <option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><font face="Arial, Helvetica, sans-serif"><%=rst("name")%></font></option>
                <%    
        rst.movenext
        loop
        end if
        rst.close%>
              </select> 
              <%
        if request("customer")<>"" then
        customer=request("customer")
        cust_name=request("cust_name")
        jtype=request("jtype")
        jtypeid=request("jtypeid")
        desc=request("desc")
        cstatus=request("cstatus")
        else
        customer=""
        jtype=""
        jtypeid=""
        desc=""
        cstatus=""
        cust_name=""
        end if
        %>
            </td>
          </tr>
        </table>
        <% else %>
        <span class="standardheader">Add New Job</span> 
        <% end if %>
      </td>
    </tr>
    <% if company <> "" then %>
    <tr> 
      <td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">All 
        fields are required except as noted<br></td>
    </tr>
    <tr> 
      <td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td width="80">Company</td>
            <td> 
              <!--<select name="company" onchange="screencompany(this.value)">-->
              <select name="company" onchange="document.forms[0].elements.customer.selectedIndex = 0;submit();">
                <%if trim(company) = "" then %>
                <option value="">Select Company</option>
                <% end if %>
                <%
        rst.Open "select * from companycodes where active = 1 order by name", cnn 'rsm
        if not rst.eof then
        do until rst.eof
        %>
                <option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><font face="Arial, Helvetica, sans-serif"><%=rst("name")%></font></option>
                <%    
        rst.movenext
        loop
        end if
        rst.close%>
              </select> 
              <%
        if request("customer")<>"" then
        customer=request("customer")
        cust_name=request("cust_name")
        jtype=request("jtype")
        jtypeid=request("jtypeid")
        desc=request("desc")
        cstatus=request("cstatus")
        else
        customer=""
        jtype=""
        jtypeid=""
        desc=""
        cstatus=""
        cust_name=""
        end if
        %>
            </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <!-- wrapper table for formatting(keep items below in first col)-->
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td>Description</td>
            <td colspan =2><input name="desc" type="text" size="30" maxlength="30" value="<%=desc%>"></td>
          </tr>
          <tr> 
            <td width="80">Status</td>
            <td> <select name="cStatus" onchange="updateprobability();">
                <%
        rst.Open "select distinct status from status where job=1 and code <> 'c' order by status desc", cnn
        if not rst.eof then
        do until rst.eof
      if cstatus=rst("status") then
        %>
                <option value="<%=rst("status")%>" selected><font face="Arial, Helvetica, sans-serif"><%=rst("status")%></font></option>
                <%          
      else
        %>
                <option value="<%=rst("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst("status")%></font></option>
                <%    
      end if      
        rst.movenext
        loop
        end if
        rst.close%>
              </select> &nbsp;&nbsp;Job Type&nbsp; <input name="jtype" type="hidden" value="<%=jtype%>">             
			  <select name="jtypeid" onChange="javascript:jtype.value=this.options[this.selectedIndex].text;checkBldgSelect();">
                <option value="">Select Job Type</option>
                <%
        strsql = "SELECT Type,Type_ID FROM master_job_types where job=1 and enable=1 and company='"&company&"' ORDER BY Type"
          rst.Open strsql, cnn
        if not rst.eof then
          do until rst.eof
      if jtypeid=trim(rst("Type_ID")) then  
        %>
                <option value="<%=rst("Type_ID")%>" selected><font face="Arial, Helvetica, sans-serif"><%=rst("Type")%></font></option>
                <%
      else
        %>
                <option value="<%=rst("Type_ID")%>"><font face="Arial, Helvetica, sans-serif"><%=rst("Type")%></font></option>
                <%
      end if
        rst.movenext
        loop
      
        end if
        rst.close
        %>
              </select> </td>
            <td><input type="checkbox" name="RFP"<%if lcase(request("RFP"))="on" or (trim(jtype)="" and lcase(company)="ge") then %><%="CHECKED"%><%end if%>>
              RFP Job</td>
            <!-- end wrapper-->
          </tr>
          <tr>
            <td>Tax Status</td>
            <td><select name="taxexempt">
                <option value=""><font face="Arial, Helvetica, sans-serif">Select 
                Tax Status</font></option>
                <%
									strsql = "SELECT * from taxstatuslist where companycode = 'all' or companycode='"&company&"'"
									rst.Open strsql, cnn
									if not rst.eof then
										do until rst.eof
												%>
                <option value="<%=rst("ID")%>" <%if trim(taxstatusid)=trim(rst("id")) then%>selected<%end if%>> 
                <font face="Arial, Helvetica, sans-serif"><%=left(rst("taxdesc"),50)%></font> 
                </option>
                <%
											rst.movenext
										loop
									end if
									rst.close
									%>
              </select>
              &nbsp;<%if not super then%>Tax Certification <input type="checkbox" name="taxcert" value="1"><%end if%></td>
            <td>&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td width="48%" colspan="2" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td width="80">Customer</td>
            <td> 
              <!-- Hold onto customer name as well as ID -->
              <input type="hidden" name="cust_name" value="<%=cust_name%>"> <select name="customer" onChange="getaddress(this, '<%=trim(company)%>')">
                <option value="00-0000" selected>Select Customer</option>
                <%
    if company<>"" then
      rst.Open "SELECT distinct customer,name, status FROM " & company & "_MASTER_ARM_CUSTOMER order by  name", cnn
          do until rst.eof
		  cstatus = lcase(trim(rst("status")))
      if trim(rst("customer"))=customer then
          		%>
                <option value="<%=trim(rst("customer"))%>" selected><%=left(trim(rst("name")),30)%></option>
                <%
      else
	  	if cstatus = "inactive" then 
          		%>
                <OPTGROUP Label="<%=left(trim(rst("name")),30)%> (inactive)"></OPTGROUP>
                <%
		else
          		%>
                <option value="<%=trim(rst("customer"))%>"><%=left(trim(rst("name")),30)%></option>
                <%
		end if
      end if
          rst.movenext
          loop
      rst.close
    end if
        %>
              </select> </td>
          </tr>
          <%
    ' if user has selected a customer, fill in address fields
    if customer <> "" and customer <> "00-0000" and company <> "GE" then
		rst.open "select Address_1,Address_2,City,State,ZIP_Code from "&company&"_master_arm_customer where customer='" & customer & "'"
		address_street = rst("Address_1")
		address_floor = rst("Address_2")
		jcity = rst("City") 
		jstate = rst("State") 
		jzip  = rst("ZIP_Code")
    rst.close
    else
    address_street = ""
      address_floor = ""
      jcity = "" 
      jstate = "" 
      jzip  = ""
    end if
    %>
          <tr> 
            <td></td>
            <td><br> <b>Address</b>&nbsp;<span style="font-size:7pt;color:#999999;">(second 
              line optional)</span><br> <input name="address_street" type="text" value="<%=address_street%>" size="40" maxlength="30"></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><input name="address_floor" type="text" value="<%=address_floor%>" size="40" maxlength="30"></td>
          </tr>
          <tr> 
            <td align="right">Floor:</td>
            <td><input name="jfloor" type="text"  size="30" maxlength="15"> </td>
          </tr>
          <tr> 
            <td align="right">City:</td>
            <td><input name="jcity" type="text" value="<%=jcity%>" size="30" maxlength="15"></td>
          </tr>
          <tr> 
            <td align="right">State:</td>
            <td><input name="jstate" type="text" value="<%=jstate%>" size="2" maxlength="4"> 
              &nbsp;&nbsp;Zip:&nbsp; <input name="jzip" value="<%=jzip%>" type="text" size="8" maxlength="10"></td>
          </tr>
        </table></td>
      <td width="52%" bgcolor="#eeeeee" colspan="2" style="border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
            <td>Project Manager:</td>
            <td> 
              <% if trim(company) <> "" then %>
              <select name="projid">
                <option value="none" selected>Select Project Manager</option>
                <%
        rst.Open "select * from Managers where Active is null and companycode = '"&trim(company)&"'order by lastname, firstname", cnn
        do until rst.eof%>
                <option value="<%=rst("mid")%>" <%If trim(request("projid"))=trim(rst("mid")) then%>selected<%end if%>><%=rst("lastname")%>, 
                <%=rst("firstname")%></option>
                <%
        rst.movenext
        loop
        rst.close
        %>
              </select> 
              <%end if%>
            </td>
          </tr>
          <tr> 
            <td>% Complete</td>
            <td><input name="comppercent" type="text" size="3" maxlength="30" value="<%=request("comppercent")%>" onBlur="checkpercent()">
              % &nbsp;&nbsp;<span style="font-size:7pt;color:#999999;">(0-100)</span> 
              <input type="checkbox" name="jobready" value="1" <%if request("jobready")="1" then %><%="CHECKED"%><%end if%>>
              Job ready to be closed             
            </td>
          </tr>
          <tr> 
            <td>Income Class</td>
            <td> <select name="primarybilling" onchange="updateFields();">
                <option value="none" selected>Select Income Class</option>
                <%
			rst.Open "select jobtype from jobtypes where companycode = '"&trim(company)&"'order by jobtype", cnn
			do until rst.eof 
				if trim(rst("jobtype")) = trim(request("primarybilling")) then
					%>
                <option value="<%=rst("jobtype")%>" selected><%=rst("jobtype")%></option>
                <%
				else
					%>
                <option value="<%=rst("jobtype")%>"><%=rst("jobtype")%></option>
                <%
				end if
				rst.movenext
			loop
			rst.close
			%>
              </select> </td>
          </tr>
          <tr>
            <td><input type="checkbox" name="tm" value="1" <%if request("tm")="1" then %><%="CHECKED"%><%end if%>>
              T&M Job            
            </td>


          </tr>
          <tr> 
            <td>Amount ($)</td>
            <td> 
              <%dim initAmt
			if trim(request("primary_amt")) = "" then
				initAmt = "ENTER AMOUNT"
			else
				initAmt = trim(request("primary_amt"))
			end if%>
              <input name="primary_amt" type="text" value="<%=initAmt%>" size="17" maxlength="40" onclick="form1.primary_amt.value=''"> 
              &nbsp;% Probability&nbsp; <input name="probability" type="text" value="<%=request("probability")%>" size="3" maxlength="3">
              <span style="font-size:7pt;color:#999999;">(0-100)</span> </td>
          </tr>
           <tr> 
           <td><input type="checkbox" name="contract" value="1" onclick="showdate();"<%if request("contract")="1" then %><%="CHECKED"%><%end if%>>
              Contract on File</td>
           <% if request("contract")=1 then %>
           <td>&nbsp;&nbsp;Expiration Date: <input name="contract_xp" type="text" value="<%=request("contract_expdate")%>"></td>
           <%end if%>
          </tr>
        </table>
        <br> </td>
    </tr>
    <tr> 
      <script>
		function checkBldgSelect(){
			var id = document.form1.jtypeid.value
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
			}
		}
		</script>
      <%
			dim pholder, bs
			pholder="block"
			bs="none"
			if request("jtypeid")="36" then
				pholder="none"
				bs="block"
			end if
			%>
      <td id="placeholder1" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;display:<%=pholder%>;">&nbsp;</td>
      <td id="placeholder2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;display:<%=pholder%>;">&nbsp;</td>
      <td id="bldgSelect1" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;display:<%=bs%>;"> 
        Job for Building: </td>
      <td id="bldgSelect2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;display:<%=bs%>;"> 
        <input name="bldgnum" type="text" value="<%=trim(lcase(bldgnum))%>" size="4"> 
        <img src="/genergy2/setup/images/aro-rt.gif" border="0">&nbsp;<a href="javascript:openwin('/genergy2/setup/quickbuildinglist.asp?name=<%=request("name")%>',260,330);">Quick 
        Building search</a> </td>    
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        Reference Job # </td>
      <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <input name="refjob" type="text" size="30" value="<%=request("refjob")%>"> 
        <img src="/genergy2/setup/images/aro-rt.gif" border="0">&nbsp; <a href="javascript:openwin('/um/opslog/timesheet-beta/joblist.asp?name=<%=request("name")%>',260,320);">Quick 
        job search</a> 
		&nbsp;&nbsp;<span style="font-size:7pt;color:#999999;">(separate multiple jobs with a semicolon)</span>      
       </td>                 
    </tr>
 
   
		 
    <tr bgcolor="#dddddd"> 
      <td colspan="4"> <div style="margin-left:89px;">
          <input type="button" value="Save" onclick="checkform();">
          &nbsp;
          <input type="button" value="Cancel" onclick="closepage('cancel');">
        </div></td>
    </tr>
    <%end if%>
  </table>

    </form>
    </body>
    </html>
  <%
    
  case "save"
	
dim rfp,tm
    rfp = request("rfp")
	if rfp = "on" then
		rfp = 1
	else
		rfp = 0
	end if
    tm = request("tm")
	
    desc = request("desc")
    cStatus = request("cstatus")
    address_street = request("address_street")
    address_floor = request("address_floor")
	jfloor=trim(request("jfloor"))
    jcity = request("jcity") 
    jstate = request("jstate") 
    jzip  = request("jzip")
    projid = request("projid")
    if request("comppercent") = "N/A" then
		comppercent = 0
	else
		comppercent = request("comppercent")
	end if
'    jobnotes = request("jobnotes")
    primarybilling = request("primarybilling")
    secondarybilling = request("secondarybilling")
    
    if request("primary_amt") = "N/A" then
		primary_amt = 0
	else
		primary_amt = request("primary_amt")
	end if
	
	probability=request("probability")
	refjob = request("refjob")
	'response.write(tm)
	if probability="" then
		probability=0
	end if
    secondary_amt = request("secondary_amt")  
    customer  = request("customer")
	cust_name  = request("cust_name")
    company = request("company")
    jtype = request("jtype")
    jtypeid = request("jtypeid")
    jid = request("jid")
	bldgnum = request("bldgnum")
	taxstatusid = request("taxexempt")
	taxcert	= request("taxcert")
		if taxcert = "" then taxcert = 0 end if 
	contract = request("contract")
	contract_xp = request("contract_xp")
	jobready = request("jobready")
	
	
    set cnn = server.createobject("ADODB.connection")
    set rst = server.createobject("ADODB.recordset")
    cnn.open getConnect(0,0,"intranet")
    if secondary_amt="" then
    	secondary_amt=0
	end if
    
   if contract then
		strsql = "insert into MASTER_JOB (company, description,customer,customer_name,address_1,address_2,floor,city, state,zip_code, project_manager,status,percent_complete,job_notes,billing_method_1,amt_1,probability,billing_method_2,amt_2, type,type_id,[user], referencejob, rfp, bldgnum, taxstatusid,taxcert,contract, contract_expdate,readytoclose, tm) values ('"&company&"','"&left(desc,30)&"','"&customer&"','"&cust_name&"','"&left(address_street,30)&"','"&left(address_floor,30)&"','"&jfloor&"', '"&jcity&"','"&jstate&"', '"&jzip&"', '"&projid&"', '" & cstatus & "', '"&comppercent&"', 'See Notes System', '"&left(primarybilling,30)&"', " & primary_amt & "," & probability & ", '', "& 0 & ",'" & jtype & "',"&cint(trim(jtypeid))&",'"&getxmlusername()& "','" & trim(refjob) & "','" & trim(rfp) & "','" &trim(bldgnum)& "','" & trim(taxstatusid) & "','" &trim(taxcert)& "', '" &trim(contract)& "','" &trim(contract_xp)& "','" &trim(jobready)& "','" &trim(tm)& "')"
	else
		strsql = "insert into MASTER_JOB (company, description,customer,customer_name,address_1,address_2,floor,city, state,zip_code, project_manager,status,percent_complete,job_notes,billing_method_1,amt_1,probability,billing_method_2,amt_2, type,type_id,[user], referencejob, rfp, bldgnum, taxstatusid,taxcert,contract,readytoclose, tm) values ('"&company&"','"&left(desc,30)&"','"&customer&"','"&cust_name&"','"&left(address_street,30)&"','"&left(address_floor,30)&"','"&jfloor&"', '"&jcity&"','"&jstate&"', '"&jzip&"', '"&projid&"', '" & cstatus & "', '"&comppercent&"', 'See Notes System', '"&left(primarybilling,30)&"', " & primary_amt & "," & probability & ", '', "& 0 & ",'" & jtype & "',"&cint(trim(jtypeid))&",'"&getxmlusername()& "','" & trim(refjob) & "','" & trim(rfp) & "','" &trim(bldgnum)& "','" & trim(taxstatusid) & "','" &trim(taxcert)& "', '" &trim(contract)& "','" &trim(jobready)& "','" &trim(tm)&"')"
	end if
	
	response.write contract
	'response.write strsql
	cnn.Execute strsql
 
   
	strsql = "Select @@identity as id"   ' Get last autonumber created from SQL Server variable
	rst.open strsql, cnn, 1

    
   if rst.recordcount = 1 then 
	
		strsql = "insert into CustomerBidTracking (jobid, customerid,[primary]) values ('"&rst("id")&"','"&trim(customer)&"','1')"
		cnn.execute strsql
	
	dim message
	message = _
	"A Drawing Review job has been opened. (Job #"&rst("id") &")"&vbcrlf&_
	"Customer: "&cust_name&", "&left(address_street,30)&" "&left(address_floor,30)&" "&jfloor&", "&jcity&", "&jstate&" "&jzip&vbcrlf
	
	'10/23/2007 N.Ambo remove line because automatic emails no longer needed
    'if trim(ucase(jtypeid)) = "38" then sendmail "R&B@genergy.com", "GSA", "A Drawing Review Job has been opened", message	
    
    'Email to be sent to David if a new job is opened for PA (Tarun 10/31/2007)
	if trim(customer) = "GY-2142" then 
		sendmail  "Davide@genergy.com", "rb@genergy.com", "A PA Job has been opened. " & " JobNo : GY-0" & rst("id") , desc		
	end if
    %>
   <script>
   alert("Update saved")   
   opener.parent.frames['jobwindow'].document.location = "viewjob.asp?jid=<%=rst("id")%>"
   window.close()
    </script>
    <%
    else
     Response.write "ERROR: JOB COULD NOT BE CREATED."
   end if
    
  case else
end select
%>