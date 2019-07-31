<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
'7/5/2008 N.Ambo added drop down box to search by project manager

dim cnn, rst, strsql,company,SearchTextValue, showadv, showopen, startlabel, ccolor

set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

company = secureRequest("company")
'company = "GY"

if company <> "" then 
	showAdv = "block"
	showOpen = "none"
	startLabel = "Go To Open Job"
else
	showOpen = "block"
	showAdv = "none"
	startLabel = "Go To Advance Search"
end if

if trim(company) = "" then 
	SearchTextValue = "Insert Search Text"
	company 		= "GY"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Job Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2">
function customerdetail(mode) {
//	company = document.formadv.company.value
	theURL = "cis_update.asp?mode=" + mode
	openwin(theURL,600,300)
}
function opencontact(mode) 
{
//	company = document.formadv.company.value
	theURL="updatecontact.asp?mode=" + mode
	openwin(theURL,500,410)
}
function newjob(mode){
	theURL = "newjob.asp?mode=" + mode
	openwin(theURL,750,450)
}
function newcustomercontact(mode){
	theURL = "cis_manage.asp?mode=" + mode
	openwin(theURL,750,450)
}

function newvendor(mode){
	theURL = "/genergy2_intranet/opsmanager/Vendors/VendorView.asp?mode=" + mode
	openwin(theURL,750,450)
}


function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

//visual feedback functions for img buttons
function buttonOver(obj,clr){
  if (arguments.length == 1) { clr = "#eeeeee"; }
  obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
  if (arguments.length == 1) { clr = "#00FF00"; }
  obj.style.border = "1px solid " + clr;
}

function buttonOut(obj,clr){
  if (arguments.length == 1) { clr = "#6699cc"; }
  obj.style.border = "1px solid " + clr;
}
function clearDefault() {
  if(document.formadv.search.value=="Insert Search Text") {
    document.formadv.search.value=""
    }
}
function screencompany(company) {
  document.location.href="index.asp?company="+company 
}
function cleartext(){
	if (document.formadv.search.value == 'Insert Search Text'){
		document.formadv.search.value=''
	}
}
function openwin(url,mwidth,mheight){
  window.name="opener";
  popwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth);
  popwin.focus();
}
function jobPicked(job){
	document.formquick.jid.value = job
	document.formquick.jid.style.backgroundColor = "#FFFFFF"
	document.formquick.submit()
}
function switchview(){
	formquick.style.display = (formquick.style.display == 'none' ? 'block' : 'none')
	formadv.style.display = (formquick.style.display == 'none' ? 'block' : 'none')
//	openjobfield.style.backgroundColor = (openjobfield.style.backgroundColor == '#6699cc' ? '#eeeeee' : '#6699cc')
//	advfield.style.backgroundColor = (advfield.style.backgroundColor == '#6699cc' ? '#eeeeee' : '#6699cc')
	opentag.innerHTML = (opentag.innerHTML == 'Go To Open Job' ? 'Go To Advance Search' : 'Go To Open Job')
}
function checkform(){
	if (document.formquick.jid.value == "" ){
		alert("Please Enter a Job Number")
		return false
	}else  if(isNaN(document.formquick.jid.value)) {
     	alert("Please Enter a Job Number")
		return false 
	} else {
        return true ; 
	}
}
//-->
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">		
</head>
<body bgcolor="#eeeeee">
  <table width="100%" border="0" cellspacing="0" cellpadding="3" >
    <tr bgcolor="#6699cc"> 
      
    <td width="35%"><span class="standardheader"><font size="2">JOB LOG </font></span></td>
      <td width="65%" align="right" nowrap><span class="standardheader">
      <label type="" src="" value="New Job" onclick="newjob('new');" onmouseover="buttonOver(this);this.style.cursor='hand'" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #6699cc;">&nbsp;Create 
      a New Job&nbsp;</label>
      &nbsp;|&nbsp;<label type="" src="" value="Manage Customers &amp; Contacts" onclick="newcustomercontact('new');" onmouseover="buttonOver(this);this.style.cursor='hand'" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #6699cc;">&nbsp;Manage 
        Customers &amp; Contacts&nbsp;</label>&nbsp;|&nbsp;<label type="" src="" value="Browse Vendors" onclick="newvendor('new');" onmouseover="buttonOver(this);this.style.cursor='hand'" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #6699cc;">&nbsp;Manage 
        Vendors&nbsp;</label>&nbsp;|&nbsp;<label type="" src="" value="Browse Vendors" onclick="parent.document.frames.jobwindow.location='./start.htm'" onmouseover="buttonOver(this);this.style.cursor='hand'" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #6699cc;">&nbsp;Job Log Quick Help&nbsp;</label></span>
      </td>
    </tr>
  </table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan=3 align="right" nowrap id="openjobfield">&nbsp;
      <strong><a onclick="switchview()"><label id="opentag" onmouseover="this.style.cursor='hand'"><%=startLabel%></label></a></strong>
      &nbsp;</td>
  </tr>
  <tr>
    <td colspan=3><form name="formquick" method="post" action="viewjob.asp" target="jobwindow" style="display:<%=showOpen%>" onsubmit="return checkform()">
  <table width="100" border="0" cellspacing="0" cellpadding="0">
    <tr valign="middle"> 
      <td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;Open Job:</td>
      <td>&nbsp; 
        <input name="jid" type="text" size="10" maxlength="10" onClick="this.style.backgroundColor='white';" style="background-color:pink;"></td>
      <td nowrap>&nbsp; 
        <input name="jobview" type="submit" value="OPEN"> <img src="/genergy2/setup/images/aro-rt.gif" border="0">&nbsp;					<a href="javascript:openwin('/um/opslog/timesheet-beta/joblist.asp',260,320);">Quick job search</a>
</td>
  </tr>
</table>
</form>
</td>
  </tr>
  <tr>
    <td colspan=3><form name="formadv" method="post" action="jobsearch.asp" target="jobwindow" style="display:<%=showAdv%>">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" >
          <tr bgcolor="#eeeeee"> 
      <td colspan="2" nowrap>
	  <table width="100%" border=0 cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="7%" nowrap>&nbsp; <input name="mode" type="hidden" value="search"> 
                    <input name="custid" type="hidden" value=""> <select name="company" onchange="screencompany(this.value);">
                      <%
        rst.Open "select * from companycodes where active = 1 order by name", cnn
        if not rst.eof then
        do until rst.eof
        %>
                      <option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><font face="Arial, Helvetica, sans-serif"><%=rst("name")%></font></option>
                      <%    
        rst.movenext
        loop
		end if
        rst.close%>
                    </select> </td>
                    
    <td width="8%"> &nbsp; 
		<select name="pm" onchange="cleartext()" >
             <option value="">Any Project Manager
             <%
			if ucase(trim(company)) = "AC" then 
				strsql = "select * from Managers where Active is null order by lastname, firstname"
			else
				strsql = "select * from Managers where Active is null and companycode = '"&trim(company)&"'order by lastname, firstname"
			end if 
			
			rst.Open strsql, cnn
			if not rst.eof then
			do until rst.eof
			%>
				<option value="<%=rst("mid")%>" <%If trim(request("projid"))=trim(rst("mid")) then%>selected<%end if%>><%=rst("lastname")%>, 
                <%=rst("firstname")%></option>
			<%
			rst.movenext
			loop
		    
			end if
			rst.close
			%>
      </select> </td>
                    
                    
                  <td width="8%"> &nbsp; <select name="jtype" onchange="cleartext()">
                      <option value="">Any job type 
                      <%
	if ucase(trim(company)) = "AC" then 
		strsql = "SELECT Type,Type_ID, enable FROM master_job_types where job=1 ORDER BY enable desc,Type"
	else
		strsql = "SELECT Type,Type_ID, enable FROM master_job_types where job=1 and company = '"&trim(company)&"' ORDER BY enable desc,Type"
	end if 
	
    rst.Open strsql, cnn
    if not rst.eof then
    do until rst.eof
		dim fonttag, unfonttag
		if not rst("enable") then
			fonttag = "style='color:#555555'"
			'unfonttag = "</i></font>"	
		end if
    %>
        <option value="<%=rst("Type_ID")%>" <%if rst("enable")=0 then%> style="color:#555555" <%else%> style="color:black" <%end if%>> <%=left(rst("Type"),30)%></option>
    <%
    fonttag=""
	'unfonttag=""
    rst.movenext
    loop
    rst.close
    end if
    %>
                    </select> </td>
                  <td width="6%"> &nbsp; <input name="search" type="text" value="<%=searchtextvalue%>" size="20" onClick="cleartext();this.style.backgroundColor='white';" style="background-color:pink;"> 
                  </td>
                  <td width="9%"  nowrap>&nbsp;<strong>Order results by:</strong>&nbsp;</td>
                  <td width="56%"><select name="order">
                      <option value="job">Job</option>
                      <option value="actual_start_date">Job Opened Date</option>
                      <option value="description">Description</option>
                      <option value="address_1">Job Address</option>
                      <option value="pm_last,pm_first">Project Manager</option>
                      <option value="percent_complete">% Complete</option>
                    </select></td>
                </tr>
              </table></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td><table width="100%" border=0 cellpadding="0" cellspacing="3">
                <tr> 
                  <td width="36"><strong>Show:</strong></td>
                  <td width="25"><input type="checkbox" name="unstarted" value="unstarted" checked></td>
                  <td width="55">Unstarted</td>
                  <td width="25"><input name="inprogress" type="checkbox" value="in progress" checked></td>
                  <td width="69" nowrap>In Progress<br></td>
                  <td width="20"><input type="checkbox" name="closed" value="closed"></td>
                  <td width="40">Closed</td>
                  <td width="25"><input type="checkbox" name="rfp" value="1" checked></td>
                  <td width="25">RFP</td>
                  <td width="11">&nbsp;|&nbsp;</td>
                  <td width="54"><strong>Exclude/Show:</strong></td>
                  <td width="372"><select name="filter_fixed">
                      <option value="1">Exclude Reoccuring (M,Q,Y)</option>
                      <option value="2">Exclude 100% Complete</option>
                      <option value="3">Exclude Internal Jobs</option>
                      <option value="6">Exclude Internal & Reoccuring</option>
                      <option value="4">Exclude Internal,100%,Reoccuring</option>
                      <option value="0" selected>Show All</option>
                      <option value="5">Show only 100% complete</option>
					  <option value="7">Show only Reoccuring (M,Q,Y)</option>
                    </select>
                    <input type="submit" name="Submit" value="Search">
                  </td>
               </tr>
              </table>
        
        
      </td>
    </tr>
  </table>

<!--
[[iframe src="null.htm" name="jobwindow" style="border-top:1px solid #cccccc;" id="mainval" width="100%" height="80%" marginwidth="0" marginheight="0"  border=0 frameborder=0]][[/iframe]]
-->
</form>
</td>
  </tr>
</table>
</body>
</html>