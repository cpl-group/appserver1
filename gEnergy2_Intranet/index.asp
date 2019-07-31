<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn, rst, strsql,company,SearchTextValue

set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open application("cnnstr_main")

company = secureRequest("company")

if trim(company) = "" then 
	SearchTextValue = "Insert Search Text"
	company 		= "AC"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Job Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2">
function customerdetail(mode) {
//	company = document.form1.company.value
	theURL = "cis_update.asp?mode=" + mode
	openwin(theURL,600,300)
}
function opencontact(mode) 
{
//	company = document.form1.company.value
	theURL="updatecontact.asp?mode=" + mode
	openwin(theURL,500,410)
}
function newjob(mode){
	theURL = "newjob.asp?mode=" + mode
	openwin(theURL,750,450)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
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
function clearDefault() {
  if(document.form1.search.value=="Insert Search Text") {
    document.form1.search.value=""
    }
}
function screencompany(company) {
  document.location.href="index.asp?company="+company 
}
function cleartext(){
	if (document.form1.search.value == 'Insert Search Text'){
		document.form1.search.value=''
	}
}
//-->
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">		
</head>
<body bgcolor="#eeeeee">
<form name="form1" method="post" action="jobsearch.asp" target="jobwindow">
  <table width="100%" border="0" cellspacing="0" cellpadding="3">
    <tr bgcolor="#6699cc"> 
      <td width="35%"><span class="standardheader">JOB LOG </span></td>
      <td width="65%" align="right"> <a href="javascript:parent.location='frameset.html';"><img src="/um/opslog/images/btn-job_log_home.gif" style="border:1px solid #336699;" onmouseover="buttonOver(this,'#000000');" onmousedown="buttonDn(this);" onmouseout="buttonOver(this,'#336699');" border="0"></a> 
        <!--[[input type="button" value="New Job" onclick="newjob('new');"]][[input type="button" value="New Customer" onclick="customerdetail('new');"]][[input type="button" value="New Contact" onclick="opencontact('new');"]]-->
      </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;"><table border=0 cellspacing="3" cellpadding="0">
          <tr> 
            <td> <input name="mode" type="hidden" value="search"> <input name="custid" type="hidden" value=""> 
            </td>
            <td> <select name="company" onchange="screencompany(this.value);">
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
            <td> <select name="jtype" onchange="cleartext()">
                <option value="">Any job type 
                <%
	if ucase(trim(company)) = "AC" then 
		strsql = "SELECT Type,Type_ID FROM master_job_types where job=1 ORDER BY Type"
	else
		strsql = "SELECT Type,Type_ID FROM master_job_types where job=1 and company = '"&trim(company)&"' ORDER BY Type"
	end if 
	
    rst.Open strsql, cnn
    if not rst.eof then
    do until rst.eof
    %>
                <option value="<%=rst("Type_ID")%>"><%=left(rst("Type"),30)%></option>
                <%
    rst.movenext
    loop
    
    end if
    rst.close
    %>
              </select> </td>
            <td><input name="search" type="text" value="<%=searchtextvalue%>" size="20" onClick="cleartext()"> 
            </td>
            <td><input type="submit" name="Submit" value="Search"></td>
            <td><input type="reset" name="Reset" value="Reset"></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;"> <table border=0 cellspacing="3" cellpadding="0">
          <tr> 
            <td>Show:</td>
            <!--
    [[td]]
    [[select name="order"]]
    [[option value="status" selected]]Status
    [[option value="job"]]Job Number
    [[option value="description"]]Job Description
    [[option value="address_1"]]Job Address
    [[option value="pm_last"]]Project Manager
    [[/select]]
    [[/td]]
-->
            <td><input type="checkbox" name="unstarted" value="unstarted" checked></td>
            <td>Unstarted</td>
            <td><input name="inprogress" type="checkbox" value="in progress" checked></td>
            <td>In Progress<br></td>
            <td><input type="checkbox" name="closed" value="closed"></td>
            <td>Closed</td>
            <td><input type="checkbox" name="rfp" value="1"></td>
            <td>RFP</td>
          </tr>
        </table>
        
      </td>
      <td style="border-bottom:1px solid #cccccc;"><table border=0 cellspacing="3" cellpadding="0">
          <tr> 
            <td>Order results by:</td>
            <!--
    [[td]]
    [[select name="order"]]
    [[option value="status" selected]]Status
    [[option value="job"]]Job Number
    [[option value="description"]]Job Description
    [[option value="address_1"]]Job Address
    [[option value="pm_last"]]Project Manager
    [[/select]]
    [[/td]]
-->
            <td><input type="radio" name="order" value="status" checked></td>
            <td>Status</td>
            <td><input name="order" type="radio" value="job"></td>
            <td>Job Number<br></td>
            <td><input type="radio" name="order" value="description"></td>
            <td>Job Description<br></td>
            <td></td>
            <td><input type="radio" name="order" value="address_1"></td>
            <td>Job Address<br></td>
            <td><input type="radio" name="order" value="pm_last"></td>
            <td>Project Manager</td>
          </tr>
        </table></td>
    </tr>
  </table>
<!--
[[iframe src="null.htm" name="jobwindow" style="border-top:1px solid #cccccc;" id="mainval" width="100%" height="80%" marginwidth="0" marginheight="0"  border=0 frameborder=0]][[/iframe]]
-->
</form>
</body>
</html>