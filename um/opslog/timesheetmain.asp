<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<%
    if isempty(getkeyvalue("name")) then
%>
<script>
parent.location="../index.asp"
</script>
<%
    end if    
    user=Session("name")
job=request("job")


ReDim Categorys(5)

Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")
'sql="sp_invoice '"& job &"'"
'flag=request("flag")
'if flag<>1 then
'response.write("overwrite")
'cnn1.Execute sql, , adCmdStoredProc
'end if
%>
<script>
function toggleInfoDisplay(tag){
  tagdiv = tag;
  if (document.all[tagdiv].style.display == "none") {
    document.all[tagdiv].style.display = "inline";
  } else {
    document.all[tagdiv].style.display = "none";
  }
}

function opencontact(mode, cid, company) 
{
//	company = document.form1.company.value
	theURL="/genergy2_intranet/opsmanager/joblog/updatecontact.asp?mode=" + mode + "&contactid=" + cid + "&company="+company
	document.form1.cc1.checked=true;
	openwin(theURL,500,450)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function screencompany(company) {
    document.location.href="cis_manage.asp?company="+company	
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><title>Request For Invoice</title></head>
<body bgcolor="ffffff" onunload="javascript:opener.document.all['genjobtable'].bgColor ='#eeeeee';">
    <% 
  sqlstr2 = "select DISTINCT cat.category as category, sum(cat.hours) as hours from  " & _
        "(select employees.category as category, sum(hours) as hours from times, master_job, [employees] where jobno=master_job.id and username=matricola and times.entry_time> master_job.last_invoice and times.jobno='"& job & "' group by employees.category) as cat group by cat.category"
' cat 0 is not displayed so below sql is not displayed since inv_sub.cat inits to 0
'select 0 as category, sum(hours) as hours from times, master_job, [employees] where jobno=master_job.id and username=matricola and times.date>=master_job.last_invoice and times.jobno='" & job & "' group by category 
 
    rst1.Open sqlstr2, cnn1, 0, 1, 1
  while not rst1.eof
  
    categorys(rst1("category")) = rst1("hours")
    rst1.movenext
  
  wend
  rst1.close
  
  sqlstr2 = "select sum(hours) as hours, sum(hours_bill) as hours_bill, sum(billable) as billable from times, master_job where jobno=master_job.id and times.entry_time> master_job.last_invoice and times.jobno='"&job&"' "
 
  rst1.Open sqlstr2, cnn1, 0, 1, 1
    if not rst1.eof  then ' calculate hour sums and display form, iframe  and  not isnull(rst1("hours"))
      dim hourlyrate,hours,totalbillhours,diff,billable,buttontxt
      hourlyrate=0.00
      hours=Trim(rst1("hours"))
      totalbillhours=Trim(rst1("hours_bill"))
      diff=totalbillhours-hours
      if hours then ' if not isnull(hours) then
        billable=rst1("billable")
      end if
      if diff <=0 then
        totalbillhours=hours
      end if
	  rst1.close
	  rst1.open "SELECT a.company,a.customer,Billing_Contact FROM MASTER_JOB as a LEFT JOIN gy_MASTER_ARM_CUSTOMER as b ON a.Customer = b.Customer WHERE a.id ="&job,cnn1
	  if not rst1.eof then
	    if rst1("customer")="" or rst1("customer")="00-0000" then
	      buttontxt="There is no customer for this job.Please <a href=/genergy2_intranet/opsmanager/joblog/updatejob.asp?jid="&job&">edit</a>"
	    elseif rst1("Billing_Contact")="" then
		  buttontxt="There is no billing contact for this job.Please <a href=/genergy2_intranet/opsmanager/joblog/cis_detail.asp?cid="&rst1("customer")&"&company="&rst1("company")&" target=_blank>edit</a>"
		else
		  buttontxt="<input type=""submit"" name=""b1"" value=""Submit"">"
		end if
	  end if
	  cid = rst1("customer")
	  company = rst1("company")
	  bcontact = rst1("billing_contact")
	  rst1.close
    'end if
    %>
  
<form name="form1" method=post action=invoiceupdate.asp>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#6699cc">
      <td height="1" width="70%">
	  	<span class="standardheader"><a href="javascript:openwin('/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=<%=job%>',550,400)">Invoice for Job # <%=job%></a> | <a href="javascript:toggleInfoDisplay('billinginfo')">Show/Hide 
        Billing Contact</a></span></td>
    <td align="right" width="30%">
    <input type="button" name="Submit" value="Exit" onClick="javascript:opener.document.all['genjobtable'].bgColor ='#eeeeee';window.close()">
    <input type="hidden" name="job" value="<%=job%>">
  </td>
</tr>
</table>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border-top:1px solid #ffffff;">
<tr bgcolor="#eeeeee"> 
      <td> 
  <div id="billinginfo" style="display:none;border:1px solid #cccccc;width:100%;padding:3px;">
	      <table width="100%" border="0" cellpadding="0">
            <tr>
              <td width="31%" bgcolor="#CCCCCC"><font color="#000000">Billing Contact</font></td>
              <td width="69%" bgcolor="#CCCCCC">&nbsp;</td>
            </tr>
            <tr> 
              <td valign="top">
<input id="cc1" name="desigContact" type="radio" value="0" checked>
                Select from contacts setup for this client</td>
              <td valign="middle"> 
                <select name="invoicecontact" onclick="document.form1.cc1.checked=true">
                  <%
  sqlstr = "SELECT * FROM " & company & "_STANDARD_ARS_CONTACT WHERE customer ='"&cid&"'"
  rst.Open sqlstr, cnn1, 0, 1, 1     
  While not rst.EOF 
    cName     = rst("contact_name")
    cType     = rst("contact_type")
    cTitle    = rst("title")
    cContactID  = rst("contact")
%>
                  <option value="<%=cName%>" <% if bcontact = rst("contact") then %> selected <%end if%> ><%=cName%>, 
                  <%=cType%></option>
                  <%
	rst.movenext
	wend
	rst.close
%>
                </select> <a href="javascript:opencontact('new', '<%=cid%>','<%=company%>');">Add 
                new contact to this customer</a></td>
            </tr>
            <tr> 
              <td height="21" colspan="2" valign="top" bgcolor="#009999">
                  <input id="cc2" type="radio" name="desigContact" value="1">
                  or enter invoice specific information
                <table width="100%" border="0">
                  <tr bgcolor="#009999"> 
                    <td width="25%">Name</td>
                    <td width="75%"><input name="invCname" type="text" id="invCname4" size="50" onclick="document.form1.cc2.checked=true"></td>
                  </tr>
                  <tr bgcolor="#009999"> 
                    <td>Telephone</td>
                    <td><input name="invCTele" type="text" id="invCTele2" size="11" onclick="document.form1.cc2.checked=true"></td>
                  </tr>
                  <tr bgcolor="#009999"> 
                    <td>Email</td>
                    <td><input name="invCemail" type="text" id="invCemail2" size="50" onclick="document.form1.cc2.checked=true"></td>
                  </tr>
                </table>
                </td>
              <td>&nbsp;</td>
            </tr>
          </table>
        </div>
        <table border=0 cellpadding="3" cellspacing="0">
  <tr valign="top">
    <td>Billing Method:</td>
    <td>
    <input type="radio" name="invtype" value="1" checked> T &amp; M&nbsp<br>
    <input type="radio" name="invtype" value="0"> Contract
    </td>
    <td rowspan="5" width="20">&nbsp;</td>
    <td rowspan="5">
    Comment<br>
    <textarea name="invoice" cols="40" rows="6" wrap="PHYSICAL">This is an invoice for services rendered in connection with</textarea><br>
    <input type="button" name="b2" value="Clear Comment" onClick='javascript:invoice.value=""'>
    </td>
  </tr>
  <tr valign="top">
    <td>Hours:</td>
    <td>
    <table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc"width="300">
    <tr> 
      <td bgcolor="#66ff66" width="20%">Entry</td>
      <td bgcolor="#339999" width="20%">Junior</td>
      <td bgcolor="#ff9900" width="20%">Mid</td>
      <td bgcolor="#cc0000" width="20%"><span style="color:#ffffff;">Senior</span></td>
      <td bgcolor="#666699" width="20%"><span style="color:#ffffff;">Admin</span></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td><%=Categorys(1)%></td>
      <td><%=Categorys(2)%></td>
      <td><%=Categorys(3)%></td>
      <td><%=Categorys(4)%></td>
      <td><%=Categorys(5)%>&nbsp;</td>
    </tr>
    </table>
    </td>
  </tr>
  <tr valign="top">
    <td width="130">Total Hours:</td>
    <td><%=hours%></td>
  </tr>
  <tr valign="top">
    <td width="130">Total Billable:</td>
    <td><b><%=totalbillhours%></b></td>
  </tr>
  <tr valign="top">
  	<td width="130">Total Amount:</td>
	<td><input type="text" name="tot_amt" value="" size="4"></td>
  </table>  
  </td>
</tr>
<tr bgcolor="#dddddd">
  <td style="border-top:1px solid #cccccc;"><%=buttontxt%></td>
</tr>
</table>  
<IFRAME name="oplog" width="100%" height="250" src="timesheetsearch.asp?job=<%=job%>&edit=no" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
</form>
    <%
   else
     'response.Write(rst1.sqlstatement)
     rst1.close
       strsql="select id from master_job where id='" & job &"'"
     rst1.Open strsql, cnn1, 0, 1, 1
     if not rst1.eof then
    %>
    <br>
    <table border=0 cellpadding="3" cellspacing="0">
    <tr>
      <td>&nbsp; No time available for invoicing</td>
    </tr>
    <tr>
      <td><input type="button" name="Submit" value="Close Window" onClick="javascript:opener.document.all['genjobtable'].bgColor ='#eeeeee';window.close()"></td>
    </tr>
    </table>
<!--
    [[table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc"width="300"]]
    [[tr]] 
      [[td bgcolor="#66ff66" width="20%"]]Entry[[/td]]
      [[td bgcolor="#339999" width="20%"]]Junior[[/td]]
      [[td bgcolor="#ff9900" width="20%"]]Mid[[/td]]
      [[td bgcolor="#cc0000" width="20%"]][[span style="color:#ffffff;"]]Senior[[/span]][[/td]]
      [[td bgcolor="#666699" width="20%"]][[span style="color:#ffffff;"]]Admin[[/span]][[/td]]
    [[/tr]]
    [[/table]]
-->
    <%
     else    
    %>
    <br>
    <table border=0 cellpadding="3" cellspacing="0">
    <tr>
      <td>&nbsp; No such job</td>
    </tr>
    <tr>
      <td><input type="button" name="Submit" value="Close Window" onClick="javascript:opener.document.all['genjobtable'].bgColor ='#eeeeee';window.close()"></td>
    </tr>
    </table>

    <%
       end if
  end if
    %>    

</body>
</html>