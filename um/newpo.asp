<%@Language="VBScript"%>
<%


Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")

cnn1.Open application("cnnstr_main")

dim jid, fullname, sessionname
jid = request("jid")
if Session("name")<>"" then
  sessionname = (Session("name"))
end if
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
Dim tcolor,Desc,company,job,jtype,cstatus
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
    <a href="<%="/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=" & jid %>">General Info</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobtime.asp?jid=" &jid%>">Job Time</a> &nbsp;|&nbsp; <a href="<%="/genergy2_intranet/opsmanager/joblog/jobfolder.asp?jid="&jid%>">Job Folder</a> &nbsp;|&nbsp; <b><a href="/um/opslog/posearch.asp?caller=<%=caller%>&select=jobnum&findvar=<%=jid%>">Purchase Orders</a></b>&nbsp;|&nbsp; <a href="/um/war/jc/jc<%=permissionflag%>.asp?c=<%=company%>&j<%=lcase(left(company,1))%>=<%=job%>&jid=<%=jid%>">Job Cost</a>
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

    <b>New Purchase Order</b>
    <input type="hidden" name="job" value="<%=job%>">

  </td>
</tr>
<tr> 
  <td colspan="2" bgcolor="#eeeeee"> 
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
      <td><input type="text" name="vendor"></td>
    </tr>
    <tr>
      <td>Requisitioner</td>
      <td>  
      <select name="req">
      <%Set rst3 = Server.CreateObject("ADODB.recordset")
      sqlstr = "select [first name], [last name], [last name]+', '+ [first name] as name, substring(username,7,20) as user1, active from employees order by [last name]"
      rst3.Open sqlstr, cnn1, 0, 1, 1
      if not rst3.eof then
        do until rst3.eof 
        if rst3("active") then 
        fullname = trim(rst3("first name")) + " " + trim(rst3("last name"))
        %> 
        <option value="<%=rst3("user1")%>" <% if trim(fullname) = trim(sessionname) then %> selected<% end if %>><%=rst3("name")%></option>
        <%
        end if
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
      <select name="jobnum">
      <%Set rst2 = Server.CreateObject("ADODB.recordset")
      sqlstr = "select distinct id from master_job order by id desc"
      rst2.Open sqlstr, cnn1, 0, 1, 1
      if not rst2.eof then
        do until rst2.eof 
        %>
        <option value="<%=rst2("id")%>"><%=rst2("id")%></option>
        <%
        rst2.movenext
        loop
      end if 
      rst2.close
      %>
      </select>
      <% end if %>
      </td>
    </tr>
    </table>
    <!-- end first column -->
    </td>
    <td>
    <!-- begin second column -->
    <p>PO Description<br><textarea name="description" cols="50" rows="3"></textarea></p>
    <!-- end second column -->
    </td>
  </tr>
  </table>
  </td>
</tr>
<tr> 
  <td colspan="2" bgcolor="#dddddd" style="border-top:1px solid #cccccc;"> 
  <input type="submit" name="choice" value="Save">
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
    <td><img src="/um/opslog/images/btn-new_po.gif" value="New Purchase Order" id="new_po" align="middle" onclick="newPO();" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;">&nbsp;</td>
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
