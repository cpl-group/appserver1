<html>
<head>
<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
    if isempty(getKeyValue("name")) then
%>
<script>
parent.location="../index.asp"

</script>
<%
      Response.Redirect "http://www.genergyonline.com"
    end if    
    user=getKeyValue("name")
job=request("job")


ReDim Categorys(5)

Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")
'sql="sp_invoice '"& job &"'"
'flag=request("flag")
'if flag<>1 then
'response.write("overwrite")
'cnn1.Execute sql, , adCmdStoredProc
'end if
%>
<script>
function checktype(tm, cont){



}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="ffffff">
    <% 
  sqlstr2 = "select DISTINCT cat.category as category, sum(cat.hours) as hours from  " & _
        "(select employees.category as category, sum(hours) as hours from times, master_job, [employees] where jobno=master_job.id and username=matricola and times.date> master_job.last_invoice and times.jobno='"& job & "' group by employees.category) as cat group by cat.category"
' cat 0 is not displayed so below sql is not displayed since inv_sub.cat inits to 0
'select 0 as category, sum(hours) as hours from times, master_job, [employees] where jobno=master_job.id and username=matricola and times.date>=master_job.last_invoice and times.jobno='" & job & "' group by category 
  
    rst1.Open sqlstr2, cnn1, 0, 1, 1
  while not rst1.eof
  
    categorys(rst1("category")) = rst1("hours")
    rst1.movenext
  
  wend
  rst1.close
  
  sqlstr2 = "select sum(hours) as hours, sum(hours_bill) as hours_bill, sum(billable) as billable from times, master_job where jobno=master_job.id and times.date> master_job.last_invoice and times.jobno='"&job&"' "
      
  rst1.Open sqlstr2, cnn1, 0, 1, 1
    if not rst1.eof and  not isnull(rst1("hours")) then
        hourlyrate=0.00
      hours=Trim(rst1("hours"))
    totalbillhours=Trim(rst1("hours_bill"))
    diff=totalbillhours-hours
    if not isnull(hours) then
      billable=rst1("billable")
    end if
    if diff <=0 then
        totalbillhours=hours
    end if
    'end if
    %>
  
<form method=post action=invoiceupdate.asp>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#6699cc">
    <td height="1" width="71%"><span class="standardheader">Invoice for Job # <%=job%></span></td>
    <td align="right" width="29%">
    <input type="button" name="Submit" value="Exit" onClick="javascript:opener.document.all['genjobtable'].bgColor ='#eeeeee';window.close()">
    <input type="hidden" name="job" value="<%=job%>">
  </td>
</tr>
</table>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border-top:1px solid #ffffff;">
<tr bgcolor="#eeeeee"> 
  <td>
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
  </table>  
  </td>
</tr>
<tr bgcolor="#dddddd">
  <td style="border-top:1px solid #cccccc;"><input type="submit" name="b1" value="Submit"></td>
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