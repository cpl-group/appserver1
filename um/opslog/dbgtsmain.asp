<html>
<head>
<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
parent.location="../index.asp"

</script>
<%
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
job=request("job")


ReDim Categorys(5)

Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_main")
sql="sp_invoice '"& job &"'"
flag=request("flag")
if flag<>1 then
'response.write("overwrite")
'cnn1.Execute sql, , adCmdStoredProc
end if
%>
<script>
function checktype(tm, cont){



}
</script>
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
	
	sqlstr2 = "select sum(hours) as hours, sum(hours_bill) as hours_bill, sum(billable) as billable from times, master_job where jobno=master_job.id and times.date>= master_job.last_invoice and times.jobno='"&job&"' "
    	
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
    %>
	
<form method=post action=file://///10.0.7.110/VSS/VSS/temp/invoiceupdate.asp>
<table width="100%" height="5%" border="0" bgcolor="#3399CC">
  <tr >
      <td height="1" width="71%"><font face="Arial, Helvetica, sans-serif"><i><b> 
        Invoice for Job # <%=job%> </b></i></font> </td>
    <td align="right" width="29%">
	<input type="button" name="Submit" value="Exit" onClick='javascript:document.location="null.htm"'>
    <input type="hidden" name="job" value="<%=job%>">
	</td>
  </tr>
</table>
  <table width="100%">
    <tr> 
      <td width="50%"> 
        <input type="submit" name="b1" value="Submit">
      </td>
      <td width="21%">Invoice Comment &nbsp 
        <input type="button" name="b2" value="Clear" onClick='javascript:invoice.value=""'>
      </td>
    </tr>
    <tr> 
      <td width="50%" height="26"> 
        <table width="100%" border="0">
          <tr> 
            <td><font face="Arial, Helvetica, sans-serif">T &amp; M&nbsp 
              <input type="radio" name="invtype" value="1" checked>
              Contract 
              <input type="radio" name="invtype" value="0">
              </font></td>
            <td bgcolor="#CCCCCC">&nbsp;</td>
          </tr>
          <tr> 
            <td><font face="Arial, Helvetica, sans-serif">Total Hours</font></td>
            <td bgcolor="#CCCCCC"> 
              <div align="right"><%=hours%> </div>
            </td>
          </tr>
          <tr> 
            <td><font face="Arial, Helvetica, sans-serif">Total Billable</font></td>
            <td bgcolor="#CCCCCC"> 
              <div align="right"> <%=totalbillhours%></div>
            </td>
          </tr>
          <tr> 
            <td height="2"> 
              <table width="100%" border="0">
                <tr> 
                  <td bgcolor="#339999" width="25%"> 
                    <div align="center"><b><%=Categorys(5)%></b></div>
                  </td>
                  <td bgcolor="#00FF00" width="25%"> 
                    <div align="center"><b><%=Categorys(1)%></b></div>
                  </td>
                  <td bgcolor="#00CC00" width="25%"> 
                    <div align="center"><b><%=Categorys(2)%></b></div>
                  </td>
                  <td bgcolor="#3399CC" width="25%"> 
                    <div align="center"><b><%=Categorys(3)%></b></div>
                  </td>
                  <td bgcolor="#FF0000" width="25%"> 
                    <div align="center"><b><%=Categorys(4)%></b></div>
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#339999" width="25%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Admin</font></b></div>
                  </td>
                  <td bgcolor="#00FF00" width="25%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Entry</font></b></div>
                  </td>
                  <td bgcolor="#00CC00" width="25%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Junior</font></b></div>
                  </td>
                  <td bgcolor="#3399CC" width="25%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Mid</font></b></div>
                  </td>
                  <td bgcolor="#FF0000" width="25%"> 
                    <div align="center"><b><font face="Arial, Helvetica, sans-serif">Senior</font></b></div>
                  </td>
                </tr>
              </table>
            </td>
            <td height="2"> 
              <div align="right"></div>
            </td>
          </tr>
        </table>
      </td>
      <td width="21%" height="26"> 
        <textarea name="invoice" cols="40" rows="6" wrap="PHYSICAL">This is an invoice for services rendered in connection with</textarea>
      </td>
    </tr>
  </table>  
</form>
<IFRAME name="oplog" width="100%" height="150" src="timesheetsearch.asp?job=<%=job%>" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
<IFRAME name="detail" width="100%" height="150" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>
		<%
		else
			rst1.close
    		strsql="select id from master_job where id='" & job &"'"
			rst1.Open strsql, cnn1, 0, 1, 1
			if not rst1.eof then
		%>
		<br><center>
  <h2><b> No Time Available For Invoicing</b></h2>
</center>
		<%
			else    
		%>
		<br><center><b><h2>
		No Such Job
		</h2></b></center>
		<%
    		end if
	end if
    %>    

</body>
