
<script>
function Print(uname){
    //var temp="timeprint.asp"
	if (uname == 'Print All Timesheets'){
		var temp="timetemplateall.asp"
	} else {
		var temp="timetemplate.asp?user=" + uname 
		}
	window.open(temp,"", "scrollbars=yes,width=800, height=600, status=no" );
}
</script>
<body bgcolor="#FFFFFF">
<%@Language="VBScript"%>
<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")
 
			sqlstr = "select startweek as s,endweek as e from time_submission where username='payroll'"
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
			start=rst1("s")
			end1=rst1("e")
					
					end if
					rst1.close
			
Set rst2 = Server.CreateObject("ADODB.recordset")

if request.querystring("back") = 1 then 
	sqlstr= "select e.[first name]+' '+e.[last name] as name,substring(e.username,7,20) as subu,u.submitted_date as date1 from employees e join time_submission u on e.username=u.username where e.active=1"
else
	sqlstr= "select e.[first name]+' '+e.[last name] as name,substring(u.username,7,20) as subu,u.submitted_date as date1 from employees e join time_submission u on e.username=u.username where u.approved=1 and u.startweek between '"&start&"' and '"&end1&"'"
end if
'response.write sqlstr
'response.end
	


rst2.ActiveConnection = cnn1
rst2.Cursortype = adOpenStatic

'response.write sqlstr
'response.end

rst2.Open sqlstr, cnn1, 0, 1, 1

if rst2.EOF then 
%>
	<table width="100%" border="0" dwcopytype="CopyTableCell">
  <tr> 
    <td bgcolor="#3399CC" > 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Approved 
        Timesheets For Current Week</font></b></font></div>
    </td>
  </tr>
  <tr>
    <td height="21"></td>
  </tr>
  <tr> 
    <td height="2"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b>No Timesheets 
        Have Been Approved For The Current Week</b></font></div>
    </td>
  </tr>
</table>
<%
Else
x=0
%>
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="2" > 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Approved 
        Timesheets For Current Week</font></b></font></div>
    </td>
  </tr>
  <tr>
  	<td height="2"> 
      <input type="button" name="Submit22" value="Print All Timesheets" onClick="Print(this.value)">
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0">
        <tr bgcolor="#CCCCCC"> 
		  <td bgcolor="#CCCCCC" width="8%"><font face="Arial, Helvetica, sans-serif" color="#000000">Employee</font></td>   	
          <td bgcolor="#CCCCCC" width="20%"><font face="Arial, Helvetica, sans-serif" color="#000000">Start 
            Date </font></td>
          <td width="13%"><font face="Arial, Helvetica, sans-serif" color="#000000">End 
            Date</font></td>
         <td width="13%"><font face="Arial, Helvetica, sans-serif" color="#000000">Submitted Time Stamp</font></td>
        </tr>
        <% While not rst2.EOF %>
        <tr> 
          <td width="8%"><font face="Arial, Helvetica, sans-serif"><a href=<%="timesheetreview.asp?user=" & rst2("subu")%> ><%=rst2("name")%></a></font></td>
		  <input type="hidden" name="user1" value="<%="ghnet\"&rst2("subu")%>">
          <td width="20%"><font face="Arial, Helvetica, sans-serif"><%=start%></font></td>
          <td width="13%"><font face="Arial, Helvetica, sans-serif"><%=end1%> </font></td>
		   <td width="13%"><font face="Arial, Helvetica, sans-serif"><%=rst2("date1")%>
		  </font></td>
		  
        </tr>
        <%
		x=x+1
		rst2.movenext
		Wend
		%>
      </table>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF"><%=x%> 
        Timesheets Found </font></b></font></div>
    </td>
  </tr>
</table>
<%
end if
rst2.close
%>

