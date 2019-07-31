<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function processtime(userapp) {
	var temp
		temp="approvetime.asp?userapp=" + userapp
		document.location=temp
}
function rejecttime(user1) {
	var temp
		temp="rejecttime.asp?user1=" + user1 
		window.open(temp,"", "scrollbars=yes,width=600, height=300, status=no, menubar=no" );
}
</script>
<body bgcolor="#FFFFFF">
<%@Language="VBScript"%>
<%


user= Request.Querystring("user")

	

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")
		sqlstr= "select e.[first name]+' '+e.[last name] as name, u.startweek,u.endweek,u.username as u,substring(u.username,7,20) as subu ,sum(times.hours)as hours ,u.submitted_date from employees e join time_submission u on e.username=u.username join times on u.username=times.matricola where times.date between u.startweek and u.endweek and  u.submitted=1 and approved =0 group by e.[first name],e.[last name],u.startweek,u.endweek,u.username,u.submitted_date "
	


rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic



rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then 
%>
<table width="100%" border="0" dwcopytype="CopyTableCell">
  <tr> 
    <td bgcolor="#3399CC" > 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Submitted 
        Timesheets for Current Week</font></b></font></div>
    </td>
  </tr>
  <tr>
    <td height="21"></td>
  </tr>
  <tr> 
    <td height="2"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b>NO TIMESHEETS 
        WAITING FOR CURRENT WEEK</b></font></div>
    </td>
  </tr>
</table>
<%
Else
x=0
%>

<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" > 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Submitted 
        </font></b></font><font color="#FFFFFF"> Timesheets for Current Week</font></b></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0">
        <tr bgcolor="#CCCCCC"> 
		  <td bgcolor="#CCCCCC" width="17%"><font face="Arial, Helvetica, sans-serif" color="#000000">Employee</font></td>   	
          <td bgcolor="#CCCCCC" width="17%"><font face="Arial, Helvetica, sans-serif" color="#000000">Start 
            Date </font></td>
          <td width="16%"><font face="Arial, Helvetica, sans-serif" color="#000000">End 
            Date</font></td>
          <td width="16%"><font face="Arial, Helvetica, sans-serif" color="#000000">Total 
            Hours</font></td>
		  <td width="16%"><font face="Arial, Helvetica, sans-serif" color="#000000">Date 
            Submitted</font></td>
		  <td width="18%"><font face="Arial, Helvetica, sans-serif" color="#000000"></font></td>
        </tr>
        <% While not rst1.EOF %>
		<form name="form1" method="post" action="">

        <tr> 
            <td width="17%"><font face="Arial, Helvetica, sans-serif"><a href=<%="timesheetreview.asp?user=" & rst1("subu")%> ><%=rst1("name")%></a></font></td>
		  <input type="hidden" name="user1" value="<%=rst1("u")%>">
		  <input type="hidden" name="userapp" value="<%=rst1("subu")%>">
            <td width="17%"><font face="Arial, Helvetica, sans-serif"><%=rst1("startweek")%></font></td>
            <td width="16%"><font face="Arial, Helvetica, sans-serif"><%=rst1("endweek")%></font></td>
		    <td width="16%"><font face="Arial, Helvetica, sans-serif"><%=rst1("hours")%></font></td>
		    <td width="16%"><font face="Arial, Helvetica, sans-serif"><%=rst1("submitted_date")%></font></td>
		    <td width="18%"><font face="Arial, Helvetica, sans-serif"> 
              <input type="button" name="reject" value="REJECT" onClick="rejecttime(user1.value)">
          <input type="button" name="approve" value="ACCEPT " onClick="processtime(userapp.value)">
		  
		  </font></td>
        </tr></form>
        <%
		x=x+1
		rst1.movenext
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
rst1.close
%>

