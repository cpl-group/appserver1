<html>
<head>
<title></title>
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
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#eeeeee" class="innerbody">
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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


sqlstr = "select e.[first name]+' '+e.[last name] as name,substring(u.username,7,20) as subu,u.submitted_date as date1 from employees e join time_submission u on e.username=u.username where u.capproved=1 "

if request.querystring("back") = 1 then 
	'sqlstr= sqlstr & " and e.active=1 and (u.Submitted_date IS NOT NULL)"
	sqlstr= sqlstr & " and (u.Submitted_date IS NOT NULL)"
else
	sqlstr= sqlstr & " and u.startweek between '"&start&"' and '"&end1&"'"
end if

sqlstr = sqlstr & " order by name"' order by date1
'response.write sqlstr
'response.end
	


rst2.ActiveConnection = cnn1
rst2.Cursortype = adOpenStatic

'response.write sqlstr
'response.end

rst2.Open sqlstr, cnn1, 0, 1, 1

if rst2.EOF then 
%>
	<table border="0" cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#6699cc"> 
    <td><span class="standardheader">Approved Time Sheets For Current Week</span></td>
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
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc"> 
  <td><span class="standardheader">Approved Time Sheets For Current Week</span></td>
  <td align="right"><input type="button" name="Submit22" value="Print All Timesheets" onClick="Print(this.value)"></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="1" bgcolor="#ffffff" width="100%">
<tr bgcolor="#dddddd" style="font-weight:bold;"> 
  <td width="25%">Employee</td>   	
  <td width="25%">Start Date </td>
  <td width="25%">End Date</td>
  <td width="25%">Submitted Time Stamp</td>
</tr>
<% While not rst2.EOF %>
<tr bgcolor="#ffffff" onclick="document.location='<%="timesheetreview.asp?user=" & rst2("subu")%>';" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" style="cursor:hand"> 
 <td><%=rst2("name")%></td>
 <input type="hidden" name="user1" value="<%="ghnet\"&rst2("subu")%>">
 <td><%=start%></td>
 <td><%=end1%></td>
 <td><%=rst2("date1")%></td>
</tr>
<%
x=x+1
rst2.movenext
Wend
%>
</table>

	<table border="0" cellpadding="3" cellspacing="0" width="100%">
  <tr> 
    <td colspan="2"><%=x%> Time sheets found.</td>
  </tr>
</table>
<%
end if
rst2.close
%>

</body>
</html>