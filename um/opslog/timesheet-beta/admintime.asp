<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
// 1/1/2008 N.Ambo added 'noaccess() script and modified processtime and rejecttime 
//This change allows only members of the Genergy_Supervisors and IT Services groups to accept/reject timesheets and throws
//an error message to the user if the user is not a member of the group
function noaccess() {
	alert("Sorry, you do not have sufficient permissions to perform this action. Please contact your administrator if you need these permissions.")
	 parent.location = '/genergy2/main.asp';
}
function processtime(userapp) {
  var temp
     <% if allowgroups("Genergy_Supervisors,IT Services") then %>
    temp="approvetime.asp?userapp=" + userapp
    document.location=temp
   <%else %>
     noaccess()
	<% end if %>    
}
function rejecttime(user1) {
  var temp
    <% if allowgroups("Genergy_Supervisors,IT Services") then %>
	temp="rejecttime.asp?user1=" + user1 
    window.open(temp,"", "scrollbars=yes,width=600, height=300, status=no, menubar=no" );
   <%else %>
		noaccess()
	<% end if %>
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" class="innerbody">

<%


user= Request.Querystring("user")

  

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")
    sqlstr= "select e.[first name]+' '+e.[last name] as name, u.startweek,u.endweek,u.username as u,substring(u.username,7,20) as subu ,sum(times.hours)as reghours, sum(times.overt) as ot, sum(times.hours) + sum(times.overt) as hours,u.submitted_date, approved from employees e join time_submission u on e.username=u.username join times on u.username=times.matricola where times.date between u.startweek and u.endweek and  u.submitted=1 and capproved =0 group by e.[first name],e.[last name],u.startweek,u.endweek,u.username,u.submitted_date, approved"


rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic



rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then 
%>
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr> 
  <td><b>Time Sheets Submitted for Current Week</b></td>
</tr>
<tr> 
  <td>No time sheets pending for current week.</td>
</tr>
</table>
<%
Else
x=0
%>

<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr> 
  <td bgcolor="#6699cc"><span class="standardheader">Time Sheets Submitted For Current Week</span></td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="1" bgcolor="#cccccc" width="100%">
<tr bgcolor="#dddddd"> 
  <td width="10%">Employee</td>     
  <td width="10%">Start Date</td>
  <td width="10%">End Date</td>
  <td width="10%">Total Hours</td>
  <td width="10%">Reg Hours</td>
  <td width="10%">OT Hours</td>
  <td width="20%">Date Submitted</td>
  <td width="20%">&nbsp;</td>
</tr>
<% While not rst1.EOF %>
<form name="form1" method="post" action="">

<tr <% if rst1("approved") = 1 then %> bgcolor="#00FF00" <% else %> bgcolor="#ffffff" <% end if %>> 
  <td><a href=<%="timesheetreview.asp?user=" & rst1("subu")%> ><%=rst1("name")%></a></td>
  <input type="hidden" name="user1" value="<%=rst1("u")%>">
  <input type="hidden" name="userapp" value="<%=rst1("subu")%>">
  <td><%=rst1("startweek")%></td>
  <td><%=rst1("endweek")%></td>
  <td><%=rst1("hours")%></td>
  <td><%=rst1("reghours")%></td>
  <td><%=rst1("ot")%></td>
  <td><%=rst1("submitted_date")%></td>
  <td> 
  <input type="button" name="reject" value="REJECT" onClick="rejecttime(user1.value)">
  <input type="button" name="approve" value="ACCEPT " onClick="processtime(userapp.value)">
  </td>
</tr>
</form>
    <%
x=x+1
rst1.movenext
Wend
%>
</table>

<table border="0" cellpadding="3" cellspacing="0" width="100%">
  <tr> 
    <td><%=x%> time sheets found</td>
  </tr>
</table>
<%
end if
rst1.close
%>
<br><br><br><br>
</body>
</html>