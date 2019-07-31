<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("ts") < 4 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
Set rst2 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")


sql="select username, startweek, endweek from time_submission where startweek='" & Request.QueryString("revstart") &"' and endweek= '"& Request.Querystring("revend") &"' and username<>'payroll'"
'response.write sql
'response.end
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<html>
<head>
<script>
function openpopup(){
//configure "Open Logout Window

parent.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}

function resize(){
    parent.moveTo(0, 0)
}
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000" onload="resize()">
<form name="form1" method="post" action="">
  <div align="center">
    <div align="left"><i><font size="+2">Weekly Time Sheets For All Employees</font></i></div>
  </div>
 <% while not rst1.eof

	if not rst1.eof then
		user=rst1("username")
		startweek=rst1("startweek")
		endweek=rst1("endweek")
	end if
	sql2 = "SELECT employees.employee,[master_job].Description as description, Times.jobno, sum(Times.Hours) as hours, sum(Times.OverT) as overt, sum(Times.Value) as [Expense Value] FROM [master_job]  INNER JOIN Times ON [master_job].[id] = Times.JobNo,employees WHERE times.matricola=employees.username and (Times.Date BETWEEN (select startweek from time_submission where username='"&user&"') AND (select endweek from time_submission where username='"&user&"')) and times.matricola='"&user&"' group by [master_job].Description,Times.JobNo,employee"
	
	'response.write sql2
	'response.end
	rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly 
	
	

%>
  <div align="right"> <br>
    <i><br>
    <b> <%=user%> (<%=rst2("employee")%>) <br>
    From: <%=startweek%> <br>
    To: <%=endweek%> <br>
    </b></i></div>
  <br>
  <hr>
  <table width="100%">
    <tr> 
      <td width="10%"><i><b>Job No</b></i></td>
      <td width="42%"><i><b>Description</b></i></td>
      <td width="15%"><i><b>Reg. T</b></i></td>
      <td width="15%"><i><b>OverT</b></i></td>
      <td width="18%"><i><b>Expense Value</b></i></td>
    </tr>
  </table>
  <hr>
  <table width="100%">
    <% 
	total=0.0
	overt=0.0
	expense=0.0
	Do until rst2.eof
	
%>
    <tr> 
      <td width="10%" height="20"> <%=rst2("jobno")%> </td>
      <td width="42%" height="20"> <%=rst2("description")%> </td>
      <td width="15%" height="20"> <%=rst2("hours")%> </td>
	  <td width="15%" height="20"><%=rst2("overt")%></td>
	  <td width="18%" height="20"><%=rst2("expense value")%> </td>
    </tr>
    <%
	total=total+formatnumber(rst2("hours"), 2)
	overt=cdbl(overt)+formatnumber(rst2("overt"), 2)
	expense=cdbl(expense)+formatnumber(rst2("expense value"), 2)
rst2.movenext
loop

rst2.close
%>
</table><hr>
<table width="100%">
    <tr> 
      <td wi width="10%">&nbsp;</td>
      <td width="42%">Total</td>
      <td width="15%"> <%=total%> </td>
	  <td width="15%"> <%=overt%> </td>
	  <td width="18%"> <%=expense%></td>
	   
</table>
</form>
<%
rst1.movenext
wend
rst1.close
set cnn1=nothing
%>
<br>
<br>
<br>
<br>
<br>
<br>


</body>
</html>
