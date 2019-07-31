<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
    if isempty(getKeyValue("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
      '     Response.Redirect "http://www.genergyonline.com"
    end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
Set rst2 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")


sql="select time_submission.username, startweek, endweek,employees.employee from time_submission,employees where startweek='" & Request.QueryString("revstart") &"' and endweek= '"& Request.Querystring("revend") &"' and time_submission.username<>'payroll' and time_submission.username = employees.username order by employees.employee asc"
'response.write sql + "<br>"
'response.end
rst1.Open sql, cnn1', adOpenStatic, adLockReadOnly
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
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" onload="resize()">
<form name="form1" method="post" action="">
<table border=0 cellpadding="3" cellspacing="0" width="550">
<tr valign="middle">
  <td colspan="2"><div style="padding:3px;width:100%;border:1px solid #000000;"><b>Weekly Time Sheet For All Employees</b></div></td>
</tr>
 <% while not rst1.eof

  if not rst1.eof then
    user=rst1("username")
    startweek=rst1("startweek")
    endweek=rst1("endweek")
  end if
  sql2 = "SELECT employees.employee,[master_job].Description as description, Times.jobno,substring([master_job].job,1,2) as JobID, sum(Times.Hours) as hours, sum(Times.OverT) as overt, sum(Times.Value) as [Expense Value] FROM [master_job]  INNER JOIN Times ON [master_job].[id] = Times.JobNo,employees WHERE times.matricola=employees.username and (Times.Date BETWEEN (select startweek from time_submission where username='"&user&"') AND (select endweek from time_submission where username='"&user&"')) and times.matricola='"&user&"' group by [master_job].Description,Times.JobNo,[master_job].job,employee"
  
'response.Write(sql2 + "<br>")
  rst2.Open sql2, cnn1', adOpenStatic, adLockReadOnly 
  
  

  %>
<tr>
  <td colspan="2" height="8"></td>
</tr>
<tr>
  <td><b><%=user%> (<%if not rst2.eof then response.write rst2("employee")%>)</b></td>
  <td align="right">
  <table border=0 cellpadding="1" cellspacing="0">
  <tr>
    <td>From:</td>
    <td><%=startweek%></td>
  </tr>
  <tr>
    <td>To:</td>
    <td><%=endweek%></td>
  </tr>
  </table>
  </td>
</tr>
<tr>
  <td colspan="2">
  <table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #000000;">
  <tr valign="bottom" style="font-weight:bold;"> 
	 <td width="5%" style="border-bottom:1px solid #000000;">&nbsp;&nbsp;</td>
    <td width="10%" style="border-bottom:1px solid #000000;">Job&nbsp;No.</td>
    <td width="40%" style="border-bottom:1px solid #000000;">Description</td>
    <td width="15%" style="border-bottom:1px solid #000000;">Hours</td>
    <td width="15%" style="border-bottom:1px solid #000000;">Overtime</td>
    <td width="15%" style="border-bottom:1px solid #000000;">Expense</td>
  </tr>

    <% 
  total=0.0
  overt=0.0
  expense=0.0
  Do until rst2.eof
  
%>
  <tr> 
	  <td style="border-bottom:1px solid #cccccc;"> <%=trim(rst2("JobID"))%> </td>
    <td style="border-bottom:1px solid #cccccc;"> <%=rst2("jobno")%> </td>
    <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;"> <%=rst2("description")%> </td>
    <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;" align="right"> <%=rst2("hours")%> </td>
    <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;" align="right"><%=rst2("overt")%></td>
    <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;" align="right"><%=rst2("expense value")%> </td>
  </tr>
  <%
  total=total+formatnumber(rst2("hours"), 2)
  overt=cdbl(overt)+formatnumber(rst2("overt"), 2)
  expense=cdbl(expense)+formatnumber(rst2("expense value"), 2)
  rst2.movenext
  loop
  
  rst2.close
  %>
  
  <tr style="font-weight:bold;"> 
    <td colspan="2"></td>
	<td>Total</td>
    <td align="right" style="border-left:1px solid #cccccc;"> <%=total%> </td>
    <td align="right" style="border-left:1px solid #cccccc;"> <%=overt%> </td>
    <td align="right" style="border-left:1px solid #cccccc;">$ <%=expense%></td>
  </tr>
  </table>
  </td>
</tr>
<%
rst1.movenext
wend
rst1.close
set cnn1=nothing
%>
</table>
</form>
<br>
<br>
<br>
<br>
<br>
<br>


</body>
</html>
