<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
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

if Request.querystring("user")="" then
	user="ghnet\"&Session("login")
	name=Session("name")
else
	user="ghnet\"&Request.Querystring("user")
	name=Request.QueryString("user")
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
Set rst2 = Server.CreateObject("ADODB.Recordset")
cnn1.Open application("cnnstr_main")

sql="select startweek, endweek from time_submission where username='payroll'"
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
if not rst1.eof then
   
	startweek=rst1("startweek")
	endweek=rst1("endweek")
end if
sql2 = "SELECT [Job Log].Description as description, Times.jobno, sum(Times.Hours) as hours, sum(Times.OverT) as overt, sum(Times.Value) as [Expense Value],case when [job log].[entry id] > 6283 then left([entry type],2)+'-00'+convert(varchar(4),[entry id]) else '00-00'+convert(varchar(4),[entry id]) end as  tjob  FROM [Job Log]  INNER JOIN Times ON [Job Log].[Entry ID] = Times.JobNo WHERE (Times.Date BETWEEN '"&startweek&"' AND '"&endweek&"') and times.matricola='"&user&"' group by [Job Log].Description,Times.JobNo,[job log].[entry id],[job log].[entry type]"


rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly 



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
    <div align="left"><i><font size="+2">Weekly Time Sheet</font></i></div>
  </div>
  <div align="right"> <br>
    <i><br>
    <b><%=name%> <br>
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
	expense=0
	total=0.0
	overt=0
	temp=0.0
	Do until rst2.eof
	
%>
    <tr> 
      <td width="10%" height="20"> <%=rst2("tjob")%> </td>
      <td width="42%" height="20"> <%=rst2("description")%> </td>
      <td width="15%" height="20"> <%=rst2("hours")%> </td>
	  <td width="15%" height="20"><%=rst2("overt")%></td>
	  <td width="18%" height="20"><%=rst2("expense value")%> </td>
    </tr>
    <%
	total=total+formatnumber(rst2("hours"), 2)
	expense=expense+rst2("expense value")
	overt=overt+formatnumber(rst2("overt"), 2)
rst2.movenext
loop
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
<table width="100%" border="0">
  <tr>
    <td>Employee&nbsp&nbsp&nbsp___________________________________ </td>
    <td>Supervisor&nbsp&nbsp___________________________________ </td>
  </tr>
</table>
<br>
<br>
<br>
<br>
<br>
<br>


</body>
</html>
