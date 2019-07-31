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

'			Response.Redirect "http://www.genergyonline.com"
		else
			'response.write "Here Ando"
			'response.end			
		'if getKeyValue("ts") < 4 then 
				'getKeyValue("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				'Response.Redirect "../main.asp"
			'end if	
		end if

if Request.querystring("user")="" then
	user="ghnet\"&getKeyValue("login")
	name=getKeyValue("name")
else
	user="ghnet\"&Request.Querystring("user")
	name=Request.QueryString("user")
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
Set rst2 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")

sql="select startweek, endweek from time_submission where username='payroll'"
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
if not rst1.eof then
   
	startweek=rst1("startweek")
	endweek=rst1("endweek")
end if
sql2 = "SELECT [master_job].Description as description, company,Times.jobno, sum(Times.Hours) as hours, sum(Times.OverT) as overt, sum(Times.Value) as [Expense Value],case when [master_job].[id] > 6283 then left([master_job].[id],2)+'-00'+convert(varchar(4),[master_job].[id]) else '00-00'+convert(varchar(4),[master_job].[id]) end as tjob FROM [master_job] INNER JOIN Times ON [master_job].[ID] = Times.JobNo WHERE (Times.Date BETWEEN '"&startweek&"' AND '"&endweek&"') and times.matricola='"&user&"' group by [master_job].Description,Times.JobNo,[master_job].[id],company"

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

<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" onload="resize();">
<form name="form1" method="post" action="">
<table border=0 cellpadding="3" cellspacing="0" width="600">
<tr valign="middle">
  <td colspan="2"><div style="padding:3px;width:100%;border:1px solid #000000;"><b>Weekly Time Sheet</b></div></td>
</tr>
<tr>
  <td colspan="2" height="8"></td>
</tr>
<tr valign="middle">
  <td><b><%=name%></b></td>
  <td align="right">
  <table border=0 cellpadding="1" cellspacing="0">
  <tr>
    <td> From:</td>
    <td><%=startweek%></td>
  </tr>
  <tr>
    <td>To:</td>
    <td><%=endweek%></td>
  </td>
  </tr>
  </table>
  </td>
</tr>
<tr>
  <td colspan="2">
  <table border=0 cellpadding="3" cellspacing="0" width="100%" style="border:1px solid #000000;">
  <tr valign="bottom" style="font-weight:bold;"> 
    <td width="10%" style="border-bottom:1px solid #000000;">Job&nbsp;No.</td>
    <td width="10%" style="border-bottom:1px solid #000000;">Company</td>
    <td width="40%" style="border-bottom:1px solid #000000;">Description</td>
    <td width="15%" style="border-bottom:1px solid #000000;">Hours</td>
    <td width="15%" style="border-bottom:1px solid #000000;">Overtime</td>
    <td width="10%" style="border-bottom:1px solid #000000;">Expense</td>
  </tr>
  <% 
  expense=0
  total=0.0
  overt=0
  temp=0.0
  Do until rst2.eof
  %>
  <tr valign="top"> 
    <td style="border-bottom:1px solid #cccccc;"> <%=rst2("jobno")%> </td>
    <td style="border-bottom:1px solid #cccccc;"> <%=rst2("company")%> </td>
    <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;"> <%=rst2("description")%> </td>
    <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;"> <%=rst2("hours")%> </td>
    <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;"><%=rst2("overt")%></td>
    <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;"><%=rst2("expense value")%> </td>
  </tr>
  <%
  total=total+formatnumber(rst2("hours"), 2)
  expense=expense+rst2("expense value")
  overt=overt+formatnumber(rst2("overt"), 2)
  rst2.movenext
  loop
  %>
  <tr style="font-weight:bold;"> 
    <td colspan="3" style="border-left:1px solid #cccccc;">Total</td>
    <td width="15%" style="border-left:1px solid #cccccc;"> <%=total%> </td>
    <td width="15%" style="border-left:1px solid #cccccc;"> <%=overt%> </td>
    <td width="15%" style="border-left:1px solid #cccccc;"> <%=expense%></td>
  </tr>
  </table>
  
  </td>
</tr>
<tr>
  <td colspan="2">
  <br><br>
  <table width="100%" border="0">
  <tr>
    <td width="50%">___________________________________<br>Employee</td>
    <td width="50%">___________________________________<br>Supervisor</td>
  </tr>
  </table>
  </td>
</tr>
</table>
</body>
</html>
