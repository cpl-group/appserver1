<%@Language="VBScript"%>
<%
if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'	Response.Redirect "http://www.genergyonline.com"
else
	if Session("ts") < 4 then 
		Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
		Response.Redirect "../main.asp"
	end if	
end if	
user="ghnet\"& Request.querystring("user")
user1=Request.querystring("user")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_main")
sql="select startweek, endweek from time_submission where username='payroll'"


rst1.Open sql, cnn1, 0, 1, 1

if not rst1.eof then
	startweek=rst1("startweek")
	endweek=rst1("endweek")
end if
rst1.close
strsql = "SELECT t.*, t.matricola AS Expr1, e.[first name]+ ' '+e.[last name] as name1, substring(e.username,7,20) as user1 FROM Times t join employees e on e.username=t.matricola WHERE (t.matricola = '"& user &"'  and t.[date] between '" & Startweek &"' and '" & endweek &"' ) order by t.date desc"

rst1.Open strsql, cnn1, 0, 1, 1

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
//document.main.location="http://www.yahoo.com"
function updateEntry(id){
	parent.frames.bottom.location="timedetail.asp?id="+id
}
function printtime(uname){
    //var temp="timeprint.asp"
	if (uname == 'Print All Timesheets'){
		var temp="timetemplateall.asp"
	} else {
		var temp="timetemplate.asp?user=" + uname 
		}
	window.open(temp,"", "scrollbars=yes,width=800, height=600, status=no" );
}

</script>
<body bgcolor="FFFFFF">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#3399CC" width="89%"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Timesheet 
      for <%=rst1("name1")%> for the week between <%=startweek%> and <%=endweek%></font></b></td>
    <td bgcolor="#3399CC" width="11%"> 
      <div align="right"><b><font face="Arial, Helvetica, sans-serif"><i> 
        <input type="button" name="Button" value="BACK" onClick="Javascript:history.back()">
        </i></font></b></div>
    </td>
  </tr>
</table>
<form name="form2" method="post" action="">
<input type="hidden" name="user" value="<%=rst1("user1")%>">
<input type="button" name="Submit2" value="Print This Timesheet" onClick="printtime(user.value)"></form>
<table width="100%" border="1" height="8" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr bgcolor="#CCCCCC" valign="middle"> 
    <td width=9% height="34"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">DATE</font></div>
    </td width=17%>
    <td width="5%" height="34"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">JOB #</font></div>
    </td>
    <td width="61%" height="34"><font face="Arial, Helvetica, sans-serif">Description 
      of Time</font></td>
    <td width="4%" bgcolor="#00CCFF" height="34"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">Hours</font></div>
    </td>
    <td width="4%" bgcolor="#3399CC" height="34"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">BH</font></div>
    </td>
    <td width="3%" bgcolor="#0033FF" height="34"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">OT</font></div>
    </td>
    <td width="4%" bgcolor="#0066CC" height="34"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">Expense Description</font></div>
    </td>
    <td width="4%" bgcolor="#3300CC" height="34"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif">Exp $$$</font></div>
    </td>
  </tr>
  <%
if not rst1.eof then
	Do until rst1.EOF 
%>
  <tr bgcolor="#CCCCCC" valign="middle"> 
    <form name=form1 method="post" action="">
      <input type="hidden" name="key" value="<%=rst1("id")%>">
      <td width=9% height="34"> <font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("date")%> 
        </font></td width=17%>
      <td width="5%" height="34"> <font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("jobno")%> 
        </font></td>
      <td width="61%" height="34"> <font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("description")%> 
        </font></td>
      <td width="4%" bgcolor="#00CCFF" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#000000"><%=rst1("hours")%> 
          </font></div>
      </td>
      <td width="4%" bgcolor="#3399CC" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("hours_bill")%> 
          </font></div>
      </td>
      <td width="3%" bgcolor="#0033FF" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("overt")%> 
          </font></div>
      </td>
      <td width="4%" bgcolor="#0066CC" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("expense")%> 
          </font></div>
      </td>
      <td width="4%" bgcolor="#3300CC" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">$<%=rst1("value")%> 
          </font></div>
      </td>
    </form>
  </tr>
  <%  
    rst1.movenext
    loop
end if
%>
</table>
 
</body>
</html>
