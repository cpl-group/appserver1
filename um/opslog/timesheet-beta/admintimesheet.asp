<html>
<head>
<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(getKeyValue("user")) then
			Response.write "You are not currently Logged in."
		else
			if  not allowgroups("Department Supervisors,IT Services,AP_Admin")  then 
				setKeyValue "fMessage","Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "/genergy2/main.asp"
			end if	
		end if		
		user=getKeyValue("name")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
function viewtime(){
	var backtime = document.forms[0].backtime.value
	if (document.forms[0].backtime.checked) {
		var temp = 'viewtime.asp?back=' + backtime
	} else {
		var temp = 'viewtime.asp?back=0'
	}
	document.frames.admin.location= temp

}
function Print(uname){
    //var temp="timeprint.asp"
	if (uname == 'Print All Timesheets'){
		var temp="timetemplateall.asp"
	} else {
		var temp="timetemplate.asp?user=" + uname 
		}
	window.open(temp,"", "scrollbars=yes,width=800, height=600, status=no" );
}

function setUp(){
    var temp="admintimesetup.asp"
	//alert(temp);
    //window.open(temp,"", "scrollbars=yes, width=500, height=300, resizeable, status" );
	document.frames.admin.location= temp
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css"></head>
<body bgcolor="#eeeeee" text="#000000">
<form action="runpayroll.asp" method="post">
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#666699"> 
  <td><span class="standardheader">Review Employee Time Logs</span></td>
  <td align="right"><input type="button" name="Submit" value="Print Current View" onClick='javascript:document.frames.admin.focus();document.frames.admin.print()'></td>
</tr>
</table>
<table border=0 cellpadding="6" cellspacing="0" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
<tr>
  <td> 
  <a href="admintime.asp" target="admin" style="color:#333366;">Approve/Reject Time Sheets</a> &nbsp;|&nbsp; 
  <a href="javascript:viewtime()" style="color:#333366;">View Approved Time Sheets</a> &nbsp;|&nbsp; 
  <a href="javascript:setUp()" style="color:#333366;">Set Up Review Week</a>
  </td>
  <td align="right">
  	<table cellspacing="1" cellpadding="3" bgcolor="gray"><tr><td bgcolor="#DDDDDD">
    <input type="submit" name="payroll" value="Run Payroll">
	<select name="company">
		<option value="">Select Company</option>
		<%
		rst1.open "SELECT * FROM dbo.CompanyCodes WHERE code<>'AC' ORDER BY name", cnn1
		do until rst1.eof
			%><option value="<%=rst1("code")%>|<%=rst1("name")%>"><%=rst1("name")%></option><%
			rst1.movenext
		loop
		rst1.close
		%>
	</select>
	</td></tr></table>
  </td>
  <td align="right"><input type="checkbox" name="backtime" value="1">&nbsp;Back Time Sheets</td>
</tr>
</table>
<IFRAME name="admin" width="100%" height="86%" src="/um/opslog/timesheet-beta/admintime.asp" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME>
</form>
<%if trim(request("name"))<>"" then%>
<script>
alert("Pay Roll has been run for <%=request("name")%>");
</script>
<%end if%>
</body>
</html>