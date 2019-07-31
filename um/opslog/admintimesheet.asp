<html>
<head>
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("admin") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
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
    window.open(temp,"", "scrollbars=yes, width=500, height=300, resizeable, status" );
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<form>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

%>

<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Timesheets 
        Review </font></b></font></div>
    </td>
  </tr>
</table>
  <table width="100%" border="0">
    <tr>
      <td width="17%" height="2"> 
        <input type="button" name="appr" value="Approve\Reject Timesheets" onClick="Javascript:frames.admin.location='admintime.asp'">
      </td>
      <td width="41%" height="2"> 
        <div align="left"> 
          <input type="button" name="acct" value="View Approved Timesheets" onClick="viewtime()">
        </div>
      </td>
      <td height="2" width="42%"> 
        <div align="right"> 
          <input type="button" name="Submit3" value="Setup Review Week" onClick="setUp()">
        </div>
      </td>
    </tr>
    <tr>
      <td width="17%" height="2"> 
        <input type="button" name="Submit" value="Print Current View" onClick='javascript:document.frames.admin.focus();document.frames.admin.print()'>
      </td>
      <td width="41%" height="2">&nbsp; </td>
      <td height="2" width="42%"> 
        <div align="right"> 
          <input type="checkbox" name="backtime" value="1">
          <b><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Back 
          Timesheets</font></b></div>
      </td>
    </tr>
  </table>
  <p><IFRAME name="admin" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME></p>
</form></body>

</html>