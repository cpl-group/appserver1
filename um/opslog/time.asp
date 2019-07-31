<html>
<head>
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(Session("name")) then
%>
<script>
parent.location="../index.asp"

</script>
<%
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("ts") < 4 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
		user1=Session("login")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

%>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function openpopup(){
//configure "Open Logout Window
parent.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}

function setUp(){
    var temp="timesetup.asp"
	//alert(temp);
    window.open(temp,"", "scrollbars=yes, width=500, height=300, resizeable, status" );
}

function Print(name){
    //var temp="timeprint.asp"
	var temp="timeformatemp.asp"
	window.open(temp,"", "scrollbars=yes, width=700, height=500, resizeable, status" );
}
function timesheetjob(name,start,end1){
name=document.forms[0].name.value
start=document.forms[0].start.value
	document.location="processtime.asp?name=" + name+ "&start=" +start + "&end1=" + end1
}




</script>
<STYLE>
<!--
A.ssmItems:link		{color:black;text-decoration:none;}
A.ssmItems:hover	{color:black;text-decoration:none;}
A.ssmItems:active	{color:black;text-decoration:none;}
A.ssmItems:visited	{color:black;text-decoration:none;}
//-->

<script>
</STYLE>

</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr>
    <td bgcolor="#3399CC">
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">TIMESHEET 
        : <%=user%></font></b></font></div>
    </td>
  </tr>
</table>

<%sqlstr = "select startweek as s,endweek as e, username from user_cost where substring(username,7,20)='"&user1&"'"
   			rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<form name="form1" method="post" action="">
<table width="116" border="0" height="13" align="left">
  <tr> 
    <td height="45" width="56"> 
      <div align="left"> 
        <input type="button" name="Submit" value="Setup Week" onClick="setUp()">
      </div>
    </td>
    <td height="45" width="694"> 
      <div align="left"> 
		<input type="hidden" name="name" value="<%=Session("login")%>">	  
        <input type="button" name="Submit" value="Print Timesheet" onClick="Print(name.value)">
        </div>
    </td>
	<td>
	<input type="hidden" name="start" value="<%=rst1("s")%>">
	<input type="hidden" name="end1" value="<%=rst1("e")%>">
	<input type="button" name="submit1" value="Submit Timesheet" onclick="timesheetjob( name.value,start.value,end1.value)"></td>
	<tr>
</table></form>

<p>&nbsp; </p>
<p><br>
</p>
<table width="100%" border="1" height="50" align="center" cellpadding="1" cellspacing="0" bordercolor="#333333">
  <tr> 
    <td width="13%" height="45" > 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">Date</font></div>
    </td>
    <td width="5%" height="45"><font face="Arial, Helvetica, sans-serif" size="2">Job#</font></td>
    <td width="60%" height="45" ><font face="Arial, Helvetica, sans-serif" size="2">Description</font></td>
    <td width="4%" height="45" bgcolor="#00CCFF">
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Hrs</font></div>
    </td>
    <td width="5%" height="45" bgcolor="#3399CC">
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Bill 
        Hours</font></div>
    </td>
    <td width="4%" height="45" bgcolor="#0033FF">
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">OT</font></div>
    </td>
    <td width="4%" align="center" height="45" bgcolor="#0066CC"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Expense 
      Desc</font></td>
    <td width="5%" align="center" height="45" bgcolor="#3300CC"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Expense 
      Amount</font></td>
  </tr>
</table>
<iframe name="top" width="100%" height="150" src="timesheet.asp" scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
<IFRAME name="bottom" width="100%" height="150" src="timedetail.asp" scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
</body>
</html>
