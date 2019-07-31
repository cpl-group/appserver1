<%@Language="VBScript"%>
<%option explicit%>
<html>
<head>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->		<%
if isempty(getKeyValue("name")) then
	%>
	<script>
	top.location="http://www.genergyonline.com"
	
	</script>
	<%
end if

dim user, fullname, employee
user=getKeyValue("user")
employee=trim(request("employee"))
if employee<>"" then user=employee
fullname=getKeyValue("fullname")

Dim cnnMain, rs, sqlstr,cnn

Set cnnMain = Server.CreateObject("ADODB.connection")
Set cnn = Server.CreateObject("ADODB.connection")

Set rs = Server.CreateObject("ADODB.recordset")

cnnMain.Open getConnect(0,0,"intranet")
cnn.Open getConnect(0,0,"dbCore")

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

function setUp(name){
	var temp="timesetup.asp?name=" + name
    window.event.returnValue = false;
	//window.name = "opener"
	window.tsbottom.location.replace = "https://appserver1.genergy.com/um/opslog/timesheet-beta/timesetup.asp?name=" + name;
    return false;
    
   
}

function Print(name){
	var temp="timeformatemp.asp?name=" + name
	window.open(temp,"", "scrollbars=yes, width=700, height=500, resizeable=1, status=1,toolbar=0" );
}
function timesheetjob(name,start,end1){
    
	name=document.forms[0].name.value
	//start=document.forms[0].start.value
	if (confirm("You are posting time for the week starting " + start + " and ending " + end1 + ".  Today's date is <%=date()%>.  Is this correct?"))
    {
	window.event.returnValue = false;	
    //document.location="processtime.asp?name=" + name+ "&start=" +start + "&end1=" + end1
    window.location.replace = "https://appserver1.genergy.com/um/opslog/timesheet-beta/processtime.asp?name=" + name+ "&start=" +start + "&end1=" + end1;
    return false;
	}
}

function edittimesheet(name){
	document.frames[0].location ='timesheet.asp?name='+name
	document.frames[1].location ='timedetail.asp?name='+name
	document.form1.name.value = name
}
</script>

<STYLE>
	<!--
	A.ssmItems:link		{color:black;text-decoration:none;}
	A.ssmItems:hover	{color:black;text-decoration:none;}
	A.ssmItems:active	{color:black;text-decoration:none;}
	A.ssmItems:visited	{color:black;text-decoration:none;}
	//-->
</STYLE>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">		

</head>
<body bgcolor="#eeeeee" text="#000000">
<form name="form1" method="post" action="">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
	<tr>
		<td bgcolor="#6699cc"><span class="standardheader">Time Log: <%=Fullname%></span></td>
	</tr>

	<%
	sqlstr	 = "select startweek as s,endweek as e, username from user_cost where substring(username,7,20)='"&user&"'"
	rs.Open sqlstr, cnnMain, 0, 1, 1
	'response.write "<tr><td>"&sqlstr&"</td></tr>"
	%>
	
	<tr> 
		<td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;">
		
			<button name="Submit" value="Set Up Week" onClick="return setUp(document.form1.name.value)" 
				style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;" class="standard">
				<img src="/um/opslog/images/setup.gif" align="absmiddle" hspace="3" border="0">&nbsp;Set Up Week
			</button>
			
			<button name="Submit" value="Print Time Sheet" onClick="Print(document.form1.name.value)" 
				style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;" class="standard">
				<img src="/um/opslog/images/printer.gif" align="absmiddle" hspace="3" border="0">&nbsp;Print Time Sheet
			</button>
			
			<button name="submit1" value="Submit Time Sheet" onClick="return timesheetjob( name.value,start.value,end1.value)" 
				style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;font-weight:bold;" class="standard">
				<img src="/um/opslog/images/check.gif" align="absmiddle" hspace="3" border="0">&nbsp;Submit Time Sheet
			</button>
			
			<input type="hidden" name="name" value="<%=user%>">	  
			<input type="hidden" name="start" value="<%=rs("s")%>">
			<input type="hidden" name="end1" value="<%=rs("e")%>">		<%
			
			
			if allowgroups("Department Supervisors,NYE User,Genergy_Corp") then 
							
				dim usersRS
				
				set UsersRS = server.createobject("ADODB.recordset")
				UsersRS.open "select * from adusers_genergyusers where email is not null order by company,department,fullname", cnn
				UsersRS.MoveFirst
				if employee= "" then employee = user
				
				GenerateUserList "employee",UsersRS,"","",employee

				set UsersRS = nothing	%>
				<button name="submit1" value="viewtimesheet" onClick="this.form.submit();" 
					style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;font-weight:bold;" class="standard">
					<img src="/um/opslog/images/go.gif" align="absmiddle" hspace="3" border="0">&nbsp;Edit User
				</button>		<%
				
			end if %> 
		</td>
	<tr>
</table>
<iframe name="tstop" width="100%" height="280" src="timesheet.asp?name=<%=user%>" scrolling="auto" marginwidth="8" marginheight="16" border=0 frameborder=0></iframe> 
<IFRAME name="tsbottom" width="100%" height="125" src="timedetail.asp?name=<%=user%>" scrolling="auto" marginwidth="8" marginheight="16" border=0 frameborder=0></iframe> 
</form>
</body>
</html>
