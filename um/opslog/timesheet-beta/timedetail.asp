<%@Language="VBScript"%>
<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<%
'12/19/2007 N.Ambo added change to allow user to post time for unstarted jobs which before was not allowed
dim user, cnn1, rst1, sql, startweek, endweek, add1, add2
user="ghnet\"&trim(request("name"))

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")

sql="select username, startweek, endweek from user_cost where username='" & user & "'"
rst1.Open sql, cnn1, 0, 1, 1
if not rst1.eof then
	startweek=rst1("startweek")
	endweek=rst1("endweek")
end if
rst1.close

dim sql666,rst666,CompanyCheck
		sql666 = "select company from employees where username='ghnet\"&getXMLUserName()&"'"
		'response.write sql666
		Set rst666 = Server.CreateObject("ADODB.Recordset")
		rst666.Open sql666, cnn1', adOpenStatic, adLockReadOnly
		if not rst666.eof then
		CompanyCheck=rst666("Company")
		rst666.close
		end if
		
dim number, table, job, temp, rst3, sql3, description, id, status, rfp, oldjob
oldjob = trim(request("oldjob"))
number="id"
table="MASTER_JOB"
job=Request("job")
temp=Request("day")
description= ""

dim bLoadDescriptionMasterTable
bLoadDescriptionMasterTable = true

' if id is supplied that means user is editing an existing time sheet entry.
' if job is not supplied that means the user is just starting to edit it and hasnt 
' input any info yet.  if this is the case we have to lookup the job from the Times table
if (not isEmpty(Request("id"))) AND isempty(job) then
	id=Request("id")
	sql3 = "SELECT *, matricola AS Expr1 FROM Times WHERE (id='"& id &"')"
	Set rst3 = Server.CreateObject("ADODB.Recordset")
	rst3.Open sql3, cnn1, adOpenStatic, adLockReadOnly
	job=rst3("jobno")
	description=rst3("description")
	bLoadDescriptionMasterTable = false
end if

if not isNumeric(job) then						' ensure that the info we have been provided for job is numeric.
	job=""
	%>
	<script>alert("Job must be a numeric value.")</script>
	<%
elseif not isempty(job) then					' if it is numeric and not empty, then we can look up info about the job
	Set rst3 = Server.CreateObject("ADODB.Recordset")	
	sql3 = "SELECT * FROM MASTER_JOB where(id='"& job &"')"
 	rst3.Open sql3, cnn1, adOpenStatic, adLockReadOnly
	
	' now we have to verify that the job exists and time can be posted, if everything goes well
	' we can set the description based on the lookup info.  otherwise we will clear job so that 
	' the bad info is not seen in the page.

	if rst3.eof then
		%>
		<script>alert("Job <%=job%> could not be found.")</script>
		<%
		job=""		
	elseif lcase(trim(rst3("Status"))) = "closed" then
		%>
		<script>alert("Job <%=job%> is closed, no time can be posted.")</script>
		<%
		job=""
	'12/19/2007 N.Ambo blocked off follwing lines because time can now be posted for unstarted jobs
	'elseif lcase(trim(rst3("Status"))) = "unstarted" and not rst3("rfp") and trim(rst3("company")) <> "GE" then
		%>
		<script>//alert("Job <%=job%> is unstarted, no time can be posted.")</script>
		<%
		'job=""
		
		
		elseif lcase(trim(rst3("Company"))) <>  lcase(trim(CompanyCheck)) and not allowGroups("Timesheet Supervisors") then 
		dim jobcompany 
		jobcompany=rst3("Company")
		%>
		<script>alert("You're not allowed to post time to a <%=jobcompany%> Jobno only to a <%=CompanyCheck%> Jobno")</script>
		<%
		job=""
		
	elseif bLoadDescriptionMasterTable then
		description=rst3("description")
		add1 = rst3("address_1")
		add2 = rst3("address_2")			
	end if
end if
if oldjob="" then oldjob=job
sql = "SELECT [Entry id] FROM [Job Log] order by [Entry id]"
%>

<script language="JavaScript" type="text/javascript">
<%
if isempty(getKeyValue("name")) then
	%>
	top.location="http://www.genergyonline.com"
	<%
end if
%>

function openpopup(){
//configure "Open Logout Window

	top.document.location.href="../index.asp";
}

function loadpopup(){
	openpopup()
}

function jobPicked(job){
	document.form1.job.value = job
}

function setDesc(job, id, name){
	var date=document.forms[0].date.value
	if( id > 0 ){
		document.location="timedetail.asp?job="+job+"&oldjob=<%=oldjob%>&day="+date+"&name="+name+"&id="+id+"&source=<%=request("source")%>"
	}else{
		document.location="timedetail.asp?job="+job+"&oldjob=<%=oldjob%>&day="+date+"&name="+name+"&source=<%=request("source")%>"
	}
}
function weekDay(d){
    var day
	if(d==0){
	    day="Sun"
	}else if(d==1){
	    day="Mon"
	}else if(d==2){
	    day="Tue"
	}else if(d==3){
	    day="Wed"
	}else if(d==4){
		day="Thu"
	}else if(d==5){
		day="Fri"
	}else{
		day="Sat"
	}
	return day
}
function setDate(){
    var now=new Date()
	var temp=""
	var str
	var day
	if(document.forms[0].date.value==""){
		temp=now.getMonth() + 1+ "/" + (now.getDate()) + "/" + now.getFullYear()
    }else{
	    temp=document.forms[0].date.value
		ary=temp.split(" ")
		temp=ary[ary.length-1]
    }    
	day=new Date(temp)
	str=weekDay(day.getDay())+" "+temp
	document.forms[0].date.value=str   
}

function navigate(direc){
	var str
    datevalue=document.forms[0].date.value
   	var currdate = new Date(datevalue)
	if (direc == "+") {
	    currdate=new Date(currdate).valueOf() + (1 * 90000000)
	}else{
	    currdate=new Date(currdate).valueOf() - (1 * 86400000)
	}
	currdate = new Date(currdate)
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	str=new Date(currdate)
	str=weekDay(str.getDay())
	currdate=str+" "+currdate
	document.forms[0].date.value=currdate;
}

function truncate(){
    var date=document.forms[0].date.value;
	date=date.split(" ")
	document.forms[0].date.value=date[1]
	
	
}

function fillin(wkday, week){
  str = weekDay(wkday-1);
  document.forms[0].date.value = str + " " + week;
}
function openwin(url,mwidth,mheight){
  window.name="opener";
  popwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth);
  popwin.focus();
}
function delete1(key,u){
	if(confirm("Are you sure you want to delete this entry?")){
	document.location="deletetime.asp?key="+key+"&window=close&u="+u
	}
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#eeeeee" text="#000000" onload="setDate()" style="border-top:2px outset #ffffff;" class="innerbody">
<form name="form1" method="post" action="timemodify.asp">
<% 
if isEmpty(Request("id")) then			'creating a new listing
	%>
	<table border=0 cellpadding="2" cellspacing="0" width="100%">
		<tr>
			<td bgcolor="#dddddd" style="border-bottom:1px solid #999999;">
				<table border=0 cellpadding="2" cellspacing="0">
					<tr>
						<td><b>Date</b> (<%=startweek%> - <%=endweek%>):</td>
						<td>&nbsp;<input type="text" name="date" value="<%=temp%>" size="20">&nbsp;</td>
 						<!--<td><input type="button" name="minus" value=" -" onClick="navigate(this.value)" 
						style="width:20px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><input
						 type="button" name="plus" value="+" onClick="navigate(this.value)" style="width:20px;background
						 -color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>-->
    					<td>
    						<table border=0 cellpadding="0" cellspacing="0">
    							<tr>
      								<td>
										<input type="button" name="W" value="<%=shortDay(weekday(startweek))%>"
										onClick="fillin('<%=weekday(startweek)%>', '<%=startweek%>')"
										style="width:30px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
									</td>
      								<td>
										<input type="button" name="Th" value="<%=shortDay(weekday(DateAdd("d",1,startweek)))%>"
										onClick="fillin('<%=weekday(DateAdd("d",1,startweek))%>', '<%=DateAdd("d",1,startweek)%>')"
										style="width:30px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
									</td>
									<td><input type="button" name="F" value="<%=shortDay(weekday(DateAdd("d",2,startweek)))%>"
										onClick="fillin('<%=weekday(DateAdd("d",2,startweek))%>', '<%=DateAdd("d",2,startweek)%>')"
										style="width:30px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
									</td>
									<td><input type="button" name="Sa" value="<%=shortDay(weekday(DateAdd("d",3,startweek)))%>" 
										onClick="fillin('<%=weekday(DateAdd("d",3,startweek))%>', '<%=DateAdd("d",3,startweek)%>')"
										style="width:30px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
									</td>
									<td><input type="button" name="Su" value="<%=shortDay(weekday(DateAdd("d",4,startweek)))%>"
										onClick="fillin('<%=weekday(DateAdd("d",4,startweek))%>', '<%=DateAdd("d",4,startweek)%>')"
										style="width:30px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
									</td>
      								<td><input type="button" name="M" value="<%=shortDay(weekday(DateAdd("d",5,startweek)))%>"
										onClick="fillin('<%=weekday(DateAdd("d",5,startweek))%>', '<%=DateAdd("d",5,startweek)%>')"
										style="width:30px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
									</td>
									<td><input type="button" name="Tu" value="<%=shortDay(weekday(DateAdd("d",6,startweek)))%>"
										onClick="fillin('<%=weekday(DateAdd("d",6,startweek))%>', '<%=DateAdd("d",6,startweek)%>')"
										style="width:30px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table border=0 cellpadding="0" cellspacing="2">
          <tr valign="bottom"> 
            <td>Job Number</td>
            <td>&nbsp;</td>
            <td>Description</td>
            <td>Hrs</td>
            <td>OT</td>
            <td>Expense Description</td>
            <td>Expense Amount</td>
          </tr>
          <tr> 
            <td> <input type="text" name="job" onChange="setDesc(this.value,'<%=id%>','<%=trim(request("name"))%>')" value="<%=job%>" size="6"> 
              <input type="hidden" name="oldjob" value="<%=oldjob%>"> </td>
            <td>&nbsp;</td>
            <td><input type="text" name="description" value="<%=description%>" size="35"></td>
            <td><input type="text" name="hrs" size="2%" value=0></td>
            <td><input type="text" name="ot" size="2%" value=0></td>
            <td><input type="text" name="exp" size="10%"></td>
            <td> $
              <input type="text" name="value" value="0" size="4%"></td>
          </tr>
          <%if add1 <> "" then %>
		  <tr>
            <td></td>
            <td>&nbsp;</td>
            <td colspan=5><%=add1%>, <%=add2%></td>
          </tr>
          <%end if%>
		  <tr> 
            <td colspan="7"> <input type="Submit" name="modify" value="Save" onClick="truncate(this.value)"
							  style="border:1px outset #ddffdd;background-color:ccf3cc;font-weight:bold;"> 
              <%
							dim CancelOnClick
							if trim(request("source"))="review" then 
								CancelOnclick = "Javascript:window.close()"
							else
								CancelOnclick = "Javascript:document.location='timedetail.asp?name="&trim(request("name"))&"'"
							end if
							%>
              <input type="button" name="Button" value="Cancel" onclick="<%=CancelOnclick%>"
							  style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
              &nbsp;&nbsp;<img src="/genergy2/setup/images/aro-rt.gif" border="0">&nbsp; 
              <a href="javascript:openwin('joblist.asp?name=<%=request("name")%>',260,320);">Quick 
              job search</a> </td>
          </tr>
        </table>
			</td>
		</tr>
	</table>
<%
else			'request id is not empty, this means we are editing an existing entry
	id=Request("id")
	
	dim sql2, rst2
	
	sql2 = "SELECT *, matricola AS Expr1 FROM Times WHERE (id='"& id &"')"

	Set rst2 = Server.CreateObject("ADODB.Recordset")
	rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
	if not rst2.EOF then
	if trim(job)="" then job = rst2("jobno")
		%>
		<table border=0 cellpadding="2" cellspacing="0" width="100%">
			<tr>
				<td bgcolor="#dddddd" style="border-bottom:1px solid #999999;">
					<table border=0 cellpadding="2" cellspacing="0">
						<tr>
							<td><b>Date</b>:</td>
							<td><input type="button" name="minus" value=" -" onClick="navigate(this.value)"
								style="width:20px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
							</td>
							<td>&nbsp;<input type="text" name="date" value="<%=rst2("date")%>" size="20"></td>
							<td><input type="button" name="plus" value="+" onClick="navigate(this.value)"
								style="width:20px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td>
					<table border=0 cellpadding="0" cellspacing="2">
						<tr> 
							<td>Job Number</td>
							<td>&nbsp;</td>
							<td>Description</td>
							<td>Hrs</td>
							<td>OT</td>
							<td>Expense Description</td>
							<td>Expense Amount</td>
						</tr>
						<tr>
							<td> 
								<input type="text" name="job"  value="<%=job%>" onChange="setDesc(this.value, '<%=id%>','<%=trim(request("name"))%>')" size="6">
    						<input type="hidden" name="oldjob" value="<%=oldjob%>">
							</td>
							<td>&nbsp;</td>
							<td><input type="text" name="description" size="35" value="<%=description%>"></td>
							<td> 
								<%
								if isNull(rst2("overt")) Then
									rst2("overt")=0
								end if
								if isNull(rst2("hours")) Then
									rst2("hours")=0
								end if    
								%>
								<input type="text" name="hrs" size="2%" value="<%=rst2("hours")%>">
							</td>
							<td><input type="text" name="ot" size="2%" value="<%=rst2("overt")%>"></td>
							<td><input type="text" name="exp" size="10%" value="<%=rst2("expense")%>"></td>
							<td> 
								<%
								dim value
								if (rst2("value")>0) then
									value=rst2("value")
								else
									value=0
								end if
								%>
								$<input type="text" name="value" size="4%" value="<%=value%>">
								<input type="hidden" name="id" value="<%=id%>">
							</td>
						</tr>
						<tr>
							<td colspan="9">
								<input type="Submit" name="modify" value="Update" onclick="truncate(this.value)"
								  style="border:1px outset #ddffdd;background-color:ccf3cc;font-weight:bold;">
								<%
								if trim(request("source"))="review" then 
									CancelOnclick = "Javascript:window.close()"
								else
									CancelOnclick = "Javascript:document.location='timedetail.asp?name="&trim(request("name"))&"'"
								end if
								%>
								<input type="button" name="Button" value="Cancel" onclick="<%=CancelOnclick%>"
								  style="border:1px outset #ddffdd;background-color:ccf3cc;">
								  <a href="javascript:delete1('<%=id%>','0')" style="background-color:#eeeeee;"><img src="/um/opslog/delete.gif" align="absmiddle" border="0"></a>
								  
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	<%
	end if				'matches if not rst2.eof
end if
%>
<input name="source" type="hidden" value="<%=trim(request("source"))%>">
<input name="name" type="hidden" value="<%=trim(request("name"))%>">
</form>
<%
function shortDay(someday)
	select case someday
		case 1
			shortDay = "Su"
		case 2
			shortDay = "M"
		case 3
			shortDay = "Tu"
		case 4
			shortDay = "W"
		case 5
			shortDay = "Th"
		case 6
			shortDay = "F"
		case 7
			shortDay = "Sa"
	end select
end function 
%>
</body>
</html>
