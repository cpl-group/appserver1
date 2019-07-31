<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<html>
<head>
<script>
<%
		if isempty(Session("name")) then
%>
top.location="../index.asp"
<%
		else
			if Session("ts") < 4 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if	
user="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open application("cnnstr_main")
number="[Entry ID]"
table="[Job Log]"
job=Request("job")
temp=Request("day")
if not isempty(job) then
    Set rst3 = Server.CreateObject("ADODB.Recordset")
	sql3 = "SELECT description FROM [Job Log] where([Entry id]='"& job &"')"
 	rst3.Open sql3, cnn1, adOpenStatic, adLockReadOnly
	if not rst3.eof then
		description=rst3("description")
	end if
end if
sql = "SELECT [Entry id] FROM [Job Log] order by [Entry id]"

%>
function openpopup(){
//configure "Open Logout Window

top.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}


function setDesc(job, id){
    var date=document.forms[0].date.value
	if( id >0 ){
	    document.location="timedetail.asp?job="+job+"&day="+date+"&id="+id
	}else{
	    document.location="timedetail.asp?job="+job+"&day="+date
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
</script>

</head>
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
		
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")

cnn1.Open application("cnnstr_main")
contact=request("contact")
customer=request("customer")
%>
<body bgcolor="#FFFFFF" text="#000000" onload="setDate()">
<form name="form1" method="post" action="corptimesave.asp">

  <table width="100%" border="0" height="50" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="3%"><font face="Arial, Helvetica, sans-serif" size="2">User </font></td>
      <td width="2%"><font size="2"></font></td>
      <td width="5%"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">Date</font></div>
      </td>
      <td width="2%"><font size="2"></font></td>
      <td width="4%"><font face="Arial, Helvetica, sans-serif" size="2">Job#</font></td>
      <td width="25%"><font face="Arial, Helvetica, sans-serif" size="2">Description</font></td>
      <td width="4%"><font face="Arial, Helvetica, sans-serif" size="2">Hours</font></td>
      <td width="3%"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">Billable 
          Hours</font></div>
      </td>
      <td width="2%"><font face="Arial, Helvetica, sans-serif" size="2">OT</font></td>
      <td width="17%"><font face="Arial, Helvetica, sans-serif" size="2">Expense 
        Description </font></td>
      <td width="33%"><font face="Arial, Helvetica, sans-serif" size="2">Expense 
        Amount</font></td>
    </tr>
    <tr> 
      <td width="3%"> 
        <select name="user">
          <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select [last name]+', '+[first name]  as name, substring(username,7,20) as user1 from employees where active=1 order by [last name]"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
		%>
          <option value="<%=rst2("user1")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("name")%></font></option>
          <%
					rst2.movenext
					loop
					end if
					rst2.close
				%>
        </select>
        <input type="hidden" name="invday" value="<%=Request.Querystring("invday")%>">
        <input type="hidden" name="des" value="<%=Request.Querystring("des")%>">
        <input type="hidden" name="comment" value="<%=Request.Querystring("comment")%>">
		<input type="hidden" name="customer" value="<%=customer%>">
  		<input type="hidden" name="contact" value="<%=contact%>">
      </td>
      <td width="2%"> 
        <input type="button" name="minus" value=" -" onClick="navigate(this.value)">
      </td>
      <td width="5%" > 
        <input type="text" name="date" value="<%=temp%>" size="13%" >
        <font face="Arial, Helvetica, sans-serif" size="2"></font> </td>
      <td width="2%"> 
        <input type="button" name="plus" value="+" onClick="navigate(this.value)">
      </td>
      <td width="4%"> <%=job%>
        <input type="hidden" name="job" onChange="setDesc(this.value)" value="<%=job%>" size="10">
      </td>
      <td width="25%" > 
        <input type="text" name="description" value="<%=description%>" size="50%">
        <font face="Arial, Helvetica, sans-serif" size="2"></font> </td>
      <td width="4%"> 
        <div align="center"> 
          <input type="text" name="hrs" size="2%" value=0>
          <font face="Arial, Helvetica, sans-serif" size="2"></font> </div>
      </td>
      <td width="3%"> 
        <div align="center"> 
          <input type="text" name="billh" size="2%" value=0>
        </div>
      </td>
      <td width="2%"> 
        <input type="text" name="ot" size="2%" value=0>
        <font face="Arial, Helvetica, sans-serif" size="2"></font> </td>
      <td width="17%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="exp" size="40%">
        </font> </td>
      <td width="33%"> $ 
        <input type="text" name="value" value="0" size="5">
        <font face="Arial, Helvetica, sans-serif" size="2"></font> </td>
    </tr>
    <tr> 
      <input type="Submit" name="modify" value="Save" onClick="truncate(this.value)">
    </tr>
  </table>

</form>
</body>
</html>
