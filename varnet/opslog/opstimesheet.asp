<html>
<head>
<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
//top.location="../index.asp"

</script>
<%
'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("ts") < 4 then 
				'Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				'Response.Redirect "../main.asp"
			end if	
		end if		
user="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open application("cnnstr_main")
job=Request("job")
id=Request("id")
temp=Request("day")
sql = "SELECT * FROM invoice_submission where id='"& id &"'"
rst1.Open sql, cnn1, 0, 1, 1
%>
<script>
var c=0
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
	if( typeof(document.forms[0].date) == "undefined" ){return}
	else{
	temp=document.forms[0].date.value
	ary=temp.split(" ")
	temp=ary[ary.length-1]
    day=new Date(temp)
	str=weekDay(day.getDay())+" "+temp
	document.forms[0].date.value=str   
	}
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

function truncate(date){
    //var date=document.forms[0].date.value;
	d=date.split(" ")
	date=d[1]
	document.forms[0].date.value=date
	
}
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000" onload="setDate()">
<form name="form1" method="post" action="opstimemodify.asp">
  <input type="hidden" name="id" value="<%=id%>">
  <input type="hidden" name="job" value="<%=job%>">
  <table width="100%" border="0" height="50" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="4%">&nbsp;</td>
	  
	  <td width="10%"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">Date</font></div>
      </td>
	  <td width="4%">&nbsp;</td>
      <td width="48%"><font face="Arial, Helvetica, sans-serif" size="2">Description</font></td>
      <td width="3%"><font face="Arial, Helvetica, sans-serif" size="2">Hrs</font></td>
      <td width="3%"><font face="Arial, Helvetica, sans-serif" size="2">Bill H</font></td>
	  <td width="3%"><font face="Arial, Helvetica, sans-serif" size="2">OT</font></td>
	  <td width="12%"><font face="Arial, Helvetica, sans-serif" size="2">Expense</font></td>
      <td width="7%"><font face="Arial, Helvetica, sans-serif" size="2">Value</font></td>
  </tr>
  <%
  if not rst1.eof then
  %>
  <tr>
      <td width="4%"> 
        <input type="button" name="minus" value=" -" onClick="navigate(this.value)">
      </td>
      <td width="10%" > 
	    <input type="hidden" name="buffer">
        <input type="text" name="date" value="<%=temp%>" size="13%" >
        <font face="Arial, Helvetica, sans-serif" size="2"></font>
      </td>
      <td width="4%"> 
        <input type="button" name="plus" value="+" onClick="navigate(this.value)">
      </td>
      <td width="48%" > 
        <input type="text" name="description" value="<%=rst1("description")%>" size="70%" maxlength="100">
        <font face="Arial, Helvetica, sans-serif" size="2"></font> 
      </td>
      <td width="3%"> 
        <input type="text" name="hours" size="2%" value="<%=rst1("hours")%>"><font face="Arial, Helvetica, sans-serif" size="2"></font>
	</td>
	<td width="3%"> 
        <input type="text" name="billh" size="2%" value="<%=rst1("hours_bill")%>"><font face="Arial, Helvetica, sans-serif" size="2"></font>
	</td>
      <td width="3%"> 
        <input type="text" name="overt" size="2%" value="<%=rst1("overt")%>"><font face="Arial, Helvetica, sans-serif" size="2"></font>
	</td>
      <td width="12%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
	    <%
		if(rst1("value") >=0) then
		    value=rst1("value")
		else
		    value=0
		end if
		if(rst1("billable") >=0) then
		    billable=rst1("billable")
		else
		    billable=0
		end if
		%>
        <input type="text" name="expense" value="<%=rst1("expense")%>" size="10%">
        </font> </td>
      <td width="7%"> 
        <input type="text" name="v" value="<%=value%>" size="4%"><font face="Arial, Helvetica, sans-serif" size="2"></font>
	</td>
  </tr>
  	
  <%
  end if
  %>
  
</table>
<br>
<input type="Submit" name="modify" value="Update" onClick="truncate(date.value)">
</form>
</body>
</html>
