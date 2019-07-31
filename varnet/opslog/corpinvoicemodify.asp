<html>
<head>
<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
ReDim By(5)
By(0) = "None"
By(1) = "Entry"
By(2) = "Junior"
By(3) = "Mid"
By(4) = "Senior"
By(5) = "Admin"


%>
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
flag=Request("flag")
des=Request("description")
customer=Request("customer")
contact=Request("contact")
sql = "SELECT *, substring(matricola,5,20) as user1 FROM invoice_submission where id='"& id &"'"
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
function fillup(typebox){
	document.location="corpinvoicemodify.asp?typebox=" + typebox
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
	d=date.split(" ")
	date=d[1]
	document.forms[0].date.value=date
	
}
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000" onload="setDate()">
<form name="form1" method="post" action="corpinvoiceupdate.asp">
  <input type="hidden" name="id" value="<%=id%>">
  <input type="hidden" name="job" value="<%=job%>">
  <input type="hidden" name="flag" value="<%=flag%>">
  <input type="hidden" name="customer" value="<%=customer%>">
  <input type="hidden" name="contact" value="<%=contact%>">
  <table width="100%" border="0" height="50" align="center" cellpadding="0" cellspacing="0">
    <tr> 
	  <td width="3%"><font face="Arial, Helvetica, sans-serif" size="2">Category</font></td>
	  <td width="6%"><font face="Arial, Helvetica, sans-serif" size="2">UserName</font></td>
      <td width="1%">&nbsp;</td>
	  
	  <td width="5%"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">Date</font></div>
      </td>
	  <td width="1%">&nbsp;</td>
      <td width="28%"><font face="Arial, Helvetica, sans-serif" size="2">Description</font></td>
      <td width="1%"><font face="Arial, Helvetica, sans-serif" size="2">Hrs</font></td>
      <td width="1%"><font face="Arial, Helvetica, sans-serif" size="2">Bill H</font></td>
	  <td width="1%"><font face="Arial, Helvetica, sans-serif" size="2">OT</font></td>
	  <td width="4%"><font face="Arial, Helvetica, sans-serif" size="2">Expense</font></td>
      <td width="49%"><font face="Arial, Helvetica, sans-serif" size="2">Value</font></td>
  </tr>
  <%
  if not rst1.eof then
 
  %>
  <tr>
      <td width="3%"> <font face="Arial, Helvetica, sans-serif"> 
        <select name="typebox" size="1" >
          <%
		  i=0
		  while i < 6 
		  	if i = rst1("category") then 
				%>
				<option value="<%=i%>" selected><%=By(i)%></option>
				<%
			else
			%>
				<option value="<%=i%>"><%=By(i)%></option>
			<%
			end if
			i=i+1
		  wend
		  %>
		  </select>
       </font></td>
	   
      <td width="6%"> <font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("user1")%> 
        <input type="hidden" name="username" size="13" value="<%=rst1("user1")%>">
        </font> 
      </td>
      <td width="1%"> 
        <input type="button" name="minus" value=" -" onClick="navigate(this.value)">
      </td>
      <td width="5%"> 
        <input type="hidden" name="buffer">
        <input type="text" name="date" value="<%=temp%>" size="13%">
        <font face="Arial, Helvetica, sans-serif" size="2"></font>
      </td>
      <td width="1%"> 
        <input type="button" name="plus" value="+" onClick="navigate(this.value)">
      </td>
      <td width="28%" > 
        <input type="text" name="description" value="<%=rst1("description")%>" size="70%" maxlength="100">
        <font face="Arial, Helvetica, sans-serif" size="2"></font> 
      </td>
      <td width="1%"> 
        <input type="text" name="hours" size="2%" value="<%=rst1("hours")%>"><font face="Arial, Helvetica, sans-serif" size="2"></font>
	</td>
	  <td width="1%"> 
        <input type="text" name="billh" size="2%" value="<%=rst1("hours_bill")%>">
        <font face="Arial, Helvetica, sans-serif" size="2"></font> 
      </td>
      <td width="1%"> 
        <input type="text" name="overt" size="2%" value="<%=rst1("overt")%>"><font face="Arial, Helvetica, sans-serif" size="2"></font>
	</td>
      <td width="4%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
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
      <td width="49%"> 
        <input type="text" name="v" value="<%=value%>" size="4%"><font face="Arial, Helvetica, sans-serif" size="2"></font>
	</td>
  </tr>
  	
  <%
  end if
  
  %>
  
</table>
<br>
<input type=hidden name=des value="<%=des%>">
<input type="Submit" name="modify" value="Update" onClick="truncate(date.value)">
</form>
</body>
</html>
