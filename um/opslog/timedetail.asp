<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
cnn1.Open getConnect(0,0,"intranet")
number="id"
table="MASTER_JOB"
job=Request("job")
temp=Request("day")
if not isempty(job) then
    Set rst3 = Server.CreateObject("ADODB.Recordset")
	
	'sql3 = "SELECT description FROM [Job Log] where([Entry id]='"& job &"' and [current status]<>'Closed')"
	
   sql3 = "SELECT description FROM MASTER_JOB where(id='"& job &"' and status <>'Closed')"

 	rst3.Open sql3, cnn1, adOpenStatic, adLockReadOnly
	if rst3.eof then
		description="NA"
		job="NA"
		else
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
<body bgcolor="#FFFFFF" text="#000000" onload="setDate()">
<form name="form1" method="post" action="timemodify.asp">

<table width="100%" border="0" height="50" align="center" cellpadding="0" cellspacing="0">
  <tr> 
      <td width="4%">&nbsp;</td>
	  <td width="10%"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">Date</font></div>
      </td>
	  <td width="4%">&nbsp;</td>
      <td width="9%"><font face="Arial, Helvetica, sans-serif" size="2">Job#</font></td>
      <td width="48%"><font face="Arial, Helvetica, sans-serif" size="2">Description</font></td>
      <td width="3%"><font face="Arial, Helvetica, sans-serif" size="2">Hrs</font></td>
      <td width="3%"><font face="Arial, Helvetica, sans-serif" size="2">OT</font></td>
      <td width="12%"><font face="Arial, Helvetica, sans-serif" size="2">Expense 
        Description </font></td>
      <td width="7%"><font face="Arial, Helvetica, sans-serif" size="2">Expense 
        Amount</font></td>
  </tr>
  <%
  if isEmpty(Request("id")) then	
  %>
  <tr>
      <td width="4%"> 
        <input type="button" name="minus" value=" -" onClick="navigate(this.value)">
      </td>
      <td width="10%" > 
        <input type="text" name="date" value="<%=temp%>" size="13%" >
        <font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
      </td>
      <td width="4%"> 
        <input type="button" name="plus" value="+" onClick="navigate(this.value)">
      </td>
      <td width="9%"> 
	  <%if description = "NA" and job="NA" then 
	  	description=""
		job=""
	  %>
	  <script> var temp="SELECTED JOB HAS BEEN CLOSED. NO TIME CAN BE POSTED"; alert(temp) </script>
	  <% end if %>
        <input type="text" name="job" onChange="setDesc(this.value)" value="<%=job%>" size="10">
      </td>
      <td width="48%" > 
        <input type="text" name="description" value="<%=description%>" size="60%"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
      <td width="3%"> 
        <input type="text" name="hrs" size="2%" value=0><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
      <td width="3%"> 
        <input type="text" name="ot" size="2%" value=0><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
      <td width="12%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="exp" size="10%">
        </font> </td>
      <td width="7%"> $
<input type="text" name="value" value="0" size="4%"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
  </tr>
  <tr>
    <input type="Submit" name="modify" value="Save" onClick="truncate(this.value)">
  </tr>	
  <%
  else
  %>
  <tr>
  <%
    id=Request("id")
	sql2 = "SELECT *, matricola AS Expr1 FROM Times WHERE (matricola = '"& user &"') and (id='"& id &"')"

	Set rst2 = Server.CreateObject("ADODB.Recordset")
	rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
	if not rst2.EOF then
  %>
      <td width="4%"> 
        <input type="hidden" name="id" value="<%=id%>">
        <input type="button" name="minus" value=" -" onClick="navigate(this.value)">
		
	</td>
      <td width="10%" > 
        <input type="text" name="date" size="13%" value="<%=rst2("date")%>">
      
	  
    <font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
	  <td width="4%"> 
        <input type="button" name="plus" value="+" onClick="navigate(this.value)">
	</td>
      <td width="9%"> 
	  <%
	    if not isempty(Request("job")) then
		    
         	sql = "SELECT description FROM master_job where(id='"& job &"' and status <>'Closed')"
 			rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
			if rst1.EOF then
				description=""
				job=""			
			else
				description=rst1("description")
				job=Request("job")
			end if
		else
		    job=rst2("jobno")
			description=rst2("description")
		end if
		
		if description = "NA" and job="NA" then 
		description=""
		job=""
	  %>
	  <script> var temp="SELECTED JOB HAS BEEN CLOSED. NO TIME CAN BE POSTED"; alert(temp) </script>
	  <% end if %>
        <input type="text" name="job"  value="<%=job%>" onChange="setDesc(this.value, <%=id%>)" size="10">
      </td>
      <td width="48%" > 
        <input type="text" name="description" size="60%" value="<%=description%>"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
      <td width="3%"> 
	    <%
		    if isNull(rst2("overt")) Then
		        rst2("overt")=0
		    end if
		    if isNull(rst2("hours")) Then
		        rst2("hours")=0
		    end if    
		%>
        
        <input type="text" name="hrs" size="2%" value="<%=rst2("hours")%>"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
      <td width="3%"> 
<input type="text" name="ot" size="2%" value="<%=rst2("overt")%>"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
      <td width="12%"> 
        <input type="text" name="exp" size="10%" value="<%=rst2("expense")%>"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
      <td width="7%"> 
        <%
	    if (rst2("value")>0) then
		    value=rst2("value")
		else
		    value=0
		end if
	  %>
        <input type="text" name="value" size="4%" value="<%=value%>"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font>
	</td>
    </tr>
	<tr>
    <input type="Submit" name="modify" value="Update" onclick="truncate(this.value)">
	  <input type="button" name="Button" value="Close" onclick="Javascript:document.location='timedetail.asp'">
    </tr>
    <%
    end if
    %>
	
<%
  end if
%>


</table>

</form>
</body>
</html>
