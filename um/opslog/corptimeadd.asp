<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<script language="JavaScript" type="text/javascript">
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
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
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

cnn1.Open  getConnect(0,0,"intranet")
contact=request("contact")
customer=request("customer")
%>
<body bgcolor="#eeeeee" text="#000000" onload="setDate()" style="border-top:1px solid #999999;">
<form name="form1" method="post" action="corptimesave.asp">
<table border=0 cellpadding="3" cellspacing="1">
<tr>
  <td colspan="2"><b>Add Time</b></td>
</tr>
<tr>
  <td>User:</td>
  <td> 
  <select name="user">
  <%Set rst2 = Server.CreateObject("ADODB.recordset")
  sqlstr = "select [last name]+', '+[first name]  as name, substring(username,7,20) as user1 from employees where active=1 order by [last name]"
  rst2.Open sqlstr, cnn1, 0, 1, 1
  if not rst2.eof then
  do until rst2.eof
  %>
  <option value="<%=rst2("user1")%>"><%=rst2("name")%></option>
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
</tr>
</table>
  <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#cccccc">
    <tr bgcolor="#dddddd"> 
      <td>Date</td>
      <td>Job#</td>
      <td>Description</td>
      <td>Hours</td>
      <td>Billable Hours</td>
      <td>OT</td>
      <td>Expense Description </td>
      <td>Expense Amount</td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td> 
        <input type="button" name="minus" value=" -" onClick="navigate(this.value)">
        <input type="text" name="date" value="<%=temp%>" size="13%" >
        <input type="button" name="plus" value="+" onClick="navigate(this.value)">
      </td>
      <td> <%=job%>
        <input type="hidden" name="job" onChange="setDesc(this.value)" value="<%=job%>" size="10">
      </td>
      <td> 
        <input type="text" name="description" value="<%=description%>" size="35">
         </td>
      <td> 
        <div align="center"> 
          <input type="text" name="hrs" size="2%" value=0>
           </div>
      </td>
      <td> 
        <div align="center"> 
          <input type="text" name="billh" size="2%" value=0>
        </div>
      </td>
      <td> 
        <input type="text" name="ot" size="2%" value=0>
         </td>
      <td>  
        <input type="text" name="exp" size="20">
         </td>
      <td> $ 
        <input type="text" name="value" value="0" size="5">
         </td>
    </tr>
    
  </table>
<div style="padding:3px;">
<input type="Submit" name="modify" value="Save" onClick="truncate(this.value)" style="border:1px outset #ddffdd;background-color:ccf3cc;font-weight:bold;">
<input type="button" name="cancel" value="Cancel" onClick="history.back()" style="border:1px outset #ddffdd;background-color:ccf3cc;">
</div>
</form>
</body>
</html>
