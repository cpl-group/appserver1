<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
window.close()
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("ts") < 4 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if	
user=Session("name")
uid="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
sql = "SELECT startweek, endweek FROM time_submission where username='payroll' "
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
if not rst1.eof then
  	startweek=rst1("startweek")
	endweek=rst1("endweek")
	end if


%>
<html>
<head>
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
function weekDay(d){
    var day
	if(d==0){
	    day="Sunday"
	}else if(d==1){
	    day="Monday"
	}else if(d==2){
	    day="Tuesday"
	}else if(d==3){
	    day="Wednesday"
	}else if(d==4){
		day="Thursday"
	}else if(d==5){
		day="Friday"
	}else{
		day="Saturday"
	}
	return day
}
function setUp(){
    var now=new Date()
	var s
	var e
	var str
	var day
    s=document.forms[0].startweek.value
    e=document.forms[0].endweek.value
	ary=s.split(" ")
	s=ary[ary.length-1]
    ary=e.split(" ")
	e=ary[ary.length-1]
	day=new Date(s)
	str=weekDay(day.getDay())+" "+s
	document.forms[0].startweek.value=str   
	day=new Date(e)
	str=weekDay(day.getDay())+" "+e
	document.forms[0].endweek.value=str   
}

function navigate(dir, flag, i){
    var str
    var currdate
	var startweek=document.forms[0].startweek.value
	var endweek=document.forms[0].endweek.value
	currdate1 = new Date(startweek)
	currdate2 = new Date(endweek)
	if(dir=="1"){
 	    currdate1 = new Date(currdate1).valueOf() + (i * 86400000)
	    currdate2 = new Date(currdate2).valueOf() + (i * 86400000)
	}else{
	    currdate1 = new Date(currdate1).valueOf() - (i * 86400000)
    	currdate2 = new Date(currdate2).valueOf() - (i * 86400000)
	}
	currdate= new Date(currdate1)
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
    str=new Date(currdate)
   	str=weekDay(str.getDay())
    currdate=str+" "+currdate
    startweek=currdate;
	currdate= new Date(currdate2)
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
    str=new Date(currdate)
   	str=weekDay(str.getDay())
    currdate=str+" "+currdate
    endweek=currdate;
	if(flag==0){
	  	document.forms[0].startweek.value=startweek
	}else if(flag==1){
	    document.forms[0].endweek.value=endweek
	}else{
		document.forms[0].startweek.value=startweek
		document.forms[0].endweek.value=endweek
	}   
}

function truncate(){
    var date=document.forms[0].startweek.value;
	date=date.split(" ")
	startweek=date[1]
	date=document.forms[0].endweek.value;
	date=date.split(" ")
	endweek=date[1]
    document.location="admintimesave.asp?startweek="+startweek+"&endweek="+endweek
	window.close()
}
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000" onload="setUp()">
<form name="form1" method="post" action="">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
        <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">TIMESHEETS 
          : CURRENT WEEK SETUP</font></b></font></div>
      </td>
    </tr>
  </table>
  <div align="center">
    <table width="100%" border="0">
      <tr> 
        <td width="50%">&nbsp;</td>
        <td width="50%">&nbsp;</td>
      </tr>
      <tr> 
        <td width="50%"><font face="Arial, Helvetica, sans-serif">From </font></td>
        <td width="50%"><font face="Arial, Helvetica, sans-serif">To</font> </td>
      </tr>
      <tr> 
        <td width="50%"> <font face="Arial, Helvetica, sans-serif"> 
          <input type="button" name="adjust2" value=" -"  onClick="navigate(0, 0, 1)">
          </font> 
          <input type="text" name="startweek" value="<%=startweek%>">
          <font face="Arial, Helvetica, sans-serif"> 
          <input type="button" name="adjust3" value="+"  onClick="navigate(1, 0, 1)">
          </font> </td>
        <td width="50%"> 
          <input type="button" name="adjust4" value=" -"  onClick="navigate(0, 1, 1)">
          <input type="text" name="endweek" value="<%=endweek%>">
          <input type="button" name="adjust" value="+" onClick="navigate(1, 1, 1)">
        </td>
      </tr>
      <tr> 
        <td width="50%"><font face="Arial, Helvetica, sans-serif"> 
          <input type="button" name="next2" value="PREVIOUS WEEK"  onClick="navigate(0, 2, 7)">
          </font></td>
        <td width="50%"><font face="Arial, Helvetica, sans-serif"> 
          <input type="button" name="next" value="NEXT WEEK"  onClick="navigate(1, 2, 7)">
          </font></td>
      </tr>
    </table>    <div align="left">
      <input type="button" name="Submit" value="SAVE" onClick="truncate()">
    </div>
  </div>
</form>

<br>

<p>&nbsp;</p>

</body>
</html>
