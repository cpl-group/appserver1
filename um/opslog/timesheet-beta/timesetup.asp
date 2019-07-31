<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
		if isempty(getKeyValue("name")) then
%>
<script>
top.location="http://www.genergyonline.com"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if	
user=trim(request("name"))
uid="ghnet\"&trim(request("name"))

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")
sql = "SELECT name, startweek, endweek FROM user_cost where username='"& uid &"' "
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
if not rst1.eof then
  	startweek=rst1("startweek")
	endweek=rst1("endweek")
	end if


%>
<script langauge="JavaScript" type="text/javascript">
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
 	    currdate1 = new Date(currdate1).valueOf() + (i * 90000000)
	    currdate2 = new Date(currdate2).valueOf() + (i * 90000000)
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
//11/26/2007 N.Ambo modified function to include validity check for pay period dates; should be seven days with start day Wednesday and end day Wednesday
	var diff
	var startd
	var endd
	var day
	
	var date=document.forms[0].startweek.value;
	var date2=document.forms[0].endweek.value;
	date=date.split(" ")
	startweek=date[1]
	//date=document.forms[0].endweek.value;
	date2=date2.split(" ")
	endweek=date2[1]	
	day = 1000*60*60*24

	
	startd = new Date(startweek)
	endd = new Date(endweek)	
	diff = Math.round((endd.getTime()- startd.getTime())/day)
		
    //alert(diff)
	//if(date2[0]!="Tuesday" || date[0]!="Wednesday" || diff!=6){    --5/16/08 N.Ambo removed specific constraint for days since user cna have a different work week
	if(diff!=6){
		alert("Either the start or end date is invalid. Please verify that the dates are within the seven day time period.")
	}
	else{		
		document.location="timesave.asp?startweek="+startweek+"&endweek="+endweek+"&name=<%=user%>"		
	}
	window.close()
}

</script>

<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">		
</head>
<body bgcolor="#eeeeee" text="#000000" onload="setUp()">
<form name="form1" method="post" action="">
<table border="0" cellpadding="3" cellspacing="0" width="100%" style="border-top:1px solid #ffffff;">
<tr bgcolor="#6699CC">
  <td><span class="standardheader">Current Week Setup</span></td>
</tr>
<tr>
  <td>
  <table border="0" cellpadding="2" cellspacing="0">
  <tr valign="middle"> 
    <td>From:</td>
    <td> 
    <input type="button" name="adjust2" value="-"  onClick="navigate(0, 0, 1)" style="background-color:#dddddd;border:1px outset #eeeeee;color:336699;">
    <input type="text" name="startweek" value="<%=startweek%>">
    <input type="button" name="adjust3" value="+"  onClick="navigate(1, 0, 1)" style="background-color:#dddddd;border:1px outset #eeeeee;color:336699;">
    </td>
    <td><input type="button" name="next" value="Previous Week"  onClick="navigate(0, 2, 7)" style="background-color:#dddddd;border:1px outset #eeeeee;color:336699;"></td>
  </tr>
  <tr valign="middle">
    <td>&nbsp;&nbsp;To:</td>
    <td> 
    <input type="button" name="adjust4" value="-"  onClick="navigate(0, 1, 1)" style="background-color:#dddddd;border:1px outset #eeeeee;color:336699;">
    <input type="text" name="endweek" value="<%=endweek%>">
    <input type="button" name="adjust" value="+" onClick="navigate(1, 1, 1)" style="background-color:#dddddd;border:1px outset #eeeeee;color:336699;">
    </td>
    <td><input type="button" name="next" value="Next Week"  onClick="navigate(1, 2, 7)" style="background-color:#dddddd;border:1px outset #eeeeee;color:336699;"></td>
  </tr>
  <tr>
    <td></td>
    <td colspan="2">
    <input type="button" name="Submit" value="Save" onClick="truncate()" style="border:1px outset #ddffdd;background-color:ccf3cc;font-weight:bold;">
    <input type="button" name="Button" value="Cancel" onclick="document.location='timedetail.asp?name=<%=user%>'" style="border:1px outset #ddffdd;background-color:ccf3cc;">
    </td>
  </tr>
  </table>
  </td>
</tr>
</table>
</form>

</body>
</html>
