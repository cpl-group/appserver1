<html>
<head>
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
function searchjob(typebox, searchitem, comp, var2) {
//alert(typebox)
//alert(searchitem)
	var temp
	if(searchitem != ""){
		if (document.forms[0].comp.checked) {
		    if(var2 !=""){
				temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=1&var="+var2
			}else{
				temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=1"
			}
		} else {
		    if(var2 !=""){
				temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=0&var="+var2
			}else{
				temp="opslogsearch.asp?select=" + typebox + "&findvar="+searchitem+"&comp=0"
			}
		}
		document.frames.oplog.location=temp
    }else{
		alert("At least type something...")
	}
}
function report(spec, job) {
document.frames.oplog.location=spec
}

function timesheetjob(typebox, job){
//alert(typebox)
//alert(document.forms[0].findvar.value)
	var temp
	if(typebox =="[entry id]"){
	    if(job == ""){
		    alert("Please enter job number")
		}else{
            if(isNaN(job)){
				alert("Not a valid number")
            }else{
                temp="timesheetmain.asp?job="+job
//				temp="timesheetsearch.asp?job="+job
				document.frames.oplog.location=temp
			}
		}
	}else{
		temp="null.htm"
		document.frames.oplog.location=temp
	}
	
}

function fillup(item){
	document.location="admininvoices.asp?item="+item
}
function check(v){
    //truncate(date.value, item.value, date.value)">
	if (!isNaN(Date.parse(v))){
	alert(v)
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
	/*str=new Date(currdate)
	str=weekDay(str.getDay())
	currdate=str+" "+currdate
	*/
	document.forms[0].date.value=currdate;
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
		currdate = (now.getMonth() + 1) + "/" + now.getDate() + "/" + now.getFullYear()
		str=new Date(currdate)
		//str=weekDay(str.getDay())
		//currdate=str+" "+currdate
		document.forms[0].date.value=currdate;
	}
}
function Search(item, c){
    if(item=="job" && isNaN(c)){
		alert("Please enter a valid number")
	}else if(item=="day" && isNaN(Date.parse(c))){
	    alert("Please enter a valid day")
	}else{
		document.frames.invoice.location="accinvoice.asp?item="+item+"&date="+c
	}
}
function truncate(date, item, c){
    //var date=document.forms[0].date.value;
	if(c == ""){
		alert("Please enter job number")
	}else{
/*		if (item == "day"){
			d=date.split(" ")
			c=d[1]
		}
*/
		Search(item, c)
	}
}

function toggleHelp(){
  if (document.all.quickhelp.style.display == "none") {
    document.all.quickhelp.style.display = "inline"
  } else {
    document.all.quickhelp.style.display = "none"
  }
}

</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<%
'ReDim Ay(5)
'ReDim By(5)
'Ay(0) = "[entry id]"
'Ay(1) = "customer"
'Ay(2) = "manager"
'Ay(3) = "[current status]"
'Ay(4) = "[description]"
'By(0) = "Job Number"
'By(1) = "Customer ID"
'By(2) = "Manager ID"
'By(3) = "Status"
'By(4) = "Description"
msg = Request.querystring("msg")
typebox = Request("typebox")
			if isempty(msg) then
				msg="Please enter search and click the FIND button to begin"
			end if
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

		
%>
<body bgcolor="#ffffff" text="#000000">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr> 
  <td bgcolor="#666699"><span class="standardheader">Invoice Administration</span></td>
</tr>
<form name="form1">
<tr bgcolor="#eeeeee">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;">
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td>
    <a href="corpinvoice.asp">Approve/Reject Submitted Invoices</a> &nbsp;|&nbsp; 
    <a href="accview.asp">Review Approved Invoices</a>
    </td>
  </tr>
  </table>
  </td>
</tr>
</form>
</table>
<!--
      <input type="button" name="Button" value="Review Submitted Invoices" onClick="Javascript:frames.admin.location='corpinvoice.asp'">
      <input type="button" name="Button2" value="Review Approved Invoices" onClick="Javascript:frames.admin.location='accview.asp'">
-->
<!--<IFRAME name="admin" width="100%" height="92%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" border=0 frameborder=0></IFRAME>-->
</body>
</html>