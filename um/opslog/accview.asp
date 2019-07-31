<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
function fillup(item){
	document.location="accview.asp?item="+item
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

//visual feedback functions for img buttons
function buttonOver(obj,clr){
  if (arguments.length == 1) { clr = "#336699"; }
  obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
  if (arguments.length == 1) { clr = "#000000"; }
  obj.style.border = "1px solid " + clr;
}

function buttonOut(obj,clr){
  if (arguments.length == 1) { clr = "#eeeeee"; }
  obj.style.border = "1px solid " + clr;
}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" class="innerbody">
<form>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#666699"> 
  <td><span class="standardheader">Invoice Administration</span></td>
  <td align="right"><input type="button" name="Submit" value="Print Current View" onclick='javascript:document.frames.invoice.focus();document.frames.invoice.print()' style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
</tr>
<tr bgcolor="#eeeeee"> 
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;">
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td>
		<a href="corpinvoice.asp" target="invoice">Approve/Reject Submitted Invoices</a> &nbsp;|&nbsp; 
		<a href="accinvoice.asp" target="invoice">Review Approved Invoices</a> 
    </td>
  </tr>
  </table>
  </td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><img src="images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
<tr bgcolor="#eeeeee">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;">
  <table border=0 cellpadding="2" cellspacing="0">
  <tr>
    <td>View past invoices:</td>
    <td>
    <select name="item" onchange="fillup(this.value)">
    <% if request("item")="day" then %>
    <option value="day" selected>Date</option>
    <option value="job">Job No.</option>
    <% else %>
    <option value="day">Date</option>
    <option value="job" selected>Job No.</option>
    </select>
    <% end if %>
    </td>
    
    <% if request("item")="day" then %>
    <td><input type="text" name="date" value="<%=temp%>" size="13%"></td>
    <td>
    <input type="button" name="minus" value="-" onClick="navigate(this.value)" style="width:22px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><input type="button" name="plus" value="+" onClick="navigate(this.value)" style="width:22px;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
    <script>
    setDate()
    </script>
    </td>
    <% else %>
    <td><input type="text" name="date"></td>
    <% end if %>
    <td><input type="button" name="Submit" value="View" onclick="truncate(date.value, item.value, date.value)" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
  </tr>
  </table>
  </td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><img src="/gEnergy2_Intranet/opsmanager/joblog/images/quick_help.gif" alt="?" align="absmiddle" width="19" height="19" border="0"><a href="javascript:toggleHelp();" style="text-decoration:none;"><b>Quick Help</b></a>&nbsp;</td>  
</tr>
</table>
<div id="quickhelp" style="display:none;">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td>
  <br>
  <ul>
  <li><b>To review currently open invoices:</b><br>
  The list below defaults to current invoices (approved invoices that have not yet been closed).
  </ul>
  <ul>
  <li><b>To review past invoices:</b><br>
  Search for past (i.e., closed) invoices by selecting <b>Job No.</b> in the pulldown above and entering four digits in the text field<br>
  Show all closed invoices submitted since a given date by selecting <b>Date</b> in the pulldown and entering a date in the form MM/DD/CCYY<br>
  </ul>
  </td>
</tr>
</table>
</div>
<IFRAME name="invoice" width="100%" height="87%" src="corpinvoice.asp" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 
</form>
</body>
</html>
