<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
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
</script>	
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" bgcolor="#3399CC">
  <tr> 
    <td> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">REVIEW 
        APPROVED INVOICES</font></b></div>
    </td>
  </tr>
</table>
<form>
  <p><font face="Arial, Helvetica, sans-serif">View Past Invoices: &nbsp 
    <select name="item" onchange="fillup(this.value)">
      <%
if request("item")="day" then
%>
      <option value="day" selected>Date</option>
      <option value="job">Job No.</option>
      <%
else
%>
      <option value="day">Date</option>
      <option value="job" selected>Job No.</option>
    </select>
    <%
end if
if request("item")="day" then
%>
    <input type="button" name="minus" value=" -" onClick="navigate(this.value)">
    <input type="text" name="date" value="<%=temp%>" size="13%" >
    <input type="button" name="plus" value="+" onClick="navigate(this.value)">
    <script>
setDate()
</script>
    <%
else 
%>
    <input type="text" name="date">
    <%
end if
%>
    </font> 
    <input type="button" name="Submit" value="View" onclick="truncate(date.value, item.value, date.value)">
  </p>
  <p>
    <input type="button" name="Submit2" value="Invoices" onclick='javascript:document.frames.invoice.location="accinvoice.asp"'>
  </p>
  <div align="right"><input type="button" name="Submit" value="Print Current View" onclick='javascript:document.frames.invoice.focus();document.frames.invoice.print()'></div>
</form>
<IFRAME name="invoice" width="100%" height="100%" src="accinvoice.asp" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
</body>
</html>
