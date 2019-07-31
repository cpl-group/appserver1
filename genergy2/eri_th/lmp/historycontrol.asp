<%option explicit%>
<html>
<head>
<title>Calendar</title>
<%
dim billingid, utility
billingid = Request("billingid")
utility = Request("utility")
%>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
var visiblemenu = 'empty'
function showdiv(divdef){

if (visiblemenu != divdef){
document.all[divdef].style.visibility='visible'
if (visiblemenu != "empty"){
	document.all[visiblemenu].style.visibility='hidden'
}
visiblemenu=divdef
}
}
function getvalue(fieldname){
	var mapvalue = "document.forms[0]." + fieldname + ".value"
	eval (mapvalue += "=parent.document.frames.to_cal.calControl.seldate.value")
	
	
}
function loaddate(){
	if (visiblemenu=="view"){
		parent.parent.document.forms[0].startdate.value=document.forms[0].date_box.value
		parent.parent.loadchart();
		parent.document.location.href='options.asp?bldg='+parent.parent.document.forms[0].bldg.value+'&meterid='+parent.parent.document.forms[0].meterid.value+'&billingid='+parent.parent.document.forms[0].billingid.value+'&utility='+parent.parent.document.forms[0].utility.value
	} else if (visiblemenu=="download"){
	var interval;
	for(i=0;i<document.forms[0].interval.length;i++){
    if(document.forms[0].interval[i].checked) interval = document.forms[0].interval[i].value;
  }
	var temp = 'csvdownload.asp?file=show&bldg=' + parent.parent.document.forms[0].bldg.value + '&meterid=' + parent.parent.document.forms[0].meterid.value + '&sd=' +document.forms[0].from_box.value + '&ed=' + document.forms[0].to_box.value + '&billingid=<%=billingid%>&utility=<%=utility%>' + '&interval=' + interval
	parent.document.location = temp
	}


}
</script>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<body bgcolor="#FFFFFF" text="#000000" link="#000000">
<center>
<form name="history" method="post" action="">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
      <tr>
        <td valign="top" height="2">&nbsp;</td>
      </tr>
      <tr> 
        <td valign="top" style="font-family: Arial, Helvetica, sans-serif; font-size: 13px;"> 
          <p><b><input type="radio" name="action" value="view" onclick="showdiv('view')">&nbsp;View Profile</b></p>
          <p><b><input type="radio" name="action" value="download" onclick="showdiv('download')">&nbsp;Download Profile</b></p>
          <div id="view" style="position:absolute; visibility:hidden; width:236px; height:115px; z-index:1; left: 26px"> 
            <p>Date: 
              <input type="text" name="date_box">
              <img src="images/Calendar.gif" width="48" height="36" align="absmiddle" onClick="getvalue('date_box')"></p>
            <p align="center"><a href="#" onclick="loaddate();return false;">View Profile Now</a></p>
          </div>
          <div id="download" style="position:absolute; visibility:hidden; width:269px; height:115px; z-index:1"> 
					<table><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 13px;"><td valign="top">
            <div align="right">From:&nbsp;<input type="text" name="from_box">&nbsp;<img src="images/Calendar.gif" width="48" height="36" align="absmiddle" onclick="getvalue('from_box')"><br>
            To:&nbsp;<input type="text" name="to_box">&nbsp;<img src="images/Calendar.gif" width="48" height="36" align="absmiddle" onClick="getvalue('to_box')"></div>
            <div align=center><a href="#" onclick="loaddate();return false;">Download Data Now</a></div><br>
					</td><td valign="top">
					<input type="radio" name="interval" value="0" checked>&nbsp;15&nbsp;minute<br>
					<input type="radio" name="interval" value="1">&nbsp;Hourly<br>
					<input type="radio" name="interval" value="2">&nbsp;Daily<br></td>
					</tr></table>
          </div>
          <p>&nbsp;</p>
        </td>
      </tr>
    </table>
  </form>
</center>

</body>
</html>


