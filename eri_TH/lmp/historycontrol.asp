<html>
<head>
<title>Calendar</title>
<%
leaseid = Request.QueryString("luid")
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
		//var temp = 'lmp.asp?bldg=' + parent.parent.document.forms[0].b.value+'&d='+document.forms[0].date_box.value + "&luid=<%=leaseid%>"
		parent.parent.document.forms[0].d.value=document.forms[0].date_box.value
		parent.parent.loadchart();
		parent.document.location.href='options2.asp?b='+parent.parent.document.forms[0].b.value+'&m='+parent.parent.document.forms[0].m.value+'&luid='+parent.parent.document.forms[0].luid.value
		//parent.parent.document.location = temp
	} else if (visiblemenu=="download"){
		var temp = 'downloadlmpdata.asp?bldg=' + parent.parent.document.forms[0].b.value + '&m=' + parent.parent.document.forms[0].m.value + '&sd=' +document.forms[0].from_box.value + '&ed=' + document.forms[0].to_box.value + '&luid=<%=leaseid%>'
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
        <td valign="top"> 
          <p> <font face="Arial, Helvetica, sans-serif"> <font size="2"> <b> 
            <input type="radio" name="action" value="view" onclick="showdiv('view')">
            View Profile</b></font></font></p>
          <p> <b><font face="Arial, Helvetica, sans-serif" size="2"> 
            <input type="radio" name="action" value="download" onclick="showdiv('download')">
            Download Profile</font></b></p>
          <div id="view" style="position:absolute; visibility:hidden; width:236px; height:115px; z-index:1; left: 26px"> 
            <p><font face="Arial, Helvetica, sans-serif" size="2">Date: 
              <input type="text" name="date_box">
              <img src="images/Calendar.gif" width="48" height="36" align="absmiddle" onClick="getvalue('date_box')"></font></p>
            <p align="center"><font face="Arial, Helvetica, sans-serif" size="2"><a href="#" onclick="loaddate();return false;">View 
              Profile Now</a></font></p>
          </div>
          <div id="download" style="position:absolute; visibility:hidden; width:269px; height:115px; z-index:1"> 
            <p align="right"><font face="Arial, Helvetica, sans-serif" size="2">From: 
              <input type="text" name="from_box">
              <img src="images/Calendar.gif" width="48" height="36" align="absmiddle" onclick="getvalue('from_box')"> 
              </font></p>
            <p align="right"><font face="Arial, Helvetica, sans-serif" size="2">To: 
              <input type="text" name="to_box">
              <img src="images/Calendar.gif" width="48" height="36" align="absmiddle" onClick="getvalue('to_box')"></font></p>
            <p align="center"><font face="Arial, Helvetica, sans-serif" size="2"><a href="#" onclick="loaddate();return false;">Download 
              Data Now</a></font></p>
          </div>
          <p>&nbsp;</p>
        </td>
      </tr>
    </table>
  </form>
</center>

</body>
</html>


