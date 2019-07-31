<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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

<script>
function setstartdate(){
	var datetemp = parent.document.forms.form1.d.value
	document.forms[0].startdate.value = datetemp
}
function loaddate(seloption){
if (seloption=='view'){
	var temp = 'lmpindex2.asp?bldg=' + parent.document.forms[0].b.value+'&d='+document.forms[0].fromdate.value
	parent.parent.document.location = temp
}else if(seloption == 'download'){
	var temp = 'downloadlmpdata.asp?bldg=' + parent.parent.document.forms[0].b.value + '&m=' + parent.parent.document.forms[0].m.value + '&sd=' +document.forms[0].fromdate.value + '&ed=' + document.forms[0].todate.value + '&luid=' +parent.parent.document.forms[0].luid.value
	document.location = temp
}
}
</script> 
<%
leaseid = Request.Querystring("luid")
Bldgnum = Request.QueryString("b")

%>
<body bgcolor="#FFFFFF" text="#000000" onload="parent.closeLoadBox('loadFrame2')" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<form name="form1" method="post" action="">
  <div align="center"> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td bgcolor="#000000" width="50%"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"><b>View 
          / Download Historical Profile Data</b></font></td>
        <td bgcolor="#000000" width="50%">
          <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location='<%="options2.asp?b=" & request("b") & "&m=" & request("m") &"&luid="& request("luid")%>'" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return 
            To Options</a></b></font><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"></font></div>
        </td>
      </tr>
    </table>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" height="261">
      <tr> 
        <td valign="top" height="230"> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0" height="104%" align="center">
            <tr> 
              <td height="226" width="50%"> 
                <div align="center"><iframe name="control" width="100%" height="100%" src="<%="historycontrol.asp?b=" & bldgnum &"&luid="&leaseid%>" scrolling="auto" marginwidth="0" marginheight="0" ><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
                  
                  </font></b></font></iframe></div><input type="hidden" name="startdate">
              </td>
              <td height="226" width="50%"> 
                <div align="center"><iframe name="to_cal" width="100%" height="100%" src="to_cal.asp" scrolling="auto" marginwidth="0" marginheight="0" ></iframe></div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr> 
        <td valign="top" bgcolor="#000000" height="23"> </td>
      </tr>
    </table>
  </div>
</form>
<script>
setstartdate()
</script>

</body>
</html>
