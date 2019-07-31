<!-- #include file="../lmp/./adovbs.inc" -->
<% 
bldg= Request("bldg")
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function yearfill(bldg,pid){
    document.location="cost.asp?bldg="+bldg+"&pid="+pid
}
function createImg(page, y2){
var bldg=page;
if (page==3) {
	if (y2==1){
		bldg=3;
		} else {
			bldg=4;
			}
}
			
    var temp= "buildinga" + bldg + y2 + ".htm"
	document.frames.graph.location=temp;
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


<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#000000">Cost 
        Analysis </font></div>
    </td>
  </tr>
</table>
  
<form method="post" action="" name="list">
<input type="hidden" name="ip" value="10.0.7.16">
  <table width="306" border="0" cellspacing="0" cellpadding="0" align="center" height="55">
    <tr valign="top"> 
      <td height="9" width="103"> 
        <div align="left"> 
          <p><font face="Arial, Helvetica, sans-serif" size="3">Building</font></p>
        </div>
      </td>
      <td height="9" width="47"> 
        <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Year 
          1</font></div>
      </td>
      <td height="9" width="45"> 
        <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Year 
          2</font></div>
      </td>
    </tr>
    <tr> 
      <td height="27" width="103"> 
        <select name="bldg" onChange="yearfill(this.value,pid.value)">
          <option value="b" selected>Building B</option>
        </select>
      </td>

      <td height="27" width="47"> 
        <select name="y1">
          <option value="1">2001</option>
        </select>
      </td>
      <td width="45" height="27"> 
        <select name="y2">
          <option value="1" selected>2001</option>
          <option value="2">2000</option>
        </select>
      </td>
 
   
    </tr>
    <tr> 
      <td width="103" height="70" valign="middle"> 
        <p> 
          <select name="sp_procedure">
            <option value="3" selected>Total Bill Amount</option>
            <option value="2">Unit Cost KW</option>
            <option value="1">Unit Cost KWH</option>
          </select>
        </p>
        </td>
      <td height="70" width="47"> 
        <div align="center"> 
          <input type="button" name="Button" value="View" onClick="createImg(sp_procedure.value, y2.value)">
        </div>
      </td>
      <td height="70" width="45">&nbsp; </td>
    </tr>
  </table>
</form>
<p align="left"><IFRAME name="graph" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p>
</body>
</html>
