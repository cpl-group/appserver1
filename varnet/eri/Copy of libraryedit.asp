<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("eri") < 5 or Session("um") < 5 or Session("opslog") < 5 or Session("ts") < 5 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if		
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript" src="../sniffer.js"></script>
<script language="javascript1.2" src="../custom.js"></script>
<script language="javascript1.2" src="../style.js"></script>
<script>
function openpopup(){
//configure "Open Logout Window

parent.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}
//document.main.location="http://www.yahoo.com"

function filter(type1){
    var type=""
	if(type1 =="F"){
		document.main.location="library.asp"
	}else{
		if(type1 == "E"){
			type="Equipment"
		}else if(type1 == "L"){
	    	type="Lighting"
		}else if(type1 == "H"){
			type="HVAC"
		}
		document.main.filter(type)	
	}
}

function sortInOrder(dir, item){
	document.main.sortInOrder(dir, item)
}

</script>
<STYLE>
<!--
A.ssmItems:link		{color:black;text-decoration:none;}
A.ssmItems:hover	{color:black;text-decoration:none;}
A.ssmItems:active	{color:black;text-decoration:none;}
A.ssmItems:visited	{color:black;text-decoration:none;}
//-->
</STYLE>

</head>
<body bgcolor="#FFFFFF" text="#000000">
<p> 
  <script language="javascript1.2" src="../menu.js"></script>
</p>
<p>&nbsp;</p>
<table width="760" border="0" height="105" align="center">
  <tr> 
    <td valign="bottom" height="2" colspan=9> 
      <div align="right"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="115" height="24">
          <param name="MOVIE" value="libraryedit.swf">
          <param name=quality value=high>
          <param name="BGCOLOR" value="">
          <param name="SCALE" value="exactfit">
          <embed src="libraryedit.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" scale="exactfit" width="115" height="24" bgcolor="">
          </embed> 
        </object></div>
    </td>
  </tr>
  <tr> 
    <td width="122" height="41" >Type</td>
    <td width="122" height="41" >Description</td>
    <td width="112" height="41" >Amps</td>
    <td width="42" height="41" >Volt</td>
    <td width="37" height="41" >PH</td>
    <td width="37" height="41" >PF</td>
    <td width="92" height="41" >Watt</td>
    <td width="67" height="41" >MF</td>
    <td width="71" height="41" >AdjFactor 
  </tr>
  <tr> 
    <td height="2" width="122" > 
      <input type="button" name="f" value="F" onClick="filter(this.value)">
      <input type="button" name="l" value="L" onClick=filter(this.value)>
      <input type="button" name="e" value="E" onClick=filter(this.value)>
      <input type="button" name="h" value="H" onClick=filter(this.value)>
    </td>
    <td height="2" width="122" > 
      <input type="button" name="a" value="A-Z" onClick="sortInOrder(this.value, des.value)">
      <input type="button" name="z" value="Z-A" onClick="sortInOrder(this.value, des.value)">
      <input type="hidden" name="des" value="description">
    </td>
    <td height="2" width="112"> 
      <input type="hidden" name="type1" value="<%=type1%>">
      <input type="button" name="a" value="A-Z" onClick="sortInOrder(this.value, am.value)">
      <input type="button" name="z" value="Z-A" onClick="sortInOrder(this.value, am.value)">
      <input type="hidden" name="am" value="amps">
    </td>
    <td height="2" width="42">&nbsp;</td>
    <td height="2" width="37">&nbsp;</td>
    <td height="2" width="37">&nbsp;</td>
    <td height="2" width="92"> 
      <input type="button" name="a" value="A-Z" onClick="sortInOrder(this.value, wa.value)">
      <input type="button" name="z" value="Z-A" onClick="sortInOrder(this.value, wa.value)">
      <input type="hidden" name="wa" value="watt">
    </td>
    <td height="2" width="67">&nbsp;</td>
    <td height="2" width="71">&nbsp;</td>
  </tr>

</table>
 <iframe name="main" width="100%" height="500" src="library.asp" scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
</body>
</html>
