<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<HTML>
<HEAD>
<TITLE></TITLE>
<META NAME="ROBOTS" CONTENT="NOINDEX,NOFOLLOW">
<script language="JavaScript">
<!--
if ((self.name != 'menuCntrl') & (self.location.protocol != "file:")) {
	self.location = "https://appserver1.genergy.com/eri_th/index.htm";
}
if (parent.theBrowser) {
	if (parent.theBrowser.canOnError) {window.onerror = parent.defOnError;}
}


//onMouseOver="return parent.setStatus('Click to expand all folders in the menu.');" 
//onMouseOut="return parent.clearStatus();"

/* Function that displays status bar messages. */
function MM_displayStatusMsg(msgStr)  { //v3.0
	status=msgStr; document.MM_returnValue = true;
}

/* Functions that swaps images. */
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

if (document.images) {
  expandall_f2 = new Image(56 ,12); expandall_f2.src = "images/menucntrl/expandall_f2.gif";
  expandall_f1 = new Image(56 ,12); expandall_f1.src = "images/menucntrl/expandall.gif";
  collapseall_f2 = new Image(65 ,12); collapseall_f2.src = "images/menucntrl/collapseall_f2.gif";
  collapseall_f1 = new Image(65 ,12); collapseall_f1.src = "images/menucntrl/collapseall.gif";
  help_f2 = new Image(25 ,12); help_f2.src = "images/menucntrl/help_f2.gif";
  help_f1 = new Image(25 ,12); help_f1.src = "images/menucntrl/help.gif";
  home_f2 = new Image(32 ,9); home_f2.src = "images/menucntrl/home_f2.gif";
  home_f1 = new Image(32 ,9); home_f1.src = "images/menucntrl/home.gif";
  frames_f2 = new Image(39 ,9); frames_f2.src = "images/menucntrl/frames_f2.gif";
  frames_f1 = new Image(39 ,9); frames_f1.src = "images/menucntrl/frames.gif";
  noframes_f2 = new Image(59 ,9); noframes_f2.src = "images/menucntrl/noframes_f2.gif";
  noframes_f1 = new Image(59 ,9); noframes_f1.src = "images/menucntrl/noframes.gif";
  floating_f2 = new Image(78 ,12); floating_f2.src = "images/menucntrl/floating_f2.gif";
  floating_f1 = new Image(78 ,12); floating_f1.src = "images/menucntrl/floating.gif";
}
function logoff(){
    
	parent.opener.location.href="https://appserver1.genergy.com/eri_th/login.asp"
	parent.close()
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script>
</HEAD>
<BODY bgcolor="white" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/menucntrl/collapseall_f2.gif','images/menucntrl/expandall_f2.gif')">
<div align="center">
<p><br>
</p><table bgcolor="#ffffff" border="0" cellpadding="0" cellspacing="0" width="143">
<tr><td colspan="5">
              <div align="center">
          <p><font face="Arial, Helvetica, sans-serif">USER: <%=Session("loginemail") %> 
            </font></p>
        </div>
</td>
</tr>
  <tr> 
    <!-- Shim row, height 1. -->
    <td width="7"><img src="images/menucntrl/shim.gif" width="7" height="1" border="0"></td>
    <td width="39"><img src="images/menucntrl/shim.gif" width="39" height="1" border="0"></td>
    <td width="17"><img src="images/menucntrl/shim.gif" width="17" height="1" border="0"></td>
      <td width="14"><img src="images/menucntrl/shim.gif" width="8" height="1" border="0"></td>
    
  </tr>
  
  <tr valign="top">
    <!-- row 2 -->
    
    <td rowspan="2" colspan="2"><a href="javascript:parent.theMenu.openAll();" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_displayStatusMsg('Click to expand all folders in the outline.');MM_swapImage('expandall','','images/menucntrl/expandall_f2.gif',1);return document.MM_returnValue" ><img name="expandall" src="images/menucntrl/expandall.gif" width="56" height="12" border="0" alt="Expand All"></a></td>
      <td rowspan="3" colspan="2"><img src="images/menucntrl/shim.gif" width="16" height="21" border="0"></td>
      <td rowspan="2" colspan="2" width="59"><a href="javascript:parent.theMenu.closeAll();" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_displayStatusMsg('Click to collapse all folders in the outline');MM_swapImage('collapseall','','images/menucntrl/collapseall_f2.gif',1);return document.MM_returnValue" ><img name="collapseall" src="images/menucntrl/collapseall.gif" width="65" height="12" border="0" alt="Collapse All"></a></td>
  </tr>
</table>

<table>
<tr>
      <td bgcolor="#FFFFFF"> <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="100" height="22">
          <param name=movie value="reload.swf">
          <param name=quality value=high>
          <param name="BGCOLOR" value="#FFFFFF">
          <embed src="reload.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" bgcolor="#FFFFFF">
          </embed> 
        </object> </td>
       </tr>
	   <tr>
	   
      <td> <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="100" height="22">
          <param name=movie value="logoff.swf">
          <param name=quality value=high>
          <param name="BGCOLOR" value="#FFFFFF">
          <embed src="logoff.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" bgcolor="#FFFFFF">
          </embed> 
        </object></td>
  </tr>
</table>  
</div>
</BODY>
</HTML>


