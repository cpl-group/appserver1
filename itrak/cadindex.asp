<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2">
try{top.applabel("Building Plans");}catch(exception){}
function fixture(id){
	theURL="fixtureinfo.asp?id=" + id
	var w = 800;
	var h = 600;
	var winl = (screen.width - w) / 2;
    var wint = (screen.height - h) / 2;

    winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',status=yes,scrollbars=yes,resizable=no'
	popupwin=window.open(theURL,'Fixtureinfo',winprops)
	popupwin.focus('Fixtureinfo')
}
function popout(theURL){
	var w = screen.width-20
	var h = screen.height-55
    winprops = 'height='+h+',width='+w+',top=5,left=5,status=yes,scrollbars='+scroll+',resizable=yes'
	popupwin=window.open(theURL,'Fixtureinfo',winprops)
	popupwin.focus('Fixtureinfo')
}

//<!--
loaded = 0;
function preloadImg(){
  expandOn = new Image(); expandOn.src = "images/enlarge_view-1.gif";
  expandOff = new Image(); expandOff.src = "images/enlarge_view.gif";
  loaded = 1;
}

function msover(img){
  if (loaded) {
    document.images[img].src = eval(img + "On.src");
  }
}

function msout(img){
  if (loaded) {
    document.images[img].src = eval(img + "Off.src");
  }
}

//-->
</script>


</head>

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center">
        <tr>
          <td width="100%" height="10" align="right"><div align="left"><a href="javascript:popout('cadindexfull.asp?cad=<%=Request.QueryString("cad")%>');" onMouseOver="msover('expand');" onMouseOut="msout('expand');"><font face="Arial, Helvetica, sans-serif"><strong>expand 
              current view</strong></font></a></div></td>
        </tr>
        <tr> 
          <td> <div align="center">
              <object classid ="clsid:B2BE75F3-9197-11CF-ABF4-08000996E931" codebase = "whip.cab#version=-1,-1,-1,-1" width=100% height=100%>
                <param name="Filename" value="<%=Request.QueryString("cad")%>">
              </object>
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
