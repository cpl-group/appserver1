<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>
<title>Expanded View</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2">
function fixture(id){
	theURL="/itrak/fixtureinfo.asp?id=" + id
	var w = 800;
	var h = 600;
	var winl = (screen.width - w) / 2;
    var wint = (screen.height - h) / 2;

    winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',status=yes,scrollbars=yes,resizable=no'
	popupwin=window.open(theURL,'Fixtureinfodetail',winprops)
	popupwin.focus('Fixtureinfodetail')
}

</script> 
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000">
<table width="100%" height="100%" border="0">
  <tr>
    <td><table width="100%" border="1" cellspacing="0" cellpadding="0" height="100%" align="center">
        <tr> 
          <td> <object classid ="clsid:B2BE75F3-9197-11CF-ABF4-08000996E931" codebase = "whip.cab#version=-1,-1,-1,-1" height=100% width=100%>
              <param name="Filename" value="<%=Request.QueryString("cad")%>">
            </object> </td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
