<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2">
function meters(bldg,meterid) {
	theURL="https://appserver1.genergy.com/cgi-bin/pgimeterdemo.asp?bldg="+bldg+"&meterid=" + meterid
	openwin(theURL,600,400)
}
function floors(bldg,floor_) {
	theURL="https://appserver1.genergy.com/cgi-bin/floorsummary.asp?b="+bldg+"&f=" + floor_
	openwin(theURL,600,400)
}
function riser(bldg,riser) {
	theURL="https://appserver1.genergy.com/cgi-bin/riser.asp?b="+bldg+"&r='" + riser+"'"
	openwin(theURL,600,400)
}
function lmp(bldg,meterid) {
	theURL="https://appserver1.genergy.com/genergy2/eri_th/lmp/lmp.asp?meterid=" + meterid+"&bldg="+bldg+"&lmp=1"
	openwin(theURL,800,700)
}

function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000">
<table width="100%" border="1" cellspacing="0" cellpadding="0" height="100%" align="center">
  <tr> 
    <td> 
		<object classid ="clsid:B2BE75F3-9197-11CF-ABF4-08000996E931" codebase = "whip.cab#version=-1,-1,-1,-1" height=100% width=100%>
        <param name="Filename" value="<%=Request.QueryString("pgi")%>">
      	</object> </td>
  </tr>
</table>
</body>
</html>
