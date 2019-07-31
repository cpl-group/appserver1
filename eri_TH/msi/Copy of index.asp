<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript1.2">
function meters(bldg,meterid) {
	theURL="https://appserver1.genergy.com/cgi-bin/pgimeter.asp?b="+bldg+"&m=" + meterid
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
	theURL="https://appserver1.genergy.com/eri_th/lmp/lmpload2.asp?m=" + meterid+"&b="+bldg+"&lmp=1"
	openwin(theURL,800,700)
}
function iri(bldg, filename){
	theURL="https://appserver1.genergy.com/eri_th/iri/" + bldg + "/" + filename + ".pdf" 
	openwin(theURL,800,700)
}
function pgi(bldg,filename){
	theURL="index.asp?pgi=" + bldg + "/" + filename 
	document.location = theURL
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000">
<table width="600" border="0" cellspacing="0" cellpadding="0" height="480" align="center">
  <tr> 
    <td>
      <div align="center"><img src="../../images/lock.gif" width="16" height="17"></div>
    </td>
  </tr>
  <tr> 
    <td> <object classid ="clsid:B2BE75F3-9197-11CF-ABF4-08000996E931" codebase = "whip.cab#version=-1,-1,-1,-1" height=600 width=700>
        <param name="Filename" value="<%=Request.QueryString("pgi")%>">
      </object> </td>
  </tr>
</table>
</body>
</html>
