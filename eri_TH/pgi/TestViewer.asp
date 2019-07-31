<html>
<head>
<title>One-Line Diagram</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
function meters(bldg,meterid) {
	theURL="http://appserver1.genergy.com/cgi-bin/pgimeter.asp?b="+bldg+"&m=" + meterid
	openwin(theURL,600,400)
}
function floors(bldg,floor_) {
	theURL="http://appserver1.genergy.com/cgi-bin/floorsummary.asp?b="+bldg+"&f=" + floor_
	openwin(theURL,600,400)
}
function riser(bldg,riser) {
	theURL="http://appserver1.genergy.com/cgi-bin/riser.asp?b="+bldg+"&r='" + riser+"'"
	openwin(theURL,600,400)
}
function lmp(bldg,meterid) {
	theURL="/genergy2/eri_th/lmp/pgilmp.asp?meterid=" + meterid+"&bldg="+bldg+"&lmp=1&utility=2&interval=0"
	openwin(theURL,570,320)
}
function iri(bldg, filename){
	theURL="../iri/" + bldg + "/" + filename + ".pdf" 
	openwin(theURL,800,700)
}
function pgi(bldg,filename){
	theURL="index.asp?pgi=" + bldg + "/" + filename 
	document.location = theURL
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
function enflexopen(ip, meter, msg){
	
	var myUrl = "http://" + ip + "/cgi-bin/bch.tcl?meter=" + meter + "&msg=" + msg
	window.open(myUrl,"","statusbar=no, menubar=no, HEIGHT=400, WIDTH=450")

}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center">
  <tr> 
    <td>
      <div align="center"><img src="../../images/lock.gif" width="16" height="17"></div>
    </td>
  </tr>
</table>
 <object classid ="clsid:A662DA7E-CCB7-4743-B71A-D817F6D575DF" 
		codebase = "http://www.autodesk.com/global/dwfviewer/installer/DwfViewerSetup.cab" height=600 width=700 VIEWASTEXT ID="Object1">
		<param name="Src" value="<%=Request.QueryString("pgi")%>">
</object>
</body>
</html>