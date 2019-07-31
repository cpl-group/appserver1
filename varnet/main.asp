<%@Language="VBScript"%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function openfolder(){
	var temp = 'file://10.0.7.21/intranet/datafiles/'

	window.open(temp,"FolderWindow", "scrollbars=yes, width=500, height=300, resizeable, status" );
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr>
    <td>
      <div id="Layer1" style="position:absolute; width:262px; height:488px; z-index:1; left: 30px; top: 58px; background-color: #3399CC; layer-background-color: #3399CC; border: 1px none #000000; overflow: scroll"> 
        <div align="center"> 
          <p><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Company 
            News</font></b></font></p>
          <p align="left"><b></b></p>
        </div>
      </div>
      <div id="Layer1" style="position:absolute; width:262px; height:488px; z-index:1; left: 307px; top: 58px; background-color: #3399CC; layer-background-color: #3399CC; border: 1px none #000000; overflow: scroll"> 
        <div align="center"> 
          <p><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Calendar</font></b></font></p>
          <p align="left"><b></b></p>
        </div>
      </div>
      <div id="Layer1" style="position:absolute; width:262px; height:488px; z-index:1; left: 588px; top: 57px; background-color: #3399CC; layer-background-color: #3399CC; border: 1px none #000000; overflow: scroll"> 
        <div align="center"> 
          <p><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Employee 
            Links </font></b></font></p>
          <ul>
            <li> 
              <div align="left"><font face="Arial, Helvetica, sans-serif"><b><a href="P1/index.html">Operations 
                Manual</a></b></font></div>
            </li>
            <li>
              <div align="left"><b><font face="Arial, Helvetica, sans-serif"><a href="#" onclick="openfolder()">Misc 
                Files (Exhibits, Pricesheets, etc.)</a></font></b></div>
            </li>
          </ul>
          <p align="left"><b></b></p>
        </div>
      </div>
    </td>
  </tr>
</table>
<p>&nbsp; </p>
</body>
</html>
