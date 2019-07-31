<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
m = Request.QueryString("m")
b = Request.QueryString("b")
luid= Request.Querystring("luid")

%>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000">
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="119">
    <tr> 
      <td bgcolor="#000000"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="159" height="24">
          <param name=movie value="options.swf">
          <param name=quality value=high>
          <param name="BGCOLOR" value="#000000">
          <param name="SCALE" value="exactfit">
          <embed src="options.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" scale="exactfit" width="159" height="24" bgcolor="#000000">
          </embed> 
        </object></td>
    </tr>
    <tr>
      <td height="23" valign="top"></td>
    </tr>
    <tr> 
      <td height="271" valign="top"> 
        <ul>
          <li><font face="Arial, Helvetica, sans-serif"><a href="<%="pk_hist_sheet.asp?m="&m&"&b="&b&"&luid="&luid%>" style="text-decoration:none;" onMouseOver="this.style.color = 'gray'" onMouseOut="this.style.color = 'black'"><font size="2"><b>Peak 
            Demand History</b></font></a></font></li>
        </ul>
        <blockquote> 
          <p><font face="Arial, Helvetica, sans-serif"><b><font size="2">Coming 
            Soon...</font></b></font></p>
          <ul>
            <li><font face="Arial, Helvetica, sans-serif" size="2"><b>Tenant Contributions 
              to Peak Demand</b></font></li>
            <li><b><font face="Arial, Helvetica, sans-serif" size="2">Usage Projections 
              </font></b></li>
          </ul>
        </blockquote>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
