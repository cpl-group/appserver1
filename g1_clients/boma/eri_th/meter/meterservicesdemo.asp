<!-- #include file="./adovbs.inc" -->
<% 
leaseid= Request("leaseid")
profiletype=Request("profiletype")
portfolio=Request("portfolio")
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function periodfill(leaseid,bldg){
	document.location.href="meterservices.asp?leaseid=" + leaseid + "&bldg=" + bldg;
}
function loadinvoice(inv){
			var temp= "http://www.genergy.com/newdemo/meter/invoice" + inv + ".htm"  
			document.frames.invoice.location.href=temp;
	}
function print_invoice() {

document.invoice.focus();
document.invoice.print();

}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td bgcolor="#0099FF"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#000000">Meter 
          Services </font></div>
      </td>
    </tr>
  </table>
  
  <form method="post" action="" name="lmp">
    <div align="left"> </div>
    <table width="306" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr> 
        <td height="37" width="101"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Tenant</font></div>
        </td>
        <td height="37" width="92"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Period</font></div>
        </td>
      </tr>
      <tr> 
        <td height="56" width="101"> 
          <div align="left"> 
            <div align="center"> 
              <div align="left"></div>
              <div align="left"><font face="Arial, Helvetica, sans-serif" size="3"> 
			  
                <select name="leaseid" onChange="">
                  <option>Select Tenant</option>
                  <option>Tenant 1</option>
                </select>
                </font></div>
            </div>
        </div>
        </td>
        <td height="56" width="92"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">
            <select name="ypid" onChange="loadinvoice(this.value)">
              <option selected>Select Period</option>
              <option value="1">1/31/2001 - 3/2/2001</option>
              <option value="2">3/2/2001 - 4/2/2001</option>
              <option value="3">4/2/2001 - 4/27/2001</option>
            </select>
            </font></div>
        </td>
      </tr>
      <tr> 
        <td height="56" width="101"> 
          <div align="center"><font face="Arial, Helvetica, sans-serif" size="3"> 
            </font> 
            <div align="left"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif">
              <input type="button" name="Button" value="Print Invoice">
              </font></font></div>
          </div>
        </td>
        <td height="56" width="92"> 
          <div align="center"> 
            <div align="left"></div>
            <div align="left"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif" size="3"> 
              </font></font></font><font face="Arial, Helvetica, sans-serif" size="3"> 
              </font></div>
          </div>
        </td>
      </tr>
    </table>
  </form>
  <p align="left"><IFRAME name="invoice" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"> 
    </IFRAME></p>
</div>
</body>
</html>
