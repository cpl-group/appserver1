<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim tenantnum, bldg
tenantnum = request("tenantnum")
if instr(tenantnum,".")>0 then
  bldg = split(tenantnum,".")(1)
  tenantnum = split(tenantnum,".")(0)
end if
dim fMsg, loggedin
fMsg = "Welcome."
loggedin = False

loadNewXML tenantnum 'Tweaked to get to next page 

'response.write getxmlusername()

'response.write session("xmlUserObj")
'response.end

' Determine which button the Web visitor clicked and take
' appropriate action.

if not isempty(Request.Form("tenantnum")) then
ProcessLogin
end if

Sub ProcessLogin()
  'Verify that the Web visitor submitted both an email address and a passwd.
  If isempty(tenantnum) Then
    fMsg ="Enter Tenant Number."
    Exit Sub    	
  End If
  
  dim cnn1, rsVis, sql
  Set cnn1 = Server.CreateObject("ADODB.Connection")
  Set rsVis = Server.CreateObject("ADODB.Recordset")
  cnn1.Open getConnect(0,0,"dbCore")
 'sql = "SELECT * FROM "&makeIPUnionDB("tblleases","")&" t  WHERE (tenantnum= '" & tenantnum & "')"
  sql = "SELECT * FROM super_main WHERE bldgnum='"&bldg&"'"
  'SQL= "SELECT * FROM tblleases  WHERE tenantnum= '" & tenantnum & "'"
  dim ip, loggedin, user, path, pid, billlink
  'rsVis.open sql, getConnect(0,0,"dbCore")
 'response.write  getConnect(0,bldg,"Billing")
 'response.end
  rsVis.open sql, getConnect(0,bldg,"Billing")
  if not rsVis.eof then
    ip = rsVis("ip")
    pid = rsVis("pid")
 else
    fMsg = "Tenant number not found."
    exit sub
  end if
  rsVis.close
 
  sql = "SELECT * FROM tblleases WHERE tenantnum='"&tenantnum&"'"
  
  rsVis.open sql, getConnect(0,bldg,"billing")
   
  If rsVis.EOF Then
    fMsg = "Tenant number not found."
  exit sub
  Else
    path = "tenantpage.asp?tenantnum="&tenantnum&"&pid="&pid
    loggedin = True
    user = rsVis("tenantnum")
 
	End If
	rsVis.Close
     
     'response.write "here1"
     'response.end     
  DIM SERVERIP,PORT
  rsVis.open "SELECT location,SERVERIP FROM portfolio p, billtemplates bt WHERE bt.id=p.templateid AND p.id='"&pid&"'", cnn1
  if not rsVis.eof then 
  billlink = rsVis("location")
  SERVERIP = rsVis("SERVERIP")
  end if
  rsVis.close

 If loggedin Then
  ' response.write user
  ' response.end
	loadNewXML user
	IF SERVERIP <> "" THEN 
	PORT ="1433"
	else
	PORT =""
    END IF

	setBuilding bldg, ip ,null,"",PORT
    setKeyValue "bldg", bldg
    setKeyValue "billlink", billlink
    setKeyValue "pid", pid
    response.redirect path
  End If
End Sub
%>
<html>
<head>
<title>gEnergyOne Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="http://appserver1.genergy.com/genergy2/styles.css" type="text/css">
</head>
<script language="javascript">
loaded = 0;
function preloadImg(){
  btnLoginOn = new Image(); btnLoginOn.src = "/images/login/login-1.gif";
  btnLoginOff = new Image(); btnLoginOff.src = "/images/login/login.gif";
  ResetOn = new Image(); ResetOn.src = "/images/login/reset-1.gif";
  ResetOff = new Image(); ResetOff.src = "/images/login/reset.gif";
  loaded = 1;
}

mywidth = screen.availWidth - 8;
myheight = screen.availHeight - 28;
function sizeandcenter(){
  desiredwidth = 580;
  desiredheight = 430;
  window.moveTo(((mywidth/2) - (desiredwidth/2)),((myheight/2) - (desiredheight/2))); 
  window.resizeTo(desiredwidth,desiredheight);
}
function processlogin(){
login.submit();
document.getElementById('progressbar').style.display = 'block';
document.getElementById('slideshow').style.display = 'none';
}
</script>
<body bgcolor="#FFFFFF" link="#000000" vlink="#000000" alink="#000000" leftmargin="0" topmargin="0">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" valign="middle">
        <form name="form1" method="post" action="tenantlogin-g2.asp">
        <table width="600" height="292" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="65"><img src="images/g2login_wei.gif" width="600" height="80"></td>
          </tr>
          <tr> 
            <td width="47" height="224" align="center"><br> <table width="600" height="224" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="14" background="../GENERGYONEV2/images/login/lgin-left.gif">&nbsp;</td>
                  <td width="572" align="center" bgcolor="#e6e6e6" style="border-top:1px solid #CCCCCC;border-bottom:1px solid #CCCCCC;"><table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                            <tr> 
                              <td><b><font face="Arial, Helvetica, sans-serif" size="1" color="#000000">Our 
                                gEnergyOne system offers tenants of buildings 
                                serviced by Genergy's Reading &amp; Billing Services 
                                instant online access to their current and historical 
                                invoices.</font></b></td>
                            </tr>
                            <tr> 
                              <td height="1"> <hr> </td>
                            </tr>
                            <tr> 
                              <td><b><font face="Arial, Helvetica, sans-serif" size="1" color="#000000">Should 
                                you experience problems with the login process, 
                                contact Genergy&#146;s IT department at 212 664 
                                7600 ext. 103, or send us an <a href="mailto:it@genergy.com">email</a>. 
                                </font></b></td>
                            </tr>
                            <tr> 
                              <td height="21"><b><font face="Arial, Helvetica, sans-serif" size="1" color="#000000">Should 
                                you have questions concerning your account information, 
                                please contact our Energy Services Department 
                                at 212 664 7600 ext. 137, or send us an <a href="mailto:george_nemeth@genergy.com">email</a>. 
                                </font></b></td>
                            </tr>
                            <tr> 
                              <td height="21"> <hr> </td>
                            </tr>
                          </table></td>
                        <td width="50%" valign="top"> 
                          <table width="321" border="0" cellspacing="0" cellpadding="0" align="left">
                            <tr> 
                              <td height="16">&nbsp;</td>
                            </tr>
                            <tr> 
                              <td height="16"> <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=fmsg%></font></div></td>
                            </tr>
                            <tr> 
                              <td> <div align="center"><font face="Arial, Helvetica, sans-serif"><font size="2" face="Arial, Helvetica, sans-serif">Please 
                                  Enter Your Tenant Access Code<br>
                                  As Seen On The Bottom Of Your Invoice</font><font size="2"> 
                                  </font></font></div></td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td> <div align="center"><font face="Arial, Helvetica, sans-serif"> 
                                  <input type="text" name="tenantnum">
                                  <input type="submit" name="Submit" value="Log In">
                                  </font></div></td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td> <div align="center"><font size="1" face="Arial, Helvetica, sans-serif">NOTE: 
                                  <a href="http://www.adobe.com/products/acrobat/readstep2.html" target="_blank">ADOBE 
                                  ACROBAT READER</a> IS REQUIRED <br>
                                  TO VIEW INVOICES</font></div></td>
                            </tr>
                          </table>
              </td>
                      </tr>
                    </table> </td>
                  <td width="14" background="../GENERGYONEV2/images/login/lgin-right.gif">&nbsp;</td>
                </tr>
              </table></td>
          </tr>
        </table>
      </form>
</td>
  </tr>
</table>
</body>
</html>
