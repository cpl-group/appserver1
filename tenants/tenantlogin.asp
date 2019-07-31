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
  cnn1.Open application("cnnstr_supermod")
'  sql = "SELECT * FROM "&makeIPUnionDB("tblleases","")&" t  WHERE (tenantnum= '" & tenantnum & "')"
  sql = "SELECT * FROM super_main WHERE bldgnum='"&bldg&"'"

  dim ip, loggedin, user, path, pid, billlink
  rsVis.open sql, application("cnnstr_supermod")
  if not rsVis.eof then
    ip = rsVis("ip")
    pid = rsVis("pid")
  else
    fMsg = "Tenant number not found."
    exit sub
  end if
   rsVis.close
  
  sql = "SELECT * FROM tblleases t WHERE tenantnum='"&tenantnum&"'"
  rsVis.open sql, "driver={SQL Server};server="&ip&";uid=genergy1;pwd=g1appg1;database=genergy2;"
  
  If rsVis.EOF Then
    fMsg = "Tenant number not found."
    exit sub
  Else
    path = "tenantpage.asp?tenantnum="&tenantnum
    loggedin = True
    user = rsVis("tenantnum")
  End If
  rsVis.Close
  
  rsVis.open "SELECT location FROM portfolio p, billtemplates bt WHERE bt.id=p.templateid AND p.id='"&pid&"'", cnn1
  if not rsVis.eof then billlink = rsVis("location")
  rsVis.close
  If loggedin Then
   
    loadNewXML(user)
    setBuilding bldg, ip
    setKeyValue "bldg", bldg
    setKeyValue "billlink", billlink
    setKeyValue "pid", pid
    response.redirect path
  End If
End Sub
%>
<html>
<head>
<title>gEnergyOne - Tenant  Access</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function register(){

	var temp = "register.htm"
	window.open(temp,"", "scrollbars=yes,width=800, height=500, status=no, menubar=no" );


}
</script>
</head>

<body bgcolor="#000000" text="#000000" link="#000000" vlink="#000000" alink="#000000" leftmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr> 
    <td valign="middle" align="center" width="50%"> 
      <table width="843" border="0" cellspacing="0" cellpadding="10" height="529" bgcolor="#FFFFFF">
        <tr align="center"> 
          <td colspan=2 style="border-bottom:2px solid #000000;"><font face="Arial, Helvetica, sans-serif"><strong>gEnergyOne 
            Web-Enabled Invoicing</strong></font></td>
        </tr>
        <tr> 
          <td width="50%"> <table width="100%" border="0" cellspacing="0" cellpadding="5">
              <tr> 
                <td height="2"><font size="2" color="#000000">&nbsp;</font></td>
              </tr>
              <tr> 
                <td><b><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Our 
                  gEnergyOne system offers tenants of buildings serviced by Genergy's 
                  Reading &amp; Billing Services instant online access to their 
                  current and historical invoices.</font></b></td>
              </tr>
              <tr> 
                <td height="1"> <hr> </td>
              </tr>
              <tr> 
                <td><b><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Should 
                  you experience problems with the login process, contact Genergy&#146;s 
                  IT department at 212 664 7600 ex. 103, or send us an <a href="mailto:it@genergy.com">email</a>. 
                  </font></b></td>
              </tr>
              <tr> 
                <td height="21"><b><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Should 
                  you have questions concerning your account information, please 
                  contact our Energy Services Department at 212 664 7600 ex. 128, 
                  or send us an <a href="mailto:george_nemeth@genergy.com">email</a>. 
                  </font></b></td>
              </tr>
              <tr> 
                <td height="21"> <hr> </td>
              </tr>
            </table></td>
          <td width="50%"> <div align="center"> 
              <form name="form1" method="post" action="tenantlogin.asp">
                <table width="328" border="0" cellspacing="0" cellpadding="0" align="left">
                  <tr> 
                    <td> <div align="center"><font face="Arial, Helvetica, sans-serif"><img src="login_header%5B1%5D.jpg" width="321" height="40"></font></div></td>
                  </tr>
                  <tr> 
                    <td height="16"> <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=fmsg%></font></div></td>
                  </tr>
                  <tr> 
                    <td> <div align="center"><font face="Arial, Helvetica, sans-serif"><i><font size="1" face="Arial, Helvetica, sans-serif">Please 
                        Enter Your Tenant Access Code as seen on your invoice</font><font size="2"> 
                        </font></i></font></div></td>
                  </tr>
                  <tr> 
                    <td> <div align="center"><font face="Arial, Helvetica, sans-serif"> 
                        <input type="text" name="tenantnum">
                        <input type="submit" name="Submit" value="Log In">
                        </font></div></td>
                  </tr>
                  <tr> 
                    <td> <div align="center"> </div></td>
                  </tr>
                  <tr> 
                    <td> <div align="center"><font size="1" face="Arial, Helvetica, sans-serif">NOTE: 
                        <a href="http://www.adobe.com/products/acrobat/readstep2.html" target="_blank">ADOBE 
                        ACROBAT READER</a> IS REQUIRED <br>
                        TO VIEW INVOICES</font></div></td>
                  </tr>
                  <tr> 
                    <td height="37"> <div align="center"><font face="Arial, Helvetica, sans-serif"><img src="login_footer%5B1%5D.jpg" width="321" height="40"></font></div></td>
                  </tr>
                </table>
              </form>
            </div></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
