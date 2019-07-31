<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<% 
	if trim(request("join"))="1" then 
		process
	end if
%>
<head>
<title>MSPNY.ORG - Join Today!!!</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="A0B1C3">
<form>
<table width="567" border="0" align="center">
  <tr> 
    <td width="270" align="center"><font size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong>MSPNY.ORG</strong></font></td>
    <td width="8" align="center"><img src="http://www.genergy.com/mspny/images/line_4.jpg" width="14" height="124"></td>
    <td width="281" align="center"><font size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><em>Join 
      Today!</em></strong></font></td>
  </tr>	
  <tr> 
    <td colspan=3><img src="http://www.genergy.com/mspny/images/line_3.jpg" width="567" height="5"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Name</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="n"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Company</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="c"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Telephone</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="t"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Fax</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="f"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Address</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="a1"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="a2"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">City</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="ci"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">State</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="s"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Zipcode</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="z"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">MSP/MDSP Certification 
      Date</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="mmc"></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">MSP/MDSP ID</font></td>
    <td>&nbsp;</td>
    <td align="right"><input type="text" name="mmid"></td>
  </tr>
  <tr> 
    <td><input name="join" type="hidden" value="1"><input type="submit" name="Submit" value="Join!">
      <input type="reset" name="Submit2" value="Reset"></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>

</form>
</body>
</html>
<%
function process()
	
	Dim message, subject
	
	message =  	n & vbCrLf & vbCrLf 
	message =  	message & c & vbCrLf & vbCrLf 
	message =  	message & f & vbCrLf & vbCrLf 
	message = 	message & a1 & vbCrLf & vbCrLf 
	message = 	message & a2 & vbCrLf & vbCrLf 
	message = 	message & ci & vbCrLf & vbCrLf 
	message = 	message & s & vbCrLf & vbCrLf 
	message = 	message & z & vbCrLf & vbCrLf 
	message = 	message & mmc & vbCrLf & vbCrLf 
	message = 	message & mmid & vbCrLf & vbCrLf 
	
	subject = "New MSPANY Registration"
	sendmail "jose.cotto@genergy.com","MSPANY Registration Site",subject, message

	response.write "Registration information has been sent" 
	response.end
end function 
function sendmail(toadd, fromadd, subject, message)
	dim email, body
	set email = server.createObject("CDONTS.NewMail")
	email.To= toadd
	email.From= fromadd
	email.Subject = Subject
	email.Body = message
	email.Bodyformat=0
	email.Mailformat=1
	email.Send 
end function
%>