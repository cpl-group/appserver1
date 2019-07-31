<html>
<head>
<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
	
%>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="capaddbldg.asp">
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="2%" height="2">&nbsp;</td> 
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Building 
        No. </font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Address</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Square 
        Feet </font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Revision</font></td>
    </tr>
    <tr> 
      <td width=6%> 
        <input type="submit" name="choice2"  value="SAVE">
      </td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="bldgnum" >
        </font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="address">
        </font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="sqft" >
        </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="rev">
        </font></td>
    </tr>
  </table>
</form>
</body>
</html>
