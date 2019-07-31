<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>

<%
		if isempty(Session("name")) then
'			Response.Redirect "../index.asp"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
	
%>

<title>New Client</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="saveclient.asp">
 <table width="100%" border="0" cellpadding="3" cellspacing="1">
    <tr bgcolor="#eeeeee"> 
      <td width="30%" align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Company Name</span></font></td> 
      <td width="70%"><font face="Arial, Helvetica, sans-serif"><input type="text" name="bldgnum" value=""></font></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Address</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="address" value=""></font></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">City</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="city" value=""></font></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">State</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="state" value="" size="10"></font></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Zip Code </span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="zip" value="" size="10"></font></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Contact Name</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="name1" value=""></font></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Contact Phone</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="phone1" value=""></font></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Logo URL</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="logoURL" value=""></font></td>	  
    </tr>
	<tr bgcolor="#cccccc">
		<td>&nbsp;</td>
		<td><input type="submit" name="choice22"  value="Save" class="standard"></td>
	</tr>
	</table>
	
 <!--table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="16%" height="2"><font face="Arial, Helvetica, sans-serif">Company 
        Name </font></td> 
      <td width="21%" height="10"><font face="Arial, Helvetica, sans-serif">Address</font></td>
      <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif">City</font></td>
      <td width="22%" height="10"><font face="Arial, Helvetica, sans-serif">State</font></td>
      <td width="23%" height="10"><font face="Arial, Helvetica, sans-serif">Zip 
        Code </font></td>
    </tr>
    <tr> 
      <td width=16%><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="bldgnum" >
        </font> </td>
      <td width="21%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="address">
        </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="city">
        </font></td>
      <td width="22%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="state" >
        </font></td>
      <td width="23%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="zip">
        </font></td>
    </tr>
	</table>
	
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="16%" height="10"><font face="Arial, Helvetica, sans-serif">Contacts 
        Name </font></td>
      <td width="8%" height="10"><font face="Arial, Helvetica, sans-serif">Contact 
        Phone #</font></td>
      <td width="8%" height="10"><font face="Arial, Helvetica, sans-serif">Logo URL</font></td>
      <td width="75%" height="10">&nbsp;</td>
    </tr>
    <tr> 
      <td width="16%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="name1" >
        </font></td>
      <td width="8%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="phone1" >
        </font></td>
      <td width="8%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="logoURL" >
        </font></td>
      <td width="75%" height="19" >&nbsp; </td>
    </tr>
    
    <tr> 
      <td width="16%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="submit" name="choice22"  value="SAVE">
        </font></td>
      <td width="8%" height="19"> <font face="Arial, Helvetica, sans-serif"> </font></td>
      <td width="75%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        </font></td>
      <td width="1%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        </font></td>
    </tr>
  </table-->
</form>
</body>
</html>
