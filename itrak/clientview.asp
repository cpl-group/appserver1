<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
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
	
id= Request.Querystring("id")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

sqlstr = "select * from clients where id='"& id &"'"


'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if not rst1.eof then

%>
<html>
<head>
<title>Client View</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
<script src="messages.js" type="text/javascript" language="Javascript1.2"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000">

<form name="form2" method="post" action="clientupd.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #ffffff;">
<tr bgcolor="#0099ff">
	<td><span class="standard" style="color:#ffffff"><b>Client: <%=rst1("corp_name")%></b></span></td>
	<td align="right"><input type="button" value="Account Manager" onclick="document.location.href='manageaccounts.asp?cid=<%=id%>'" class="standard">&nbsp;<input type="button" value="Facilities Manager" onclick="document.location.href='managebldg.asp?cid=<%=id%>'" class="standard">&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=id%>'" class="standard">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=id%>'" class="standard"></td>
</tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="1">
    <tr bgcolor="#eeeeee"> 
      <td width="30%" align="right"><span class="standard">Company Name<input type="hidden" name="bid" value="<%=rst1("id")%>"></span></td> 
      <td width="70%"><input type="text" name="bldgnum" value="<%=rst1("corp_name")%>"></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Address</span></td>
      <td><input type="text" name="address" value="<%=rst1("address")%>"></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">City</span></td>
      <td><input type="text" name="city" value="<%=rst1("city")%>"></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">State</span></td>
      <td><input type="text" name="state" value="<%=rst1("state")%>" size="10"></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Zip Code </span></td>
      <td><input type="text" name="zip" value="<%=rst1("zip")%>" size="10"></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Contact Name</span></td>
      <td><input type="text" name="name1" value="<%=rst1("Contact")%>"></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Contact Phone</span></td>
      <td><input type="text" name="phone1" value="<%=rst1("contactphone")%>"></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><span class="standard">Logo URL</span></td>
      <td><span class="standard"><input type="text" name="logourl" value="<%=rst1("logo")%>">&nbsp;<a onMouseOut="closeHelpBox()" onMouseOver="helpbox('logo_url',event.x,event.y)"><img src="images/question.gif" width="13" height="13" hspace="4" border="0"></a><span class="standard" style="font-size:7pt;">e.g., logos/companyname.gif</span></span></td>	  
    </tr>
	<tr bgcolor="#cccccc">
		<td>&nbsp;</td>
		<td><span class="standard"><input type="submit" name="choice22"  value="Update" class="standard">&nbsp;<input type="reset" name="reset" value="Cancel" onclick="location='manageaccounts.asp?cid=<%=id%>';" class="standard"></span></td>
	</tr>
	</table>
<br><br>
	
<!--table width="100%" border="0" cellpadding="3" cellspacing="1">
    <tr bgcolor="#CCCCCC"> 
      <td width="16%" height="2"><span class="standard">Company Name<input type="hidden" name="bid" value="<%=rst1("id")%>"></span></td> 
      <td width="21%" height="10"><span class="standard">Address</span></td>
      <td width="18%" height="10"><span class="standard">City</span></td>
      <td width="22%" height="10"><span class="standard">State</span></td>
      <td width="23%" height="10"><span class="standard">Zip Code </span></td>
    </tr>
    <tr> 
      <td width=16%> 
        <input type="text" name="bldgnum" value="<%=rst1("corp_name")%>">
         </td>
      <td width="21%" height="19">  
        <input type="text" name="address" value="<%=rst1("address")%>">
        </td>
      <td width="18%" height="19">  
        <input type="text" name="city" value="<%=rst1("city")%>">
        </td>
      <td width="22%" height="19">  
        <input type="text" name="state" value="<%=rst1("state")%>">
        </td>
      <td width="23%" height="19" >  
        <input type="text" name="zip" value="<%=rst1("zip")%>">
        </td>
    </tr>
	</table>
	
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="12%" height="10">Contact 
        Name</td>
      <td width="15%" height="10">Contact 
        Phone #</td>
      <td width="75%" height="10">Logo URL</td>
    </tr>
    <tr> 
      <td width="12%" height="19"> 
        <input type="text" name="name1" value="<%=rst1("Contact")%>">
        </td>
      <td width="12%" height="19" > 
<input type="text" name="phone1" value="<%=rst1("contactphone")%>">
        </td>
      <td width="75%" height="19" > 
<input type="text" name="logourl" value="<%=rst1("logo")%>">
        </td>
    </tr>
    <tr> 
      <td width="12%" height="19" colspan="3">  
        <input type="submit" name="choice22"  value="Update">&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=id%>'">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=id%>'">
        </td>
      <td width="75%" height="19">  
        </td>
      <td width="1%" height="19" >  
        </td>
    </tr>
  </table-->
  <%end if%>
</form>

<!--#INCLUDE FILE="helpbox.htm"-->
</body>
</html>