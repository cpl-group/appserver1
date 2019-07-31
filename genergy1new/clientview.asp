
<%@Language="VBScript"%>
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

cnn1.Open application("cnnstr_lighting")

sqlstr = "select * from clients where id='"& id &"'"


'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if not rst1.eof then

%>
<html>
<head>
<title>Building Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
.standard { font-family:Arial,Helvetica,sans-serif;font-size:8pt; }
.bottomline { border-bottom:1px solid #eeeeee; }
.floorlink { font-family:Arial,Helvetica,sans-serif;font-size:8pt; color:#0099ff; }
a.floorlink:hover { color:lightgreen; }
.shrunkenheader { font-family:Arial,Helvetica,sans-serif;font-size:7pt;font-weight:bold; }
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">

<form name="form2" method="post" action="clientupd.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #ffffff;">
<tr bgcolor="#0099ff">
	<td><font face="Arial, Helvetica, sans-serif" color="#ffffff"><span class="standard"><b><%=rst1("corp_name")%></b></span></font></td>
	<td align="right"><input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=id%>'" class="standard">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=id%>'" class="standard"></td>
</tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="1">
    <tr bgcolor="#eeeeee"> 
      <td width="30%" align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Company Name<input type="hidden" name="bid" value="<%=rst1("id")%>"></span></font></td> 
      <td width="70%"><font face="Arial, Helvetica, sans-serif"><input type="text" name="bldgnum" value="<%=rst1("corp_name")%>"></font></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Address</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="address" value="<%=rst1("address")%>"></font></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">City</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="city" value="<%=rst1("city")%>"></font></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">State</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="state" value="<%=rst1("state")%>" size="10"></font></td>      
	</tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Zip Code </span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="zip" value="<%=rst1("zip")%>" size="10"></font></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Contact Name</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="name1" value="<%=rst1("Contact")%>"></font></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Contact Phone</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="phone1" value="<%=rst1("contactphone")%>"></font></td>	  
    </tr>
	<tr bgcolor="#eeeeee">
	  <td align="right"><font face="Arial, Helvetica, sans-serif"><span class="standard">Logo URL</span></font></td>
      <td><font face="Arial, Helvetica, sans-serif"><input type="text" name="logourl" value="<%=rst1("logo")%>"></font></td>	  
    </tr>
	<tr bgcolor="#cccccc">
		<td>&nbsp;</td>
		<td><input type="submit" name="choice22"  value="Update" class="standard"></td>
	</tr>
	</table>
<br><br>
	
<!--table width="100%" border="0" cellpadding="3" cellspacing="1">
    <tr bgcolor="#CCCCCC"> 
      <td width="16%" height="2"><font face="Arial, Helvetica, sans-serif"><span class="standard">Company Name<input type="hidden" name="bid" value="<%=rst1("id")%>"></span></font></td> 
      <td width="21%" height="10"><font face="Arial, Helvetica, sans-serif"><span class="standard">Address</span></font></td>
      <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif"><span class="standard">City</span></font></td>
      <td width="22%" height="10"><font face="Arial, Helvetica, sans-serif"><span class="standard">State</span></font></td>
      <td width="23%" height="10"><font face="Arial, Helvetica, sans-serif"><span class="standard">Zip Code </span></font></td>
    </tr>
    <tr> 
      <td width=16%><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="bldgnum" value="<%=rst1("corp_name")%>">
        </font> </td>
      <td width="21%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="address" value="<%=rst1("address")%>">
        </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="city" value="<%=rst1("city")%>">
        </font></td>
      <td width="22%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="state" value="<%=rst1("state")%>">
        </font></td>
      <td width="23%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="zip" value="<%=rst1("zip")%>">
        </font></td>
    </tr>
	</table>
	
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="12%" height="10"><font face="Arial, Helvetica, sans-serif">Contact 
        Name</font></td>
      <td width="15%" height="10"><font face="Arial, Helvetica, sans-serif">Contact 
        Phone #</font></td>
      <td width="75%" height="10"><font face="Arial, Helvetica, sans-serif">Logo URL</font></td>
    </tr>
    <tr> 
      <td width="12%" height="19"> <font face="Arial, Helvetica, sans-serif">
        <input type="text" name="name1" value="<%=rst1("Contact")%>">
        </font></td>
      <td width="12%" height="19" > <font face="Arial, Helvetica, sans-serif">
<input type="text" name="phone1" value="<%=rst1("contactphone")%>">
        </font></td>
      <td width="75%" height="19" > <font face="Arial, Helvetica, sans-serif">
<input type="text" name="logourl" value="<%=rst1("logo")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="12%" height="19" colspan="3"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="submit" name="choice22"  value="Update">&nbsp;<input type="button" value="Menu Configure" onclick="document.location.href='treesetup.asp?cid=<%=id%>'">&nbsp;<input type="button" value="Map Configure" onclick="document.location.href='mapsetup.asp?cid=<%=id%>'">
        </font></td>
      <td width="75%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        </font></td>
      <td width="1%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        </font></td>
    </tr>
  </table-->
  <%end if%>
</form>

</body>
</html>