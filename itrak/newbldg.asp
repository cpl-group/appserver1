<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>
<%
dim user
if isempty(Session("name")) then
'Response.Redirect "../index.asp"
else
if Session("opslog") < 2 then 
Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

Response.Redirect "../main.asp"
end if	
end if		
user=Session("name")

dim cid, nid, scroll
cid = request("cid")
nid = request("nid")
scroll = request("scroll")
%>

<title>New Building</title>
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="savebldg.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="1">
<tr bgcolor="#0099ff">
	<td colspan="2"><span class="standardheader">Add New Building</span></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Building Name</span></td> 
	<td><input type="text" name="bldgnum"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">Address</span></td>
	<td><input type="text" name="address"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">City</span></td>
	<td><input type="text" name="city"></td>
</tr>
<tr bgcolor="#eeeeee">
	<td align="right"><span class="standard">State</span></td>
	<td>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr valign="middle">
    <td><input type="text" name="state" size="4"></td>
    <td width="12"><span class="standard">&nbsp;</span></td>
    <td><span class="standard">Zip Code&nbsp;</span></td>
    <td><input type="text" name="zip" size="10"></td>
  </tr>
  </table>
	</td>
</tr>
<tr bgcolor="#eeeeee"> 
	<td align="right"><span class="standard">Square Feet</span></td>
	<td>
  <table border=0 cellpadding="0" cellspacing="0">
  <tr valign="middle">
    <td><input type="text" name="sqft" size="10" maxlength="10"></td>
    <td width="4"><span class="standard">&nbsp;</span></td>
    <td><span class="standard">SQFT</span></td>
  </tr>
  </table>
	</td>
</tr>
<tr bgcolor="#cccccc"> 
	<td><span class="standard">&nbsp;</span></td>
	<td><input type="submit" name="choice22"  value="Save" class="standard"> <input type="reset" value="Cancel" onclick="location='managebldg.asp?cid=<%=cid%>';" class="standard"></td>
</tr>
</table>
<input type="hidden" name="cid" value="<%=cid%>">
<input type="hidden" name="scroll" value="<%=scroll%>">
<input type="hidden" name="nid" value="<%=nid%>">
</form>
</body>
</html>
