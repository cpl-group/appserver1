<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<!-- #include file="./adovbs.inc" -->
<%
id=Request("id")
tenant_no=Request("tenant_no")
Response.Write(id)
%>
<html>
<head>
<title>Building Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="surveyadd.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><font face="Arial, Helvetica, sans-serif"><b><u>Survey Date</u></b></font></td>
	<td><font face="Arial, Helvetica, sans-serif"><b><u>Location</u></b></font></td>
  </tr>
  <tr> 
    <td><input type="hidden" name="Survey Date" value="">
		<input type="hidden" name="Location" value="">
        
    </td>
  </tr>
  <tr>
    <td><font face="Arial, Helvetica, sans-serif"><b><u>Floor</u></b></font></td>
	<td><font face="Arial, Helvetica, sans-serif"><b><u>Order No</u></b></font></td>
  </tr>
  
  <tr> 
    <td>
		<input type="text" name="Floor" size="50" value="">
        <input type="text" name="Order No" size="50" value="">
    </td>
  </tr>
</table>
<input type="submit" name="Submit" value="OK / Close">
</form>

</body>
</html>
