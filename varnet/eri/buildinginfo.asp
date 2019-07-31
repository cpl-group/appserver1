<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
opener.location="../index.asp"
window.close()
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<!-- #include file="./adovbs.inc" -->
<%
if Request("infotype") = "bld" then

%>
<html>
<head>
<title>Building Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="update_bldg_info.asp">
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

sqlstr = "Select * from buildings where bldgnum= " & Request("bldgnum")

rst1.Open sqlstr, cnn1, adOpenStatic, adLockReadOnly

If not rst1.EOF then 
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><font face="Arial, Helvetica, sans-serif"><b><u>Building Address</u></b></font></td>
  </tr>
  <tr> 
    <td><input type="hidden" name="infotype" value="<%=request("infotype")%>">
		<input type="hidden" name="bldgnum" value="<%=request("bldgnum")%>">
        <input type="text" name="bldgname" size="50" value="<%=rst1("bldgname") %>">
    </td>
  </tr>
  <tr> 
    <td>
        <input type="text" name="strt" size="50" value="<%=rst1("Strt") %>">
    </td>
  </tr>
  <tr> 
    <td>
        <input type="text" name="city" size="20" value="<%=rst1("city") %>">
        <input type="text" name="state" size="15" value="<%=rst1("state") %>">
        <input type="text" name="zip" size="10" value="<%=rst1("zip")%>">
    </td>
  </tr>
</table>
<% else %> 
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><font face="Arial, Helvetica, sans-serif"><b><u>Building Address</u></b></font></td>
  </tr>
  <tr> 
   <td><font face="Arial, Helvetica, sans-serif">Data Unavailable Or Not Found For This Building</font></td>
  </tr>
</table>
<% end if %>
  <input type="submit" name="Submit" value="Save / Close">
  <input type="button" name="Submit2" value="Cancel" onclick="Jacascript:window.close()">
</form>

</body>
</html>
<% else %>
<html>
<head>
<title>Building Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="update_bldg_info.asp">
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

sqlstr = "Select * from buildings where bldgnum= " & Request("bldgnum")

rst1.Open sqlstr, cnn1, adOpenStatic, adLockReadOnly

If not rst1.EOF then 
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><font face="Arial, Helvetica, sans-serif"><b><u>Billing Address</u></b></font></td>
  </tr>
  <tr> 
    <td>
        <input type="text" name="btbldgname" size="50" value="<%=rst1("btbldgname") %>">
    </td>
  </tr>
  <tr> 
    <td><input type="hidden" name="infotype" value="<%=request("infotype")%>">
		<input type="hidden" name="bldgnum" value="<%=request("bldgnum")%>">
        <input type="text" name="btstrt" size="50" value="<%=rst1("btStrt") %>">
    </td>
  </tr>
  <tr> 
    <td>
        <input type="text" name="btcity" size="20" value="<%=rst1("btcity") %>">
        <input type="text" name="btstate" size="15" value="<%=rst1("btstate") %>">
        <input type="text" name="btzip" size="10" value="<%=rst1("btzip")%>">
    </td>
  </tr>
</table>
<% else %> 
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><font face="Arial, Helvetica, sans-serif"><b><u>Building Address</u></b></font></td>
  </tr>
  <tr> 
   <td><font face="Arial, Helvetica, sans-serif">Data Unavailable Or Not Found For This Building</font></td>
  </tr>
</table>
<%end if %>
  <input type="submit" name="Submit" value="Save / Close">
  <input type="button" name="Submit2" value="Cancel" onclick="Javascript:window.close()">
</form>

</body>
</html>
<% end if %>