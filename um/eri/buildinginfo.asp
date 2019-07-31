<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<!-- #include file="./adovbs.inc" -->
<html>
<head>
<title>Building Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="styles.css" type="text/css">   
</head>

<body bgcolor="#eeeeee" text="#000000">
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"engineering")

sqlstr = "Select * from buildings where bldgnum= '" & Request("bldgnum") & "'"
'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, adOpenStatic, adLockReadOnly

If not rst1.EOF then 
%>
<form name="form1" method="post" action="update_bldg_info.asp">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td colspan="2"><b>Building Address</b></td>
  <td colspan="2"><b>Billing Information</b></td>
</tr>
<tr>
  <td>Building Name</td>
  <td>
  <input type="hidden" name="infotype" value="<%=request("infotype")%>">
  <input type="hidden" name="bldgnum" value="<%=request("bldgnum")%>">
  <input type="text" name="bldgname" size="30" value="<%=rst1("bldgname") %>">
  </td>
  <td>Building Name</td>
  <td><input type="text" name="btbldgname" size="30" value="<%=rst1("btbldgname") %>"></td>
</tr>
<tr>
  <td>Street</td>
  <td><input type="text" name="strt" size="30" value="<%=rst1("Strt") %>"></td>
  <td>Street</td>
  <td><input type="text" name="btstrt" size="30" value="<%=rst1("btStrt") %>"></td>
</tr>
<tr>
  <td nowrap>City, State, Zip</td>
  <td nowrap>
  <input type="text" name="city" size="14" value="<%=rst1("city") %>">
  <input type="text" name="state" size="4" value="<%=rst1("state") %>">
  <input type="text" name="zip" size="10" value="<%=rst1("zip")%>">
  </td>
  <td nowrap>City, State, Zip</td>
  <td nowrap>
  <input type="text" name="btcity" size="14" value="<%=rst1("btcity") %>">
  <input type="text" name="btstate" size="4" value="<%=rst1("btstate") %>">
  <input type="text" name="btzip" size="10" value="<%=rst1("btzip")%>">
  </td>
</tr>
<tr bgcolor="#dddddd">
  <td>&nbsp;</td>
  <td colspan="3">
  <%if not(isbuildingoff(request("bldgnum"))) then%><input type="submit" name="Submit" value="Save"><%end if%>
  <input type="button" name="Submit2" value="Cancel" onclick="history.go(-1);">
  </td>
</tr>
</table>

<% else %> 
<table width="100%" border="0" cellspacing="3" cellpadding="0">
<tr> 
  <td>Data unavailable or not found for this building</td>
</tr>
</table>
<% end if %>
</form>

</body>
</html>
