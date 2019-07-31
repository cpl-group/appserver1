<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

  <%
tenant_no= Request("tenant_no")



Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")


Set rst1 = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tenant_info WHERE (tenant_no='" & tenant_no & "')"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

' Write a browser-side script to update another frame (named
' detail) within the same frameset that displays this page.

If not rst1.EOF then 
%>

<html>
<head>
<title>Edit Tenant</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="styles.css" type="text/css">   
</head>

<body bgcolor="#eeeeee" text="#000000">

<form name="tenantform" method="post" action="ti_update.asp">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Edit Tenant</span></td>
</tr>
</table>

<table border=0 cellpadding="0" cellspacing="0">
<tr valign="top">
  <td>
  <!-- begin first column -->
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td>Tenant Number<input type="hidden" name="tenant_no_real" value="<%=rst1("tenant_no")%>"></td>
    <td><input type="text" name="tenant_no" value="<%=rst1("tenant_no")%>" size="10" maxlength="10"></td>
  </tr>
  <tr>
    <td>Tenant Name</td>
    <td><input type="text" name="tenantname" value="<%=rst1("tenantname")%>"></td>
  </tr>
  <tr>
    <td>Commencement Date</td>
    <td><input type="text" name="effective_date" value="<%=rst1("effective_date")%>"></td>
  </tr>
  <tr>
    <td>Lease Exp.</td>
    <td><input type="text" name="lease_exp_date" value="<%=rst1("lease_exp_date")%>"></td>
  </tr>
  <tr>
    <td>Moveout</td>
    <td><input type="text" name="move_out_date" value="<%=rst1("move_out_date")%>"></td>
  </tr>
  <tr>
    <td>SqFt</td>
    <td><input type="text" name="sqft" value="<% if isnull(rst1("sqft")) then %>0<% else %><%=rst1("sqft")%><% end if %>"></td>
  </tr>
  </table>
  <!-- end first column -->
  </td>
  <td width="30">&nbsp;</td>
  <td>
  <!-- begin second column -->
  <table border=0 cellpadding="3" cellspacing="0">
  <tr> 
    <td>Base Date</td>
    <td><input type="text" name="eri_base_date" value="<%=rst1("eri_base_date")%>"></td>
  </tr>
  <tr> 
    <td>ERI Monthly Base</td>
    <td><input type="text" name="eri_base_month" value="<% if isNull(rst1("eri_base_month")) then %>0<% else %><%=FormatCurrency(rst1("eri_base_month"),2)%><%end if%>"></td>
  </tr>
  <tr> 
    <td>Current Yearly Charge</td>
    <td><input type="text" name="ccy" value="<% if isnull(rst1("ccy")) then%>0<% else %><%=FormatCurrency(rst1("ccy"),2)%><%end if%>"></td>
  </tr>
  <tr> 
    <td>Current Monthly Charge</td>
    <td><input type="text" name="ccm" value="<% if isnull(rst1("ccm")) then%>0<% else %><%=FormatCurrency(rst1("ccm"),2)%><%end if%>"></td>
  </tr>
  <tr> 
    <td>Surveyed KWH</td>
    <td><input type="text" name="last_sur_kwh" value="<% if IsNull(rst1("last_sur_kwh")) then %>0<% else %><%=rst1("last_sur_kwh")%><% end if %>"></td>
  </tr>
  <tr> 
    <td>Surveyed KW</td>
    <td><input type="text" name="last_sur_kw" value="<% If IsNull(rst1("last_sur_kw")) then %>0<% else %><%=rst1("last_sur_kw")%><% end if %>"></td>
  </tr>
  </table>  
  <!-- end second column -->
  </td>
  <td width="30">&nbsp;</td>
  <td>
  <!-- begin third column -->
  <table border=0 cellpadding="3" cellspacing="0">
  <tr> 
    <td>Lease Rate</td>
    <td><input type="text" name="bldg_rate" value="<%=rst1("bldg_rate")%>"></td>
  </tr>
  <tr> 
    <td>Base Hrs</td>
    <td><input type="text" name="base_hours" value="<% If isnull(rst1("base_hours")) then %>0<% else %><%=rst1("base_hours")%><% end if %>"></td>
  </tr>
  <tr> 
    <td>Current Cost/SqFt</td>
    <td><input type="text" name="cost_sqft" value="<% If isnull(rst1("cost_sqft")) then %>0<% else %><%=FormatCurrency(rst1("cost_sqft"),2)%><%end if%>"></td>
  </tr>
  <tr> 
    <td>Note</td>
    <td><textarea name="notes" value="<%=rst1("notes")%> cols="40"><%=rst1("notes")%></textarea></td>
  </tr>
  <tr> 
    <td>Floor</td>
    <td><input type=text  name="floor" value="<%=rst1("floor")%>"  ID="floor"></td>
    <td>Online</td>
     <% If rst1("Online") = True then %>
    <td><input type=checkbox name="Online"  checked ></td>
    <% Else %>
    <td><input type=checkbox name="Online"></td>
    <% End If%>
  </tr>  
  <tr>
    <td></td>
    <td><input type="submit" name="Submit" value="Update"><input type="button" name="Cancel" value="Cancel" onclick="history.go(-1);"></td>
  </tr>
  </table>
  <!-- end third column -->
  </td>
</tr>
</table>  
</form>
<% End if

rst1.close
set cnn1=nothing
%>
</body>
</html>
