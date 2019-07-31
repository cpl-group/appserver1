<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>New Page 1</title>

<meta name="Microsoft Theme" content="none, default">
<link rel="Stylesheet" href="styles.css" type="text/css">
<style type="text/css">
.topline { border-top:1px solid #ffffff; }
.tblunderline td { border-bottom:1px solid #ffffff; }
</style>
</head>

<%
Tenant = Request("qcatnr")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"engineering")


strsql = "SELECT * FROM tenant_history WHERE (tenant_no='" & Tenant & "') order by date_event desc "  
 
rst1.Open strsql, cnn1, adOpenStatic
%>
<body bgcolor="#eeeeee">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td><b>Tenant <%=Tenant%></b></td>
  <%'  If Session("eri") > 2 then%>
<td align="right"><%if not(isBuildingOff(request("bldg"))) then%><input type="button" value="Edit Tenant" onClick="document.location='ti_edit.asp?tenant_no=<%=Tenant%>';">
<%end if%></td>
<% 'end if %>
</tr>
</table>
<div id="infotable" style="overflow:auto;height:180px;width:100%;border:1px solid #dddddd;">
<table border=0 cellpadding="3" cellspacing="1" width="100%" class="tblunderline">
  <tr bgcolor="#dddddd" style="font-weight:normal;"> 
<!--
    [[td]]Tenant Number[[/td]]
-->
    <td>Date Event</td>
    <td>Sur_KWH</td>
    <td>Sur_KW</td>
    <td>Sqft</td>
    <td>% Rate</td>
    <td>% MAC</td>
    <td>Charge</td>
    <td>Note</td>
    <td>&nbsp;</td>
  </tr>
  <form name="form1" method="post" action="infoadd.asp"> 
 <%if not(isbuildingoff(request("bldg"))) then%> 
 <tr>
<!--
    [[td]][[input type="text" name="tenant_no" value="[[%=Tenant%]]" size="15"]][[/td]]
-->
	    <td><input type="hidden" name="tenant_no" value="<%=Tenant%>">
		<input type="text" name="date_event" value="" size="6"></td>
    <td><input type="text" name="sur_kwh" value="0" size="4"></td>
    <td><input type="text" name="sur_kw" value="0" size="4"></td>
    <td><input type="text" name="sqft" value="0" size="6"></td>
    <td><input type="text" name="rate" value="0" size="6"></td>
    <td><input type="text" name="fuel" value="0" size="8"></td>
    <td><input type="text" name="charge" value="0" size="8"></td>
    <td><input type="text" name="note" value="" size="16"></td>
    <td><input type="submit" name="Submit" value="Add" style="padding-left:7px;padding-right:7px;">&nbsp;</td>
  </tr><%end if%>
  </form>
<%
Do While Not rst1.EOF
%>
  <form name="form1" method="post" action="infoupdate.asp"> 
  <tr valign="middle" bgcolor="#eeeeee">
<!--
    [[td]][[input type="text" name="tenant_no" value="[[%=rst1("Tenant_no")%]]" size="15"]][[/td]]
-->
    <td><input type="text" name="date_event" value="<%=rst1("date_event")%>" size="6"></td>
    <td>
    <% if isnull(rst1("sur_kwh")) then %>
      <input type="text" name="sur_kwh" value="0" size="4">
    <%else %>
      <input type="text" name="sur_kwh" value="<%=rst1("sur_kwh")%>" size="4">      
    <%end if %>
    </td>
    <td>
    <% if isnull(rst1("sur_kw")) then %>
      <input type="text" name="sur_kw" value="0" size="4">
    <% else %>
      <input type="text" name="sur_kw" value="<%=rst1("sur_kw")%>" size="4">      
    <% end if %>
     </td>
    <td>
    <% if isnull(rst1("sqft")) then %>
      <input type="text" name="sqft" value="<%=rst1("sqft")%>" size="6">
    <% else %>
      <input type="text" name="sqft" value="<%=rst1("sqft")%>" size="6">      
    <% end if %>
      </td>
    <%if IsNull(rst1("rate")) Then %>
    <td><input type="text" name="rate" value="<% =FormatPercent(0,2) %>" size="6"></td>
    <%else %>
    <td><input type="text" name="rate" value="<%=FormatPercent(rst1("rate"))%>" size="6"></td>
    <% end if %>
    <%if IsNull(rst1("fuel")) Then %>
    <td><input type="text" name="fuel" value="<% =FormatPercent(0,2) %>" size="8"></td>
    <%else %>
    <td><input type="text" name="fuel" value="<%=FormatPercent(rst1("fuel"),2)%>" size="8"></td>
    <% end if %>
    <%if IsNull(rst1("Charge")) Then %>
    <td><input type="text" name="charge" value="<%=FormatCurrency(0,2)%>" size="8"></td>
    <%else %>
    <td><input type="text" name="charge" value=" <%=FormatCurrency(rst1("charge"),2)%>"  size="8"></td>
    <% end if %>
    <td><input type="text" name="note" value="<%=rst1("note")%>" size="16"></td>
    <td> 
      <input type="hidden" name="id" value="<%=rst1("id") %>">
      <%if not(isbuildingoff(request("bldg"))) then%><input type="submit" name="Submit" value="Update"><%end if%>
    </td>
  </tr>
  </form>
  <%
rst1.MoveNext  
Loop

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing
%>
</table>
</div>
</body>

</html>