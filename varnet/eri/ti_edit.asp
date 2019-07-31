<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if	
	
tenant_no= Request("tenant_no")



Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"


Set rst1 = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tenant_info WHERE (tenant_no='" & tenant_no & "')"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

' Write a browser-side script to update another frame (named
' detail) within the same frameset that displays this page.

If not rst1.EOF then 
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#999999" text="#000000">
<form name="tenantform" method="post" action="ti_update.asp">
  <table width="100%" height="100%" cellpadding="0" cellspacing="0">
    <tr valign="top"> 
      <td width="18%"><font face="Arial, Helvetica, sans-serif"><b>Tenant</b></font></td>
      <td width="27%"><font face="Arial, Helvetica, sans-serif"></font></td>
      <td width="22%"><font face="Arial, Helvetica, sans-serif"></font></td>
      <td width="33%">
        <div align="right"><font face="Arial, Helvetica, sans-serif"><i><b><font color="#66ccFF">EDIT 
          TENANT</font></b></i></font></div>
      </td>
    </tr>
    <tr valign="top"> 
      <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="tenant_no" value="<%=rst1("tenant_no")%>" size="10" maxlength="10">
        <input type="hidden" name="tenant_id" value="<%=rst1("tenant_no")%>">		
        <input type="hidden" name="bldg_no" value="<%=rst1("bldg_no")%>">
        </font></td>
      <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="tenantname" value="<%=rst1("tenantname")%>">
        </font></td>
      <td width="22%"><font face="Arial, Helvetica, sans-serif"></font></td>
      <td width="33%"><font face="Arial, Helvetica, sans-serif"></font></td>
    </tr>
    <tr valign="top"> 
      <td width="18%"><b><font face="Arial, Helvetica, sans-serif">Commencement 
        Date</font></b></td>
      <td width="27%"><b><font face="Arial, Helvetica, sans-serif">Lease Exp.</font></b></td>
      <td width="22%"><b><font face="Arial, Helvetica, sans-serif">Moveout</font></b></td>
      <td width="33%"><b><font face="Arial, Helvetica, sans-serif">SqFt</font></b></td>
    </tr>
    <tr valign="top"> 
      <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="effective_date" value="<%=rst1("effective_date")%>">
        </font></td>
      <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="lease_exp_date" value="<%=rst1("lease_exp_date")%>">
        </font></td>
      <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="move_out_date" value="<%=rst1("move_out_date")%>">
        </font></td>
      <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
      <% if isnull(rst1("sqft")) then %>
        <input type="text" name="sqft" value="0">
	  <% else %> 
        <input type="text" name="sqft" value="<%=rst1("sqft")%>">
	  <% end if %>
        </font></td>
    </tr>
    <tr valign="top"> 
      <td width="18%" height="2"><b><font face="Arial, Helvetica, sans-serif">Base 
        Date</font></b></td>
      <td width="27%" height="2"><b><font face="Arial, Helvetica, sans-serif">ERI 
        Monthly Base</font></b></td>
      <td width="22%" height="2"><b><font face="Arial, Helvetica, sans-serif">Current 
        Yearly Charge</font></b></td>
      <td width="33%" height="2"><b><font face="Arial, Helvetica, sans-serif">Current 
        Monthly Charge</font></b></td>
    </tr>
    <tr valign="top"> 
      <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="eri_base_date" value="<%=rst1("eri_base_date")%>">
        </font></td>
      <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
	  <% if isNull(rst1("eri_base_month")) then %>
	  <input type="text" name="eri_base_month" value="0">
	  <%else%>
        <input type="text" name="eri_base_month" value="<%=FormatCurrency(rst1("eri_base_month"),2)%>">
        
		<%end if%></font></td>
      <td width="22%">
	   <font face="Arial, Helvetica, sans-serif">
	   <% if isnull(rst1("ccy")) then%>
	   <input type="text" name="ccy" value="0">
	   <%else%>
        <input type="text" name="ccy" value="<%=FormatCurrency(rst1("ccy"),2)%>">
		<%end if%>
        </font></td>
      <td width="33%"> 
	  <font face="Arial, Helvetica, sans-serif"> 
	   <% if isnull(rst1("ccm")) then%>
	   <input type="text" name="ccm" value="0">
	   <%else%>
        <input type="text" name="ccm" value="<%=FormatCurrency(rst1("ccm"),2)%>">
		<%end if%>
        </font></td>
    </tr>
    <tr valign="top"> 
      <td width="18%" height="2"><b><font face="Arial, Helvetica, sans-serif">Surveyed 
        KWH</font></b></td>
      <td width="27%" height="2"><b><font face="Arial, Helvetica, sans-serif">Surveyed 
        KW</font></b></td>
      <td width="22%" height="2"><b><font face="Arial, Helvetica, sans-serif">Lease 
        Rate</font></b></td>
      <td width="33%" height="2"><b><font face="Arial, Helvetica, sans-serif">Base 
        Hrs</font></b></td>
    </tr>
    <tr valign="top"> 
      <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
        <% if IsNull(rst1("last_sur_kwh")) then %>
        <input type="text" name="last_sur_kwh" value="0">
        <% Else %>
        <input type="text" name="last_sur_kwh" value="<%=rst1("last_sur_kwh")%>">
        <% End if %>
        </font></td>
      <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
        <% If IsNull(rst1("last_sur_kw")) then %>
        <input type="text" name="last_sur_kw" value="0">
        <% else %>
        <input type="text" name="last_sur_kw" value="<%=rst1("last_sur_kw")%>">
        <% end if %>
        </font></td>
      <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="bldg_rate" value="<%=rst1("bldg_rate")%>">
        </font></td>
      <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
        <% If isnull(rst1("base_hours")) then %>
        <input type="text" name="base_hours" value="0">
        <% else %>
        <input type="text" name="base_hours" value="<%=rst1("base_hours")%>">
        <% end if %>
        </font></td>
    </tr>
    <tr valign="top"> 
      <td width="18%" height="2"><b><font face="Arial, Helvetica, sans-serif">Current 
        Cost/SqFt</font></b></td>
      <td width="27%" height="2"><b><font face="Arial, Helvetica, sans-serif"></font></b></td>
      <td width="22%" height="2"><b><font face="Arial, Helvetica, sans-serif">Note</font></b></td>
      <td width="33%" height="2"><b><font face="Arial, Helvetica, sans-serif"></font></b></td>
    </tr>
    <tr valign="top"> 
      <td width="18%" height="2"> <font face="Arial, Helvetica, sans-serif"> 
	  <% If isnull(rst1("cost_sqft")) then %>
	  <input type="text" name="cost_sqft" value="0">
	  <%else%>
        <input type="text" name="cost_sqft" value="<%=FormatCurrency(rst1("cost_sqft"),2)%>">
		<%end if%>
        </font></td>
      <td width="27%" height="2"><font face="Arial, Helvetica, sans-serif"></font></td>
      <td width="22%" height="2"> <font face="Arial, Helvetica, sans-serif"> 
        <textarea name="notes" value="<%=rst1("notes")%> cols="40"><%=rst1("notes")%></textarea>
        </font></td>
      <td width="33%" valign="bottom" height="2"><font face="Arial, Helvetica, sans-serif"> 
        <input type="submit" name="Submit" value="Update">
		</font></td>
    </tr>
  </table>
</form>
<% End if

rst1.close
set cnn1=nothing
%>
</body>
</html>
