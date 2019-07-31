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
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function canceladd(){
document.location.href="null.htm";
parent.document.frames.info.location.href="null.htm";
}
</script>
</head>

<body bgcolor="#999999" text="#000000">
<form name="form1" method="post" action="ti_add_tenant.asp">
  <table width="100%" height="100%" cellpadding="0" cellspacing="0">
    <tr valign="top"> 
      <td width="18%"><font face="Arial, Helvetica, sans-serif"><b>Tenant ID/#</b></font></td>
      <td width="27%"><font face="Arial, Helvetica, sans-serif"><b>Tenant Name</b></font></td>
      <td width="22%"><font face="Arial, Helvetica, sans-serif"></font></td>
      <td width="33%">
        <div align="right"><font face="Arial, Helvetica, sans-serif"><i><b><font color="#66ccFF">ADD 
          TENANT</font></b></i></font></div>
      </td>
    </tr>
    <tr valign="top"> 
      <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="tenant_no" value="" size="10" maxlength="10">
        <input type="hidden" name="bldg_no" value="<%Response.Write Request("bldg")%>">
        </font></td>
      <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="tenantname" value="" size="30">
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
        <input type="text" name="effective_date" value="">
        </font></td>
      <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="lease_exp_date" value="">
        </font></td>
      <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="move_out_date" value="">
        </font></td>
      <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="sqft" value="0">
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
        <input type="text" name="eri_base_date" value="">
        </font></td>
      <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="eri_base_month" value="">
        </font></td>
      <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="ccy" value="">
        </font></td>
      <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="ccm" value="">
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
        <input type="text" name="last_sur_kwh" value="0">
        </font></td>
      <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
         <input type="text" name="last_sur_kw" value="0">
        </font></td>
      <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="bldg_rate" value="">
        </font></td>
      <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="base_hours" value="0">
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
        <input type="text" name="cost_sqft" value="">
        </font></td>
      <td width="27%" height="2"><font face="Arial, Helvetica, sans-serif"></font></td>
      <td width="22%" height="2"> <font face="Arial, Helvetica, sans-serif"> 
        <textarea name="notes" value="" cols="40"></textarea>
        </font></td>
      <td width="33%" valign="bottom" height="2"><font face="Arial, Helvetica, sans-serif"> 
        <input type="submit" name="Submit" value="Save New Tenant">
        <input type="button" name="Submit2" value="Cancel" onclick="canceladd()">
        </font></td>
    </tr>
  </table>
</form>
</body>
</html>
