<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function canceladd(){
history.go(-1);
parent.document.frames.info.history.go(-1);
}
</script>
<link rel="Stylesheet" href="styles.css" type="text/css">   
</head>

<body bgcolor="#eeeeee" text="#000000">
<form name="form1" method="post" action="ti_add_tenant.asp">

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Add Tenant</span></td>
</tr>
</table>

<table border=0 cellpadding="0" cellspacing="0">
<tr valign="top">
  <td>
  <!-- begin first column -->
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td>Tenant Number</td>
    <td>
    <input type="text" name="tenant_no" value="" size="10" maxlength="10">
    <input type="hidden" name="bldg_no" value="<%Response.Write Request("bldg")%>">
    </td>
  </tr>
  <tr>
    <td>Tenant Name</td>
    <td><input type="text" name="tenantname" value=""></td>
  </tr>
  <tr>
    <td>Commencement Date</td>
    <td><input type="text" name="effective_date" value=""></td>
  </tr>
  <tr>
    <td>Lease Exp.</td>
    <td><input type="text" name="lease_exp_date" value=""></td>
  </tr>
  <tr>
    <td>Moveout</td>
    <td><input type="text" name="move_out_date" value=""></td>
  </tr>
  <tr>
    <td>SqFt</td>
    <td><input type="text" name="sqft" value=""></td>
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
    <td><input type="text" name="eri_base_date" value=""></td>
  </tr>
  <tr> 
    <td>ERI Monthly Base</td>
    <td><input type="text" name="eri_base_month" value=""></td>
  </tr>
  <tr> 
    <td>Current Yearly Charge</td>
    <td><input type="text" name="ccy" value=""></td>
  </tr>
  <tr> 
    <td>Current Monthly Charge</td>
    <td><input type="text" name="ccm" value=""></td>
  </tr>
  <tr> 
    <td>Surveyed KWH</td>
    <td><input type="text" name="last_sur_kwh" value=""></td>
  </tr>
  <tr> 
    <td>Surveyed KW</td>
    <td><input type="text" name="last_sur_kw" value=""></td>
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
    <td><input type="text" name="bldg_rate" value=""></td>
  </tr>
  <tr> 
    <td>Base Hrs</td>
    <td><input type="text" name="base_hours" value=""></td>
  </tr>
  <tr> 
    <td>Current Cost/SqFt</td>
    <td><input type="text" name="cost_sqft" value=""></td>
  </tr>
  <tr> 
    <td>Note</td>
    <td><textarea name="notes" value="" cols="40"></textarea></td>
  </tr>
  <tr> 
    <td>Floor</td>
    <td><input type=text  name="floor" value="" ID="floor"></td>
  </tr>  
  <tr> 
    <td>Online</td>
    <td><input type=checkbox name="Online" value="true" ID="Online"></td>
  </tr>  
  
  <tr>
    <td></td>
    <td valign="bottom">
    <input type="submit" name="Submit" value="Save New Tenant">
    <input type="button" name="Submit2" value="Cancel" onclick="canceladd()">
    </td>
  </tr>
  </table>
  <!-- end third column -->
  </td>
</tr>
</table>  

</form>
</body>
</html>
