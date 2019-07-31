<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<%
dim cnn, rst, strsql
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open application("cnnstr_main")

dim cid, company,cName,cType,cTrade,cStatus, cAddress_street,cAddress_floor,cCity,cState,cZip,cPhone ,cFax_phone,date_established,cMaincontact, bAddress_street,bAddress_floor,bCity,bState,bZip,bMainContact, tcolor, cTitle, cContactID, email, contactmsg, customer

cid = Request("cid")
company = Request("company")

if trim(cid)<>"" then
  rst.Open "SELECT * FROM " & company & "_MASTER_ARM_CUSTOMER WHERE customer='"&cid&"'", cnn
  if not rst.EOF then
    cName   = rst("name")
    cType = rst("customer_type")
    cTrade  = rst("trade")
    cStatus = rst("status")
    Select Case cStatus
      case "Active"
        tcolor = "#33FF00"
      case "Inactive"
        tcolor = "#FFFF00"
    end select 
    cAddress_street = rst("address_1")
    cAddress_floor = rst("address_2")
    cCity = rst("city")
    cState = rst("state")
    cZip = rst("zip_code")
    cPhone = rst("telephone")
    cFax_phone = rst("fax")
    date_established = rst("date_established")
    cMaincontact = rst("contact_1")
    bAddress_street = rst("billing_address_1")
    bAddress_floor = rst("billing_address_2")
    bCity = rst("billing_city")
    bState = rst("billing_state")
    bZip = rst("billing_zip_code")  
    bMaincontact= rst("billing_contact")
  end if
  rst.close
end if
%>
<title>Job Search</title>
<script language="JavaScript" type="text/javascript">
//<!--

  function customerdetail(cid) {
    theURL="cis_detail.asp?cid=" + cid + "&company=" + '<%=company%>'
    openwin(theURL,600,475)
  }
  
  function openwin(url,mwidth,mheight){
  window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
  }
  
  function screencompany(company) {
      document.location.href="cis_manage.asp?company="+company	
  }

//-->
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">		
</head>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#ffffff">
<form name="form1">
<table border=0 cellpadding="3" cellspacing="0">
<tr>
  <td width="80">Company</td>
  <td>
  <input type="radio" name="company" value="GY"<%if company="GY" then response.Write(" checked") end if%> onclick="screencompany('GY')">&nbsp;Genergy &nbsp;&nbsp;&nbsp;
  <input type="radio" name="company" value="IL"<%if company="IL" then response.Write(" checked") end if%> onclick="screencompany('IL')">&nbsp;I-Lite<br>
  </td>
</tr>
<tr>
  <td width="80">Customer</td>
  <td>
  <!-- Hold onto customer name as well as ID -->
  <input type="hidden" name="cust_name" value="">
  
  <select name="customer" onChange="customerdetail(this.value)">
  <option value="none" selected>Select Customer</option>
  <%
if company<>"" then
rst.Open "SELECT distinct customer,name FROM " & company & "_MASTER_ARM_CUSTOMER order by name", cnn
    do until rst.eof
if trim(rst("customer"))=customer then
    %>
    <option value="<%=trim(rst("customer"))%>" selected><%=left(trim(rst("name")),30)%></option>
    <%
else
    %>
    <option value="<%=trim(rst("customer"))%>"><%=left(trim(rst("name")),30)%></option>
    <%
end if
    rst.movenext
    loop
rst.close
end if
  %>
  </select>
  </td>
</tr>
</table>
</form>
</body>
</html>