<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<%
dim cnn, rst, strsql, ticket
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")
dim cid, company,cName,cType,cTrade,cStatus, cAddress_street,cAddress_floor,cCity,cState,cZip,cPhone,cCell,cFax_phone,date_established,cMaincontact, bAddress_street,bAddress_floor,bCity,bState,bZip,bMainContact, tcolor, cTitle, cContactID, email, contactmsg, managername

cid = secureRequest("cid")
company = secureRequest("company")

if trim(cid)<>"" then
  rst.Open "SELECT mac.*, m.lastname+', '+m.firstname as managername FROM " & company & "_MASTER_ARM_CUSTOMER mac LEFT JOIN Managers m ON mac.acct_manager=m.mid WHERE customer='"&cid&"'", cnn
  
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
	managername = rst("managername")
  end if
  rst.close
end if
set ticket = New tickets
ticket.Label="Customer"
ticket.Note="Customer Master Ticket "
ticket.requester = "JOBLOGADMIN"
ticket.department = "OPERATIONS"
ticket.userid = "JOBLOGADMIN"
ticket.findtickets "customerid", cid
%>
<html>
<head>
<title>Customer Details </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function editcustomer(cid, company)
{
  pageURL = "cis_update.asp?cid=" + cid + "&company=" + company + "&mode=edit"
  document.location = pageURL
  //window.resizeTo(600,300)
}
function opencontact(contactid,company, mode) 
{
  theURL="updatecontact.asp?contactid=" + contactid + "&company=" + company + "&mode=" + mode
  openwin(theURL,500,525)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">   
</head>
<body bgcolor="#dddddd">
<form name="form2" method="post" action="">

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr valign="top" bgcolor="#6699cc">
  <td>
  <table border=0 cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td><span class="standardheader">Customer Details</span></td>
    <td align="right"><input type="button" value="Edit Customer" onclick="editcustomer('<%=cid%>', '<%=company%>')">&nbsp;<%ticket.MakeButton%></td>
  </tr>
  </table>
  </td>
</tr>
<tr valign="top" bgcolor="#eeeeee">
  <td style="border-bottom:1px solid #cccccc;">
  <table border=0 cellpadding="3" cellspacing="0">
  <tr>
    <td>Customer Name</td>
    <td><%=cName%> (<%=cid%>) <%=cStatus%></td>
  </tr>
  <tr valign="middle" bgcolor="#eeeeee">
    <td>Customer Type</td>
    <td><%=cType%></td>
  </tr>
  <tr valign="middle" bgcolor="#eeeeee">
            <td nowrap>Account Manager</td>
    <td><%=managername%></td>
  </tr>
  <tr valign="middle" bgcolor="#eeeeee">
            <td nowrap></td>
    <td><%ticket.Display 0,true, true, false%></td>
  </tr>
  </table>
  </td>
</tr>
<tr valign="top" bgcolor="#eeeeee">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td width="35%"><b>Main Address</b></td>
    <td width="30" rowspan="4">&nbsp;</td>
    <td><b>Billing Address</b></td>
  </tr>
  <tr>
    <td><%=cAddress_street%></td>
    <td><%=bAddress_street%></td>
  </tr>
  <tr>
    <td><%=cAddress_floor%></td>
    <td><%=bAddress_floor%></td>
  </tr>
  <tr>
    <td><%=cCity%>, <%=cState%> &nbsp;&nbsp;<%=cZip%></td>
    <td><%=bCity%>, <%=bState%> &nbsp;&nbsp;<%=bZip%></td>
  </tr>
  <tr><td colspan="2" height="8"></td></tr>
  </table>
  </td>
</tr>
<tr valign="top" bgcolor="#eeeeee">
  <td style="border-top:1px solid #ffffff;">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td><b>Customer Contacts</b></td>
    <td align="right"><input type="button" value="New Contact" onclick="opencontact('<%=cid%>', '<%=company%>','new')"></td>
  </tr>
  </table>
  </td>
</tr>
<tr valign="top" bgcolor="#eeeeee">
  <td>
  <div style="overflow:auto;height:235;width:100%;">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
        <%
  rst.Open "SELECT * FROM " & company & "_STANDARD_ARS_CONTACT WHERE customer ='"&cid&"'", cnn      
  While not rst.EOF 
    cName     = rst("contact_name")
    cType     = rst("contact_type")
    cTitle    = rst("title")
    cContactID  = rst("contact")
    email     = rst("email_address")
    cAddress_street = rst("address_1")
    cAddress_floor = rst("address_2")
    cCity = rst("city")
    cState = rst("state")
    cZip = rst("zip_code")
    cPhone = rst("telephone")
    cCell = rst("cellular") '1/30/2008 N.ambo added option for cellular
    cFax_phone = rst("fax")
    if trim(ccontactid) = trim(cMaincontact) and trim(ccontactid) = trim(bMaincontact) then 
      contactmsg = "(Primary and Billing Contact)"
    else
      if trim(ccontactid) = trim(cMaincontact) then 
        contactmsg = "(Primary Contact)"
      else
        if trim(ccontactid) = trim(bMaincontact) then 
        contactmsg = "(Billing Contact)"
        end if  
      end if
    end if 

  %>
   <tr bgcolor="#eeeeee"> 
   <td style="color:#999999;" width="18%">Name:</td>
    <td width="82%"><%=cName%></td>
  </tr>
  <%if contactmsg <> "" then %>
  <tr>
    <td style="color:#999999;">Description:</td>
    <td><%=contactmsg%></td>
  </tr>
  <%end if%>
  <% if cTitle <> "" then %>
  <tr>
    <td style="color:#999999;">Title:</td>
    <td><%=cTitle%></td>
  </tr>
  <% end if %>
  <% if cType <> "" then %> 
  <tr>
    <td style="color:#999999;">Contact Type:</td>
    <td><%=cType%></td>
  </tr>
  <% end if %>
  <tr> 
    <td style="color:#999999;">Telephone:</td>
    <td><%=cphone%></td>
  </tr>
  <tr> 
    <td style="color:#999999;">Cellular:</td>
    <td><%=cCell%></td>
  </tr>
  <tr> 
    <td style="color:#999999;">Fax:</td>
    <td><%=cFax_phone%></td>
  </tr>
  <tr> 
    <td style="color:#999999;">Email:</td>
    <td><%=email%></td>
  </tr>
  <tr> 
    <td style="color:#999999;">Address:</td>
    <td><%=cAddress_street%><br><%=cAddress_floor%></td>
  </tr>
  <tr> 
    <td style="color:#999999;">&nbsp;</td>
    <td><input type="button" value=" Edit " onclick="opencontact('<%=ccontactid%>', '<%=company%>','edit')"></td>
  </tr>
  <tr><td height="8"></td><td></td></tr>
  <%
  rst.movenext
  contactmsg =""
  wend
  rst.close
  %>
  </table>
  </div>
  </td>
</tr>
<tr bgcolor="#dddddd"><td align="right" style="border-top:1px solid #999999;"><input type="button" value="Close Window" onclick="window.close();opener.window.document.location=opener.window.document.location;"></td></tr>
</table>    
<br>

    
</form>
</body>
</html>






