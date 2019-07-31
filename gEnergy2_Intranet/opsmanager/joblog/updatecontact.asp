<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim company
company = secureRequest("company")
if company="" then company = "EM"
Select case trim(secureRequest("mode"))
  Case "edit"
  %>
    <%
    dim cnn, rst, strsql
    set cnn = server.createobject("ADODB.connection")
    set rst = server.createobject("ADODB.recordset")
    cnn.open getConnect(0,0,"intranet")
    
    dim cid, cName,cType,cTrade,cStatus, cAddress_street,cAddress_floor,ccity,cstate,czip,cPhone,cCell, cFax_phone,date_established,cMaincontact, bAddress_street,bAddress_floor,bCitystatezip,bMainContact, tcolor, cTitle, cContactID, customer, email, primarycontact, billingcontact
    
    ccontactid = secureRequest("contactid")
    company= secureRequest("company")
    
          rst.Open "SELECT * FROM " & company & "_STANDARD_ARS_CONTACT WHERE contact ='"&ccontactid&"'", cnn      
          if not rst.eof then 
            cName     = rst("contact_name")
            cType     = rst("contact_type")
            cTitle    = rst("title")
            cContactID  = rst("contact")
            email     = rst("email_address")
            cAddress_street = rst("address_1")
            cAddress_floor = rst("address_2")
            cCity = rst("city")
            cstate  = rst("state") 
            czip = rst("zip_code")
            cPhone = rst("telephone")
            cCell = rst("Cellular")'1/30/2008 added by N.Ambo
            cFax_phone = rst("fax")
            customer = rst("customer")
          else 
            response.write "ERROR: SYSTEM FAILED TO LOCATE CONTACT"
            response.end
          end if 
          rst.close
          rst.Open "select contact_1, billing_contact from " & company & "_MASTER_ARM_CUSTOMER where customer ='" & customer&"'", cnn
       
		  if not rst.EOF then 
            primarycontact = rst("contact_1")
            billingcontact = rst("billing_contact")           
          end if 
          rst.close
          %>
    
    <html>
    <head>
    <title>Update : <%=cName%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script>
    function editcustomer(cid, company)
    {
      pageURL = "cis_update.asp?cid=" + cid + "&company=" + company
      document.location = pageURL
      //window.resizeTo(600,300)
    }
    function closepage()
    {
      if (confirm("Cancel changes?")){
        window.close()
      }
    }

	function checkform(frm)
	{ var err = "";
	  if(frm.cType.value=='') err+="Select Contact Type\n";
		if(frm.cName.value=='') err+="No name entered\n";
		if(frm.cphone.value=='') err+="No phone number entered\n";
		if(frm.cAddress_street.value=='') err+="No address entered\n";
		if(frm.cCity.value=='') err+="No city entered\n";
		if(frm.cstate.value=='') err+="No state entered\n";
		if(frm.czip.value=='') err+="No zip code entered\n";
		if(err=="")	frm.submit();
		else alert(err);
		 
	}
    </script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">   
    </head>
    <body bgcolor="#dddddd">
    <form name="form1" method="post" action="updatecontact.asp">
    
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc">
      <td><span class="standardheader">Update Customer Contact</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-bottom:1px solid #cccccc;">
      <table border=0 cellpadding="3" cellspacing="0" width="100%">
      <tr>
        <td width="18%"></td>
        <td width="82%">
        <input type="checkbox" name="contact_1" value="1" <%if trim(cContactid)=trim(primarycontact) then%> checked <%end if %>>
        Primary Contact 
        <input type="checkbox" name="billingcontact" value="1" <%if trim(cContactid)=trim(billingcontact)  then%> checked <%end if %>>
        Billing Contact
        </td>
      </tr>
      <tr>
        <td>Contact Name</td>
        <td><input name="cName" type="text" value="<%=cName%>" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td>Company:</td>
        <td>
          <select name="customer">
            <%
        rst.Open "SELECT distinct customer,name FROM " & company & "_MASTER_ARM_CUSTOMER order by name", cnn
        do until rst.eof 
          if trim(rst("customer")) = trim(customer) then
        %>
            <option value="<%=trim(rst("customer"))%>" selected><%=left(rst("name"),30)%></option>
            <%
          else
        %>
            <option value="<%=trim(rst("customer"))%>"><%=left(rst("name"),30)%></option>
            <%
          end if
        rst.movenext
        loop
        rst.close
        %>
          </select>
           </td>
      </tr>
      <tr> 
        <td>Title:</td>
        <td><input name="cTitle" type="text" value="<%=cTitle%>" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td>Contact Type:</td>
        <td><input name="cType" type="text" value="<%=cType%>" size="20" maxlength="20"></td>
      </tr>
      <tr> 
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr> 
        <td>Telephone:</td>
        <td>
			<input name="cphone" type="text" value="<%=cphone%>" size="15" maxlength="10">
			Cellular:&nbsp;<input name="ccell" type="text" value="<%=ccell%>" size="15" maxlength="10">
		</td>
      </tr>
      <tr> 
        <td>Fax:</td>
        <td><input name="cFax_phone" type="text" value="<%=cFax_phone%>" size="15" maxlength="10"></td>
      </tr>
      <tr> 
        <td>Email:</td>
        <td><input name="email" type="text" id="email" value="<%=email%>" size="25" maxlength="50"></td>
      </tr>
      <tr><td></td><td><b>Address</b></td></tr>
      <tr> 
        <td>Street:</td>
        <td><input name="cAddress_street" type="text" value="<%=cAddress_street%>" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td>&nbsp;</td>
        <td><input name="cAddress_floor" type="text" value="<%=cAddress_floor%>" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td>City</td>
        <td>
          <input name="cCity" type="text" value="<%=cCity%>" size="15" maxlength="15">
          State:&nbsp;<input name="cstate" type="text" value="<%=cstate%>" size="2"  maxlength="4">
          &nbsp;&nbsp;Zip:&nbsp;<input name="czip" type="text" value="<%=czip%>" size="10" maxlength="10">
          </td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>
        <input type="hidden" name="mode" value="update">
        <input type="hidden" name="company" value="<%=company%>">
        <input type="hidden" name="ccontactid" value="<%=ccontactid%>">
        <input type="button" value="Update" onClick="checkform(document.forms[0])">&nbsp;<input type="button" value="Cancel" onclick="closepage();">
        </td>
      </tr>
      </table>
      </td>
    </tr>
    </table>
    <br>
    
    </form>
    </body>
    </html>
  <%
  case "update" 
            Dim contact_1, contactupdate
            cName     = secureRequest("cName") 'contact_name
            cType   = secureRequest("cType") 'contact_type
            cTitle    = secureRequest("ctitle") 'Title
            cContactID  = secureRequest("ccontactid") 'contact
            cAddress_street = secureRequest("cAddress_street") 'address_1
            cAddress_floor = secureRequest("cAddress_floor") 'address_2
            cCity = secureRequest("ccity") ' city
            cstate  = secureRequest("cstate") 'state
            czip = secureRequest("czip") 'zip_code
            cPhone = left(secureRequest("cPhone"),10) 'telephone
            cCell = left(secureRequest("ccell"),10) 'cellular  1/30/2008 added by N.Ambo
            email = secureRequest("email")
            cFax_phone = left(secureRequest("cFax_phone"),10) 'fax
            customer = secureRequest("customer") 'customer
            company = secureRequest("company")
            
            contact_1 = secureRequest("contact_1")
            billingcontact = secureRequest("billingcontact")
          
		    if billingcontact="" then billingcontact=0
			if contact_1="" then contact_1=0
			
			contactupdate = cint(contact_1) + cint(billingcontact)
            
            set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open getConnect(0,0,"intranet")
            
            strsql = "UPDATE " & company & "_STANDARD_ARS_CONTACT set contact_name='"&cName&"', contact_type='"&cType&"', title='"&cTitle&"', address_1='"&cAddress_street&"', address_2='"&cAddress_floor&"', city='"&ccity&"', state='"&cstate&"', zip_code='" & czip & "',email_address='" & email & "', telephone='"&cPhone&"', cellular='"&cCell&"', fax='"&cFax_phone&"', customer='"&customer&"' WHERE contact='" &  cContactid & "'"
         
            cnn.Execute strsql            
            
            select case contactupdate
              case 2
                  strsql = "update " & company & "_MASTER_ARM_CUSTOMER set contact_1 = '" & cContactid & "', billing_contact='"&cContactid&"' where customer='" & customer & "'"
              case 1
                if cint(contact_1) = 1 then 
                  strsql = "update " & company & "_MASTER_ARM_CUSTOMER set contact_1 = '" & cContactid &"',billing_contact = ' ' where customer = '" & customer &"'"
                else
                  strsql = "update " & company & "_MASTER_ARM_CUSTOMER set billing_contact = '" & cContactid &"',contact_1 = ' ' where customer = '" & customer &"'"
                end if 
              case else    
                  strsql = "update " & company & "_MASTER_ARM_CUSTOMER set contact_1 = ' ', billing_contact=' ' where customer='" & customer & "'"
            end select
			cnn.Execute strsql                  
          %>
            <script>
            alert("Update Saved")
            opener.window.document.location = opener.window.document.location
            window.close()
            </script>
            <%
  case "new"
  %>
    <%
    set cnn = server.createobject("ADODB.connection")
    set rst = server.createobject("ADODB.recordset")
    cnn.open getConnect(0,0,"intranet")
    'When mode=new contactid = customer
    company= secureRequest("company")
    customer = secureRequest("contactid")
	dim newcust
	newcust=secureRequest("newcust")%>
    
	<script>
	function getaddress() {
	alert('Andomic');
	//document.form1.cust_name.value=dd.options[dd.selectedIndex].text
	//if (company != "GE") {
		document.form1.mode.value="new"
		document.form1.submit();
	//}
}
	</script>
	
	
    
    <html>
    <head>
    <title>New Contact</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script>
    function editcustomer(cid, company)
    {
      pageURL = "cis_update.asp?cid=" + cid + "&company=" + company
      document.location = pageURL
      //window.resizeTo(600,300)
    }
    function closepage()
    {
      if (confirm("Cancel changes?")){
        window.close()
      }
    }
	function checkform(frm)
	{ var err = "";
		if(frm.cName[0].value=='') err+="No name entered\n";
		if(frm.cphone.value=='') err+="No phone number entered\n";
		if(frm.cAddress_street.value=='') err+="No address entered\n";
		if(frm.cCity.value=='') err+="No city entered\n";
		if(frm.cstate.value=='') err+="No state entered\n";
		if(frm.czip.value=='') err+="No zip code entered\n";
		if(err=="")	frm.submit();
		else alert(err);
	}
function screencompany(company) {
  document.location.href="updatecontact.asp?mode=new&company="+company	
}

    </script>
    <link rel="Stylesheet" href="../../styles.css" type="text/css">   
    </head> 
    <body bgcolor="#dddddd">
    <form name="form1" method="post" action="updatecontact.asp">
    <input type="hidden" name="newcust" value="<%=newcust%>">
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc">
      <td><span class="standardheader">New Customer Contact</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-bottom:1px solid #cccccc;">
      <table border=0 cellpadding="3" cellspacing="0" width="100%">
      <tr>
        <td width="18%">Company</td>
        <td width="82%">
<select name="company" onchange="screencompany(this.value)">
		<%if trim(company) = "" then %>
		<% end if %>
                <%
        rst.Open "select * from companycodes where code <> 'AC' AND active = 1 order by name", cnn
        if not rst.eof then
        do until rst.eof
        %>
		<option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><font face="Arial, Helvetica, sans-serif"><%=rst("name")%></font></option>
                <%    
        rst.movenext
        loop
        end if
        rst.close%>
              </select>        <input type="checkbox" name="contact_1" value="1">
        Primary Contact &nbsp;
        <input type="checkbox" name="billingcontact" value="1" <%if trim(lcase(company)) = "ge" then%>checked<%end if%>>
        Billing Contact &nbsp;
        </td>
      </tr>

      <tr>
        <td width="18%">Contact Name(Last)</td>
        <td width="82%"><input name="cName" type="text" value="" size="30"  maxlength="30"></td>
      </tr>
      <tr> 
        <td>First Name</td>
        <td><input name="cName" type="text" value="" size="30" maxlength="30"></td>
      </tr>
      <tr>
        <td>Company</td>
        <td>
        <select name="customer">
        <%
		if company<>"AC" then
          rst.Open "SELECT distinct customer,name FROM " & company & "_MASTER_ARM_CUSTOMER order by name", cnn
        
          do until rst.eof 
            if trim(rst("customer")) = trim(customer) then
          %>
          <option value="<%=trim(rst("customer"))%>" selected><%=left(rst("name"),30)%></option>
          <%
            else
          %>
          <option value="<%=trim(rst("customer"))%>"><%=left(rst("name"),30)%></option>
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
      <tr> 
        <td>Title</td>
        <td><input name="cTitle" type="text" value="" size="30"  maxlength="30"></td>
      </tr>
      <tr> 
        <td>Contact Type</td>
        <td>
        <select name="cType">
        <%
        rst.Open "SELECT type FROM contact_type", cnn
        
        do until rst.eof 
        %>
        <option value="<%=trim(rst("type"))%>"><%=left(rst("type"),20)%></option>
        <%
        rst.movenext
        loop
        rst.close
        %>
        </select>
        </td>
      </tr>
      <tr> 
        <td>Telephone</td>
        <td>
			<input name="cphone" type="text" value="" size="15" maxlength="10">
			Cellular:&nbsp;<input name="ccell" type="text" value="" size="15" maxlength="10">
        </td>
      </tr>
      <tr> 
        <td>Fax</td>
        <td><input name="cFax_phone" type="text" value="" size="15" maxlength="10"></td>
      </tr>
      <tr>
        <td>Email</td>
        <td><input name="email" type="text" id="email"  size="25"  maxlength="50"></td>
      </tr>
      <tr>
        <td>Project Manager</td>
        <td>
		<% if trim(company) <> "" then %>
        <select name="projid">
        <option value="none" selected>Select Project Manager</option>
        <%
        rst.Open "select * from Managers where Active is null and companycode = '"&trim(company)&"'order by lastname, firstname", cnn
        do until rst.eof%>
        <option value="<%=rst("mid")%>" <%If trim(request("projid"))=trim(rst("mid")) then%>selected<%end if%>><%=rst("lastname")%>, 
        <%=rst("firstname")%></option>
        <%
        rst.movenext
        loop
        rst.close
        %>
        </select>
		<%end if%>
		</td>
      </tr>
      <tr>
        <td></td>
        <td><br><b>Address</b></td>
      </tr>
      <tr> 
        <td>Street</td>
        <td><input name="cAddress_street" type="text" value="" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td>&nbsp;</td>
        <td><input name="cAddress_floor" type="text" value="" size="30" maxlength="30"></td>
      </tr>
      <tr> 
        <td>City</td>
        <td><input name="cCity" type="text" value="" size="15" maxlength="15">&nbsp;&nbsp;&nbsp;State:&nbsp;<input name="cstate" type="text" value="" size="2" maxlength="4">&nbsp;&nbsp;&nbsp;Zip:&nbsp;<input name="czip" type="text" value="" size="5" maxlength="10"></td>
      </tr>
      <tr>
        <td></td>
        <td>
        <input type="hidden" name="mode" value="save">
        <input type="button" value="Save" onClick="checkform(document.forms[0])">&nbsp;<input type="button" value="Cancel" onclick="closepage();">
        </td>
      </tr>
      </table>
      </td>
    </tr>
    </table>
    
    </form>
    </body>
    </html>
  <%
  case "save"
    
            cName     = left(secureRequest("cName"),30) 'contact_name
            cType   = secureRequest("cType") 'contact_type
            cTitle    = secureRequest("ctitle") 'Title
            cContactID  = secureRequest("ccontactid") 'contact
            cAddress_street = secureRequest("cAddress_street") 'address_1
            cAddress_floor = secureRequest("cAddress_floor") 'address_2
            cCity = secureRequest("ccity") ' city
            cstate  = secureRequest("cstate") 'state
            czip = secureRequest("czip") 'zip_code
            email = secureRequest("email") 'email_address
            cPhone = left(secureRequest("cPhone"),10) 'telephone
            cCell = left(secureRequest("ccell"),10) 'cellular
            cFax_phone = left(secureRequest("cFax_phone"),10) 'fax
            customer = secureRequest("customer") 'customer
            company = secureRequest("company")
            contact_1 = secureRequest("contact_1")
            billingcontact = secureRequest("billingcontact")'rst("billing_contact")           

            set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open getConnect(0,0,"intranet")
            
            strsql = "insert into " & company & "_STANDARD_ARS_CONTACT  (contact_name,contact_type, title, address_1,address_2,city, [state], zip_code, email_address, telephone, cellular, fax, customer) values ('"&cName&"', '"&cType&"','"&cTitle&"','"&cAddress_street&"','"&cAddress_floor&"','"&ccity&"', '"&cstate&"','" & czip & "','"&email&"','"&cPhone&"', '"&cCell&"', '"&cFax_phone&"','"&customer&"')"
            cnn.Execute strsql
            
            
            rst.open "SELECT top 1 * FROM " & company & "_STANDARD_ARS_CONTACT ORDER BY id desc", cnn
            if not rst.eof then cContactid = rst("contact")
            
            if trim(billingcontact)="1" then
              strsql = "update " & company & "_MASTER_ARM_CUSTOMER set billing_contact='"&cContactid&"' where customer='" & customer & "'"
              cnn.Execute strsql
            end if
            if trim(contact_1)="1" then
              strsql = "update " & company & "_MASTER_ARM_CUSTOMER set contact_1='" & cContactid & "' where customer='" & customer & "'"
              cnn.Execute strsql
            end if
            %>
            <script>
            alert("New Contact Saved")
			<%
			if secureRequest("newcust")<>"yes" then
			  response.Write("opener.document.location.reload()")
			end if
			%>
            window.close()
            </script>
            <%
  case else
end select
%>
