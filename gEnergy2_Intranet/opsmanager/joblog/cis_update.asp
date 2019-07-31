<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
Select case trim(secureRequest("mode"))
	Case "update"
			dim sql
			cName 	= left(secureRequest("cname"),30) 'name
			cType	= secureRequest("ctype") 'customer_type
			cTrade 	= secureRequest("ctrade") 'trade
			cStatus = secureRequest("cstatus") 'status
			cAddress_street = secureRequest("cAddress_street") 'address_1
			cAddress_floor = secureRequest("cAddress_floor") 'address_2
			cCity	= secureRequest("ccity") 'city 
			cstate	= secureRequest("cstate") 'cstate 
			czip 	= secureRequest("czip") 'zip_code
			cPhone = secureRequest("cPhone") 'telephone
			cFax_phone = secureRequest("cfax_phone") 'fax
			date_established = secureRequest("date_established") ' date_established
			cMaincontact = secureRequest("cMaincontact") 'contact_1
			bAddress_street = secureRequest("baddress_street") 'billing_address_1
			bAddress_floor = secureRequest("baddress_floor") 'billing_address_2
			bCity	= secureRequest("bcity") 'billing_city
			bstate	= secureRequest("bstate") 'billing_state
			bzip	= secureRequest("bzip") 'billing_zip_code
			bMainContact= secureRequest("bMainContact") 'billing_contact
			company = secureRequest("company")
			cid 	= secureRequest("cid")	
			projid = secureRequest("projid")
			set cnn = server.createobject("ADODB.connection")
			set rst = server.createobject("ADODB.recordset")
			cnn.open getConnect(0,0,"intranet")
			
			sql = "Update " & company & "_MASTER_ARM_CUSTOMER set name='"&cname&"', customer_type='" & ctype & "', address_1='"&cAddress_street &"', address_2='" & caddress_floor & "', city='" & cCity & "', state='"&cstate & "', zip_code='"&czip&"', billing_address_1='" & bAddress_street &"', billing_address_2='" & baddress_floor & "', billing_city='" & bCity & "', billing_state='"&bstate & "', billing_zip_code='"&bzip&"', status='" & cstatus &"', acct_manager="&projid&" ,credit_limit=0,days_before_due=0 where customer='"&cid&"'"
			cnn.Execute sql
			%>
			<script>
			alert("Update saved.")
			document.location = "<%="cis_detail.asp?cid=" & cid &"&company=" & company%>"
			</script>
			<%
			
case "edit"
		
		dim cnn, rst, strsql
		set cnn = server.createobject("ADODB.connection")
		set rst = server.createobject("ADODB.recordset")
		cnn.open getConnect(0,0,"intranet")
		
		dim cid, company,cName,cType,cTrade,cStatus, cAddress_street,cAddress_floor,cPhone ,cFax_phone,date_established,cMaincontact, bAddress_street,bAddress_floor,bCity,bstate,bzip,bMainContact, tcolor, cTitle, cContactID,ccity,cstate,czip, projid
		
		cid = secureRequest("cid")
		company = secureRequest("company")
		
		if trim(cid)<>"" then
			rst.Open "SELECT * FROM " & company & "_MASTER_ARM_CUSTOMER WHERE customer='"&cid&"'", cnn
			if not rst.EOF then
				cName 	= rst("name")
				cType	= rst("customer_type")
				cTrade 	= rst("trade")
				cStatus = rst("status")
				Select Case cStatus
					case "Active"
						tcolor = "#33FF00"
					case "Inactive"
						tcolor = "#FFFF00"
				end select 
				cAddress_street = rst("address_1")
				cAddress_floor = rst("address_2")
				cCity	= rst("city") 
				cstate	= rst("state") 
				czip 	= rst("zip_code")
				cPhone = rst("telephone")
				cFax_phone = rst("fax")
				date_established = rst("date_established")
				cMaincontact = rst("contact_1")
				bAddress_street = rst("billing_address_1")
				bAddress_floor = rst("billing_address_2")
				bCity	= rst("billing_city") 
				bstate	= rst("billing_state")
				bzip	= rst("billing_zip_code")	
				bMainContact= rst("billing_contact")
				projid = rst("acct_manager")
			end if
			rst.close
		end if
		%>
		<html>
		<head>
		<title>Customer Details - UPDATE</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<script>
		function closewin(cid, company)
		{
			//window.resizeTo(600,480)
		if (confirm("Cancel changes?")){
			theURL="cis_detail.asp?cid=" + cid + "&company=" + company
			document.location = theURL	
		}

		}
		
		function checkform(frm){
			var err = "";
			if(frm.company.value==0) err+="Select company\n";
			if(frm.cName.value=='') err+="No customer name entered\n";
			if(frm.cAddress_street.value=='') err+="No customer address entered\n";
			if(frm.bAddress_street.value=='') err+="No billing address entered\n";
			if(frm.cCity.value=='') err+="No customer city entered\n";
			if(frm.bCity.value=='') err+="No billing city entered\n";
			if(frm.cstate.value=='') err+="No customer state entered\n";
			if(frm.czip.value=='') err+="No customer zip code entered\n";
			if(frm.bstate.value=='') err+="No billing state entered\n";
			if(frm.bzip.value=='') err+="No billing zip code entered\n";
			if(err=="")	frm.submit();
			else alert(err);
		}
		</script>
    <link rel="Stylesheet" href="../../styles.css" type="text/css">		
		</head>
		<body bgcolor="#eeeeee" onload="">
		<form name="form1" method="post" action="cis_update.asp">
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc">
      <td><span class="standardheader">Update Customer</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-bottom:1px solid #cccccc;">
      <table border=0 cellpadding="3" cellspacing="0">
      <tr>
        <td>Customer Name</td>
        <td>
        <input name="cName" type="text" value="<%=cName%>" size="30" maxlength="255"> (<%=cid%>)
        <select name="cstatus" id="cstatus">
        <option value="Inactive" <% if trim(cStatus) = "Inactive" then %> selected <%end if %>>Inactive</option>
        <option value="Active" <% if trim(cStatus) = "Active" then %> selected <%end if %>>Active</option>
        </select>
        </td>
      </tr>
      <tr valign="middle" bgcolor="#eeeeee">
        <td>Customer Type</td>
		<td><select name="cType">
        <%
        rst.Open "SELECT customer_type FROM CUSTOMER_TYPE order by customer_type", cnn
        do until rst.eof 
          if trim(rst("customer_type")) = trim(cType) then
          %>
        <option value="<%=trim(rst("customer_type"))%>" selected><%=rst("customer_type")%></option>
          <%
          else
          %>
        <option value="<%=trim(rst("customer_type"))%>"><%=rst("customer_type")%></option>
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
	  	    <td>Account Manager</td>
	  	<td>
		<% if trim(company) <> "" then %>
        <select name="projid">
        <option value="0" selected>Select Account Manager</option>
        <%
        rst.Open "select * from Managers where companycode = '"&trim(company)&"' and Acctmanager='1' order by lastname, firstname", cnn
        do until rst.eof%>
        <option value="<%=rst("mid")%>" <%If trim(projid)=trim(rst("mid")) then%>selected<%end if%>><%=rst("lastname")%>, 
        <%=rst("firstname")%></option>
        <%
        rst.movenext
        loop
        rst.close
        %>
        </select>
		<%end if%>
	  </td></tr>
      </table>
      </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-top:1px solid #ffffff;">
      
      <table border=0 cellpadding="3" cellspacing="0">
      <tr>
        <td></td>
        <td><b>Main Address</b></td>
        <td width="30" rowspan="6">&nbsp;</td>
        <td></td>
        <td><b>Billing Address</b></td>
      </tr>
      <tr>
        <td align="right">Street</td>
        <td><input name="cAddress_street" type="text" value="<%=cAddress_street%>" size="30" maxlength="30"></td>
        <td align="right">Street</td>
        <td><input name="bAddress_street" type="text" value="<%=bAddress_street%>" size="30" maxlength="30"></td>
      </tr>
      <tr>
        <td></td>
        <td><input name="cAddress_floor"  type="text" value="<%=cAddress_floor%>" size="30" maxlength="30"></td>
        <td></td>
        <td><input name="bAddress_floor"  type="text" value="<%=bAddress_floor%>" size="30" maxlength="30"></td>
      </tr>
      <tr>
        <td align="right">City</td>
        <td><input name="cCity" type="text" value="<%=cCity%>" size="15" maxlength="15"></td>
        <td align="right">City</td>
        <td><input name="bCity" type="text" value="<%=bCity%>" size="15" maxlength="15"></td>
      </tr>
      <tr>
        <td align="right">State</td>
        <td><input name="cstate" type="text" value="<%=cstate%>" size="4" maxlength="4"> &nbsp;&nbsp;Zip: <input name="czip" type="text" value="<%=czip%>" size="10" maxlength="10"></td>
        <td align="right">State</td>
        <td><input name="bstate" type="text" value="<%=bstate%>" size="4" maxlength="4"> &nbsp;&nbsp;Zip: <input name="bzip" type="text" value="<%=bzip%>" size="10" maxlength="10"></td>
      </tr>
      <tr><td colspan="4" height="8"></td></tr>
      <tr bgcolor="#eeeeee">
        <td></td>
        <td colspan="3">
        <input type="hidden" name="mode" value="update">
        <input type="hidden" name="company" value="<%=company%>">
		<input type="hidden" name="cid" value="<%=cid%>">
        <input type="button" value="Update" onclick="checkform(document.forms[0])"> &nbsp;<input type="button" value="Cancel" onclick="closewin('<%=cid%>', '<%=company%>');">
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
case "new"		
		company = secureRequest("company") 'rsm took out secure request and set to blank to force 'select company'
		set cnn = server.createobject("ADODB.connection")
		set rst = server.createobject("ADODB.recordset")
		cnn.open getConnect(0,0,"intranet")
	%>
		
		<html>
		<head>
		<title>New Customer</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<script>
		function closewin()
		{
			window.close()
		}
		function copymain()
		{
			document.form1.bAddress_street.value = document.form1.cAddress_street.value
			document.form1.bAddress_floor.value = document.form1.cAddress_floor.value 
			document.form1.bCity.value = document.form1.cCity.value 
			document.form1.bstate.value = document.form1.cstate.value  
			document.form1.bzip.value = document.form1.czip.value 
		}
		function screencompany(company) {
  		  document.location.href="cis_update.asp?mode=new&company="+company	
		}

		function checkform(frm){
			var err = "";
			if(frm.cName.value=='') err+="Select company name\n";
			if(frm.cAddress_street.value=='') err+="No company address entered\n";
			if(frm.bAddress_street.value=='') err+="No billing address entered\n";
			if(frm.cCity.value=='') err+="No company city entered\n";
			if(frm.bCity.value=='') err+="No billing city entered\n";
			if(frm.cstate.value=='') err+="No company state entered\n";
			if(frm.czip.value=='') err+="No company zip code entered\n";
			if(frm.bstate.value=='') err+="No billing state entered\n";
			if(frm.bzip.value=='') err+="No billing zip code entered\n";
			if(err=="")	frm.submit();
			else alert(err);
		}
		</script>
    <link rel="Stylesheet" href="../../styles.css" type="text/css">		
    </head>
		<body bgcolor="#eeeeee" marginwidth=0 marginheight=0 topmargin=0 leftmargin=0>
		<form name="form1" method="post" action="cis_update.asp">
		
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc">
      <td><span class="standardheader">New Customer</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
	 	<td>
		Company&nbsp;
		<select name="company" onchange="screencompany(this.value)">
		<%if trim(company) = "" then %>
		<option value="">Select Company</option>
		<% end if %>
                <%
        rst.Open "select * from companycodes where active = 1 order by name", cnn
        if not rst.eof then
        do until rst.eof
        %>
		<option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><font face="Arial, Helvetica, sans-serif"><%=rst("name")%></font></option>
                <%    
        rst.movenext
        loop
        end if
        rst.close%>
              </select>		</td>
      </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-bottom:1px solid #cccccc;">
      <table border=0 cellpadding="3" cellspacing="0">
      <tr>
        <td>Customer Name</td>
        <td>
        <input name="cName" type="text" size="30" maxlength="255">
        <select name="cstatus" id="cstatus">
        <option value="Inactive" >Inactive</option>
        <option value="Active" selected>Active</option>
        </select>
        </td>
      </tr>
      <tr valign="middle" bgcolor="#eeeeee">
        <td>Customer Type</td>
		<td><select name="cType">
		<option value="None">Select Customer Type</option>
        <%
        rst.Open "SELECT customer_type FROM CUSTOMER_TYPE order by customer_type", cnn
        do until rst.eof 
          %>
        <option value="<%=trim(rst("customer_type"))%>"><%=rst("customer_type")%></option>
          <%
          rst.movenext
        loop
        rst.close
        %>
        </select>
		</td>
      </tr>
	  <tr>
	  	<td>Project Manager</td>
	  	<td>
		<% if trim(company) <> "" then %>
        <select name="projid">
        <option value="0" selected>Select Account Manager</option>
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
	  </td></tr>
      </table>
      </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-top:1px solid #ffffff;">
      
      <table border=0 cellpadding="3" cellspacing="0">
      <tr>
        <td></td>
        <td><b>Main Address</b></td>
        <td width="30" rowspan="6">&nbsp;</td>
        <td></td>
        <td><b>Billing Address</b>&nbsp;&nbsp;&nbsp;<input name="copy_main" type="checkbox" value="" onclick="copymain()">&nbsp;Same as main address</td>
      </tr>
      <tr>
        <td align="right">Street</td>
        <td><input name="cAddress_street" type="text" size="30" maxlength="30"></td>
        <td align="right">Street</td>
        <td><input name="bAddress_street" type="text" size="30" maxlength="30"></td>
      </tr>
      <tr>
        <td></td>
        <td><input name="cAddress_floor"  type="text" size="30" maxlength="30"></td>
        <td></td>
        <td><input name="bAddress_floor"  type="text" size="30" maxlength="30"></td>
      </tr>
      <tr>
        <td align="right">City</td>
        <td><input name="cCity" type="text" size="15" maxlength="15"></td>
        <td align="right">City</td>
        <td><input name="bCity" type="text" size="15" maxlength="15"></td>
      </tr>
      <tr>
        <td align="right">State</td>
        <td><input name="cstate" type="text" size="4" maxlength="4"> &nbsp;&nbsp;Zip: <input name="czip" type="text" size="10" maxlength="10"></td>
        <td align="right">State</td>
        <td><input name="bstate" type="text" size="4" maxlength="4"> &nbsp;&nbsp;Zip: <input name="bzip" type="text" size="10" maxlength="10"></td>
      </tr>
      <tr><td colspan="4" height="8"></td></tr>
      <tr bgcolor="#eeeeee">
        <td></td>
        <td colspan="3">
        <input type="hidden" name="mode" value="save">
        <input type="button" value="Save" onClick="checkform(document.forms[0])"> &nbsp;<input type="button" value="Cancel" onclick="closewin();">
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
	case "save"
				cName 	= secureRequest("cName") 'name
				cType	= secureRequest("cType") 'customer_type
				cTrade 	= secureRequest("cTrade") 	'trade
				cStatus = secureRequest("cStatus")	'status
				cAddress_street = secureRequest("cAddress_street") 'address_1
				cAddress_floor = secureRequest("cAddress_floor") 'address_2
				cCity	= secureRequest("cCity") 'city
				cstate	= secureRequest("cState") 	'state
				czip 	= secureRequest("czip") 'zip_code
				date_established = secureRequest("date_established") 	'date_established
				bAddress_street = secureRequest("bAddress_street") 	'billing_address_1
				bAddress_floor = secureRequest("bAddress_floor") 	'billing_address_2
				bCity	= secureRequest("bCity")	'billing_city
				bstate	= secureRequest("bstate") 'billing_state
				bzip	= secureRequest("bzip") 	'billing_zip_code
				projid =  secureRequest("projid")
				company = secureRequest("company")
				
				set cnn = server.createobject("ADODB.connection")
				set rst = server.createobject("ADODB.recordset")
				cnn.open getConnect(0,0,"intranet")
				
				
				strsql = "insert into " & company & "_MASTER_ARM_CUSTOMER (name,customer_type, trade, address_1,address_2,city, [state], zip_code, date_established, billing_address_1, billing_address_2, billing_city, billing_state, billing_zip_code, status, acct_manager) values ('"&cName&"', '"&cType&"','NONE','"&cAddress_street&"','"&cAddress_floor&"','"&ccity&"', '"&cstate&"','" & czip & "','" &date()& "','"&bAddress_street&"','"&bAddress_floor&"','"&bcity&"', '"&bstate&"','" & bzip & "','"& cStatus & "', "&projid&")"
						cnn.Execute strsql
				strsql = "select customer from " & company & "_MASTER_ARM_CUSTOMER where ltrim(name) = '" & trim(cName) & "' and ltrim(address_1)='"&trim(cAddress_street) &"'"
				rst.open strsql, cnn,1
				if rst.recordcount = 1 then 
					cid = rst("customer")
				else
					Response.write "ERROR: NO CUSTOMER ID WAS ASSIGNED TO THIS CUSTOMER."
					response.end
				end if 
						%>
						<script>
						function opencontact(contactid,company, mode) 
						{
							theURL="updatecontact.asp?contactid=" + contactid + "&company=" + company + "&mode=" + mode+"&newcust=yes"
							openwin(theURL,500,410)
						}
						function openwin(url,mwidth,mheight)
						{
						window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
						}
						alert("New Customer Saved. Please be sure to complete the New Contact Form.")
						opencontact('<%=cid%>', '<%=company%>','new')
						window.close()
						</script>
<%				
	Case else
end select
%>





