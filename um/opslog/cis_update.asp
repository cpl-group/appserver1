<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
Select case trim(request("mode"))
	Case "update"
			dim sql
			cName 	= left(request("cname"),30) 'name
			cType	= request("ctype") 'customer_type
			cTrade 	= request("ctrade") 'trade
			cStatus = request("cstatus") 'status
			cAddress_street = request("cAddress_street") 'address_1
			cAddress_floor = request("cAddress_floor") 'address_2
			cCity	= request("ccity") 'city 
			cstate	= request("cstate") 'cstate 
			czip 	= request("czip") 'zip_code
			cPhone = request("cPhone") 'telephone
			cFax_phone = request("cfax_phone") 'fax
			date_established = request("date_established") ' date_established
			cMaincontact = request("cMaincontact") 'contact_1
			bAddress_street = request("baddress_street") 'billing_address_1
			bAddress_floor = request("bAddress_floor ") 'billing_address_2
			bCity	= request("bcity") 'billing_city
			bstate	= request("bstate") 'billing_state
			bzip	= request("bzip") 'billing_zip_code
			bMainContact= request("bMainContact") 'billing_contact
			company = request("company")
			cid 	= request("cid")	
			set cnn = server.createobject("ADODB.connection")
			set rst = server.createobject("ADODB.recordset")
			cnn.open application("cnnstr_main")
			
			sql = "Update " & company & "_MASTER_ARM_CUSTOMER set name='"&cname&"', customer_type='" & ctype & "', address_1='"&cAddress_street &"', address_2='" & caddress_floor & "', city='" & cCity & "', state='"&cstate & "', zip_code='"&czip&"', billing_address_1='" & bAddress_street &"', billing_address_2='" & baddress_floor & "', billing_city='" & bCity & "', billing_state='"&bstate & "', billing_zip_code='"&bzip&"', status='" & cstatus &"',credit_limit=0,days_before_due=0 where customer='"&cid&"'"

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
		cnn.open application("cnnstr_main")
		
		dim cid, company,cName,cType,cTrade,cStatus, cAddress_street,cAddress_floor,cPhone ,cFax_phone,date_established,cMaincontact, bAddress_street,bAddress_floor,bCity,bstate,bzip,bMainContact, tcolor, cTitle, cContactID,ccity,cstate,czip
		
		cid = Request("cid")
		company = Request("company")
		
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
		</script>
    <link rel="Stylesheet" href="../../GENERGY2_INTRANET/styles.css" type="text/css">		
		</head>
		<body bgcolor="#eeeeee" onload="">
		<form name="form1" method="post" action="../../GENERGY2_INTRANET/OPSMANAGER/JOBLOG/cis_update.asp">
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
        <input name="cName" type="text" value="<%=cName%>" size="30" maxlength="30"> (<%=cid%>)
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
        <input type="submit" value="Update"> &nbsp;<input type="button" value="Cancel" onclick="closewin('<%=cid%>', '<%=company%>');">
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
		company = request("company") 
		set cnn = server.createobject("ADODB.connection")
		set rst = server.createobject("ADODB.recordset")
		cnn.open application("cnnstr_main")
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
		function checkform() {
		  if(document.form1.cType.selectedIndex==0) {
		    alert('Select Customer Type')
		  }
		  else {
		    if(!document.form1.company[0].checked &&!document.form1.company[1].checked){
			  alert('Select Company')
			  }
			else {
		    document.form1.submit()
			}
		  }
		}


		</script>
    <link rel="Stylesheet" href="../../GENERGY2_INTRANET/styles.css" type="text/css">		
    </head>
		<body bgcolor="#eeeeee" marginwidth=0 marginheight=0 topmargin=0 leftmargin=0>
		<form name="form1" method="post" action="../../GENERGY2_INTRANET/OPSMANAGER/JOBLOG/cis_update.asp">
		
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc">
      <td><span class="standardheader">New Customer</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
	 	<td>
		Company&nbsp;
        <input type="radio" name="company" value="GY" onClick="screencompany('GY')"<%if company="GY" then response.Write(" CHECKED") end if%>>gEnergy &nbsp;
        <input type="radio" name="company" value="IL" onClick="screencompany('IL')"<%if company="IL" then response.Write(" CHECKED") end if%>>I-Lite
		</td>
      </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-bottom:1px solid #cccccc;">
      <table border=0 cellpadding="3" cellspacing="0">
      <tr>
        <td>Customer Name</td>
        <td>
        <input name="cName" type="text" size="30" maxlength="30">
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
        <input type="button" value="Save" onClick="checkform()"> &nbsp;<input type="button" value="Cancel" onclick="closewin();">
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
				cName 	= request("cName") 'name
				cType	= request("cType") 'customer_type
				cTrade 	= request("cTrade") 	'trade
				cStatus = request("cStatus")	'status
				cAddress_street = request("cAddress_street") 'address_1
				cAddress_floor = request("cAddress_floor") 'address_2
				cCity	= request("cCity") 'city
				cstate	= request("cState") 	'state
				czip 	= request("czip") 'zip_code
				date_established = request("date_established") 	'date_established
				bAddress_street = request("bAddress_street") 	'billing_address_1
				bAddress_floor = request("bAddress_floor") 	'billing_address_2
				bCity	= request("bCity")	'billing_city
				bstate	= request("bstate") 'billing_state
				bzip	= request("bzip") 	'billing_zip_code
				
				company = request("company")
				
				set cnn = server.createobject("ADODB.connection")
				set rst = server.createobject("ADODB.recordset")
				cnn.open application("cnnstr_main")
				
				
				strsql = "insert into " & company & "_MASTER_ARM_CUSTOMER (name,customer_type, trade, address_1,address_2,city, [state], zip_code, date_established, billing_address_1, billing_address_2, billing_city, billing_state, billing_zip_code, status) values ('"&cName&"', '"&cType&"','NONE','"&cAddress_street&"','"&cAddress_floor&"','"&ccity&"', '"&cstate&"','" & czip & "','" &date()& "','"&bAddress_street&"','"&bAddress_floor&"','"&bcity&"', '"&bstate&"','" & bzip & "','"& cStatus & "')"
				response.Write(strsql)
				response.end
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





