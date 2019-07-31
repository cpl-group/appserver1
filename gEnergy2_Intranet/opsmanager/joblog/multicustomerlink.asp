<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<% 
Dim appmode,customerlist, jid,cnn,cmd, rst, sql, customer, company, cstatus, templist,jstatus, custid

appmode = 	lcase(trim(request("mode")))
jid 	=	request("jid")
company = 	lcase(trim(request("company")))
jstatus =   lcase(trim(request("jstatus")))
custid	=	trim(request("custid"))
		
if jstatus = "closed" then 
	appmode = "justshow"
end if 

set cnn = server.createobject("ADODB.connection")		
cnn.open getConnect(0,0,"intranet")

select case appmode

	case "add"
		customerlist = split(request("MasterList"),",")
				
		for each customer in customerlist
		sql = "insert into CustomerBidTracking (jobid, customerid) values ('"&jid&"','"&trim(customer)&"')"
		cnn.execute sql
		next
		response.redirect "multicustomerlink.asp?jid="&jid&"&company="&company
			
	case "remove","removeall"
		customerlist = split(request("JobList"),",")
		
		if ubound(customerlist) > -1 then 
			for each customer in customerlist
				if templist = "" then 
					templist = "'" & trim(customer)&"'"
				else 
					templist = templist & ",'" & trim(customer)&"'"
				end if
			next
			sql = "delete from CustomerBidTracking where jobid='"&jid&"' and customerid in ("&templist&")"
			cnn.execute sql
		end if
		
		set rst = server.createobject("ADODB.recordset")
		sql = "select * from CustomerBidTracking where jobid = '"&jid&"' and [primary] = 1"
		rst.open sql, cnn
		
		if rst.eof then  
			response.write "<script>"
			response.write "opener.updateCustomer('-1','NO CUSTOMER SELECTED' );"
			response.write "document.location = 'multicustomerlink.asp?jid="&jid &"&company="& company &"';"
			response.write "</script>"
			response.write "done."
			response.end
		end if 
		rst.close
		
		response.redirect "multicustomerlink.asp?jid="&jid&"&company="&company	
	case "setprimary" 
		set rst = server.createobject("ADODB.recordset")
		customerlist = split(request("JobList"),",")
		if ubound(customerlist) > -1 then 
			sql = "update CustomerBidTracking set [primary] = 0 where jobid = '"&jid&"'; update CustomerBidTracking set [primary] = 1 where jobid='"&jid&"' and customerid = '"&customerlist(0)&"'" 
			cnn.execute sql
			sql	= "select name from "&company&"_MASTER_ARM_CUSTOMER where customer='"&customerlist(0)&"'" 
			rst.open sql, cnn
			if not rst.eof then 
				response.write "<script>"
				response.write "opener.updateCustomer('"&customerlist(0)&"','"&rst("name")&"' );"
					if company <> "ge" then
						response.write "opener.getaddress('"&customerlist(0)&"', '"&trim(company)&"');"				
					end if
				response.write "document.location = 'multicustomerlink.asp?jid="&jid &"&company="& company &"';"
				response.write "</script>"
				response.write "done."
				response.end
			end if
		end if
	case "justshow"
		%>
		<html>
		<head>
		<title>Link Multiple Customers to Job</title>
		<link rel="Stylesheet" href="../../styles.css" type="text/css">   
		</head>
		<body>
		<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		</tr>
		<tr> 
		<td colspan=3 align="center" style="border-bottom:2px solid black;">Link Customers 
		to Job</td>
		</tr>
		<tr valign="top"> 
		<td align="center" valign="middle">&nbsp;</td>
		<td align="center" valign="middle">&nbsp;</td>
		<td align="center" valign="middle">&nbsp;</td>
		</tr>
		<tr> 
		<td colspan = 3 align="center" valign="top">
		<%
		set rst = server.createobject("ADODB.recordset")
		sql = "SELECT distinct customer,name, status, [primary] FROM CustomerBidTracking cbt inner join " & company & "_MASTER_ARM_CUSTOMER mac on mac.customer = cbt.customerid where jobid = "&jid&" order by name"
		rst.Open sql, cnn
		do until rst.eof
		cstatus = lcase(trim(rst("status")))
		if cstatus <> "inactive" then
		%>
		<%=left(trim(rst("name")),30)%>&nbsp;&nbsp;&nbsp;<%if rst("primary") then%>[PRI]<%end if%><br>
		<%
		end if 
		rst.movenext
		loop
		rst.close
		%>
		</td>
		</tr>
		<tr> 
		<td align="center">&nbsp;</td>
		<td align="center"> <input id="editjob222" name="editjob222" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="window.close()" value="Close Window"> 
		</td>
		<td align="center">&nbsp;</td>
		</tr>
		</table>
		</body>
		</html>
		<%	
	case else
		%>
		<html>
		<head>
		<title>Link Multiple Customers to Job</title>
		<link rel="Stylesheet" href="../../styles.css" type="text/css">   
		<script>
		function addCustomers(){
			addForm.submit()
		}
		function removeCustomers(){
			removeForm.mode.value = "remove"
			removeForm.submit()
		}
		function selectAll(){
			var ops = document.removeForm.Joblist.options;
			for (var i=0;i<ops.length;i++)
			{
			ops[i].selected = true;
			}
			removeForm.mode.value = "removeall"
			removeForm.submit()
		}
		function setPrimary(){
			removeForm.mode.value = "setprimary"
			removeForm.submit()
		}
		</script>
		</head>
		<body>
		  
		
		<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		</tr>
		<tr> 
		<td colspan=3 align="center" style="border-bottom:2px solid black;">Link Customers 
		to Job</td>
		</tr>
		<tr valign="top"> 
		<td align="center" valign="middle">&nbsp;</td>
		<td align="center" valign="middle">&nbsp;</td>
		<td align="center" valign="middle">&nbsp;</td>
		</tr>
		<tr> 
		<td align="center" valign="top"> <form name="addForm" method="post" action="multicustomerlink.asp">
		<select name="MasterList" size="10" multiple id="MasterList" style="width:200">
		<%
		set rst = server.createobject("ADODB.recordset")
		sql = "SELECT distinct customer,name, status FROM " & company & "_MASTER_ARM_CUSTOMER where customer not in (select customerid from CustomerBidTracking where jobid = '"&jid&"') order by name"
		rst.Open sql, cnn
		do until rst.eof
		cstatus = lcase(trim(rst("status")))
		if cstatus <> "inactive" then
		%>
		<option value="<%=trim(rst("customer"))%>"><%=left(trim(rst("name")),30)%></option>
		<%
		end if 
		rst.movenext
		loop
		rst.close
		%>
		</select>
		<input name="jid" type="hidden" value="<%=jid%>">
		<input name="company" type="hidden" value="<%=company%>">
		<input name="mode" type="hidden" value="add">
		</form></td>
		<td align="center" valign="top"> 
		<input id="editjob3" name="editjob" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onClick="document.all.updating.style.display='block';addCustomers();" value="ADD ">
		<br>
		<br>
		<input id="editjob23" name="editjob2" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onClick="document.all.updating.style.display='block';removeCustomers();" value="REMOVE"><br><br><input id="editjob23" name="editjob2" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onClick="document.all.updating.style.display='block';selectAll();" value="REMOVE ALL"><br><br><label id="updating" style="display:none">Updating...</label>
		</td>
		<td align="center" valign="top"> <form name="removeForm" method="post" action="multicustomerlink.asp">
		<select name="Joblist" size="10" multiple id="Joblist" style="width:200">
		<%
		set rst = server.createobject("ADODB.recordset")
		sql = "SELECT distinct customer,name, status, [primary] FROM CustomerBidTracking cbt inner join " & company & "_MASTER_ARM_CUSTOMER mac on mac.customer = cbt.customerid where jobid = "&jid&" order by name"
		rst.Open sql, cnn
		do until rst.eof
		cstatus = lcase(trim(rst("status")))
		if cstatus <> "inactive" then
		%>
		<option value="<%=trim(rst("customer"))%>" <%if rst("primary") then%>style="background-color:#cccccc;"<%end if%>><%=left(trim(rst("name")),30)%>&nbsp;&nbsp;&nbsp;<%if rst("primary") then%>[PRI]<%end if%> </option>
		<%
		end if 
		rst.movenext
		loop
		rst.close
		%>
		</select>
		<input name="jid" type="hidden" value="<%=jid%>">
		<input name="company" type="hidden" value="<%=company%>">
		<br>
		<input name="mode" type="hidden" value="remove">
		<input id="editjob22" name="editjob22" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="setPrimary();document.all.updating.style.display='block';" value="Set As Primary for Job">
		</form></td>
		</tr>
		<tr> 
		<td align="center">&nbsp;</td>
		<td align="center"> <input id="editjob222" name="editjob222" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="window.close()" value="Close Window"> 
		</td>
		<td align="center">&nbsp;</td>
		</tr>
		</table>
		</body>
		</html>
		<%	
end select
%>
