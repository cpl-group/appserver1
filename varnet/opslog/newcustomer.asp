<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%

dim mkid, cust, cname, custfirst, custlast, title, address, city, state, zip, country, phone, fax

mkid = request.querystring("mkid")
cust = request.querystring("cust")


Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")

if trim(mkid)<>"" and trim(cust)<>"" then
    rst1.open "SELECT * from contacts where id="& cust, cnn1
    if not(rst1.eof) then
        cname = rst1("Company")
        custfirst = rst1("First_name")
        custlast = rst1("Last_name")
        title = rst1("Title")
        address = rst1("Address")
        city = rst1("city")
        state = rst1("state")
        zip = rst1("zip")
        country = rst1("country")
        phone = rst1("phone")
        fax = rst1("fax")
    end if
    rst1.close
end if 

sqlstr = "select max(customerid)+1 as cid from customers"

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then

%>

<table width="100%" border="0">
  <tr>
    <td>
      <div align="center">
        <p><font face="Arial, Helvetica, sans-serif"><i>Customer not found
          - please resubmit query or contact your system administrator </i></font></p>
        <p><font face="Arial, Helvetica, sans-serif"><i>
          <input type="button" name="Button" value="BACK" onclick="Javascript:history.back()">
          </i></font></p>
      </div>
    </td>
  </tr>
</table>
<%
else
%>
<form name="form1" method="post" action="savecustomer.asp">
<table width="100%" border="0">
  <tr> 
      <td bgcolor="#3399CC" height="30"> 
        <table width="100%" border="0" height="33">
          <tr> 
            <td width="73%" height="29"><b><i><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">New Customer
			<input type="hidden" name="cid" value="<%=rst1("cid")%>"></font></i></b></td>
            <td height="29" width="27%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif"><i> 
                <input type="button" name="Button2" value="BACK" onClick="Javascript:history.back()">
                </i></font></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="2"> 
      <div align="left"> 
        <table width="100%" border="0">
          <tr bgcolor="#CCCCCC"> 
              <td width="33%"><font face="Arial, Helvetica, sans-serif">Company Name:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Contact First Name:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Contact Last Name:</font></td>
          </tr>
          <tr> 
              <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
               <input type="text" name="CompanyName" size="40" maxlength="40" value="<%=cname%>">
              </font></td>
			  
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
			  <input type="text" name="first" size="40" maxlength="40" value="<%=custfirst%>">
                
                </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif">
			   <input type="text" name="last" size="40" maxlength="40" value="<%=custlast%>"></font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Contact Title: </font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Billing Address: </font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">City:</font></td>
          </tr>
          <tr> 
            <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="title" size="40" maxlength="40" value="<%=title%>">
              </font></td>
            <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="addr" size="40" maxlength="40" value="<%=address%>">
              </font></td>
            <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="city" value="<%=city%>">
              </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="33%"><font face="Arial, Helvetica, sans-serif">State:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Zip Code:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Country:</font></td>
          </tr>
          <tr> 
            <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="state" value="<%=state%>">
              </font></td>
            <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="zip" value="<%=zip%>">
              </font></td>
            <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="country" value="<%=country%>">
              </font></td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr bgcolor="#CCCCCC"> 
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Phone Number: </font></td>
              <td width="63%"><font face="Arial, Helvetica, sans-serif">Fax Number:</font></td>
              
             
          </tr>
          <tr> 
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="phone" value="<%=phone%>">
              </font></td>
              <td width="63%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="fax" value="<%=fax%>">
			
              </font></td>
             
          </tr>
        </table>
            
         
          
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="submit" name="saveit" value="SAVE" >
		<input type="hidden" name="mkid" value="<%=mkid%>">
          <input type="button" name="cancel" value="CANCEL" onclick='javascript:parent.document.location="oplogindex.asp"'>
          </i></font></div>
    </td>
  </tr>
</table>

</form>
<%
end if
%>
</body>
</html>
