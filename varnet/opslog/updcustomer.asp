<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%


Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")
cid=Request.Querystring("custid")
'response.write cid

sqlstr = "select * from customers where customerid="& cid&""

'response.write sqlstr
'response.end
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
<form name="form1" method="post" action="custupdate.asp">
<table width="100%" border="0">
  <tr> 
      <td bgcolor="#3399CC" height="30"> 
        <table width="100%" border="0" height="33">
          <tr> 
            <td width="73%" height="29"><b><i><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">New Customer
			<input type="hidden" name="cid" value="<%=rst1("customerid")%>"></font></i></b></td>
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
              <td width="33%"><font face="Arial, Helvetica, sans-serif">Company 
                Name:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Contact 
                First Name:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Contact 
                Last Name:</font></td>
          </tr>
          <tr> 
              <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
               <input type="text" name="CompanyName" value="<%=rst1("companyname")%>">
              </font></td>
			  
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
			  <input type="text" name="first" value="<%=rst1("contactfirstname")%>">
                
                </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif">
			   <input type="text" name="last" value="<%=rst1("contactlastname")%>"></font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Contact 
                Title: </font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Billing 
                Address: </font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">City:</font></td>
          </tr>
          <tr> 
            <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="title" value="<%=rst1("contacttitle")%>">
              </font></td>
            <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="addr" value="<%=rst1("billingaddress")%>">
              </font></td>
            <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="city" value="<%=rst1("city")%>">
              </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="33%"><font face="Arial, Helvetica, sans-serif">State:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Zip Code:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Country:</font></td>
          </tr>
          <tr> 
            <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="state" value="<%=rst1("stateorprovince")%>">
              </font></td>
            <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="zip" value="<%=rst1("postalcode")%>">
              </font></td>
            <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="country" value="<%=rst1("country")%>"> 
              </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Phone 
                Number: </font></td>
              <td width="63%"><font face="Arial, Helvetica, sans-serif">Fax Number:</font></td>
              <td width="63%"><font face="Arial, Helvetica, sans-serif">Email Address</font></td> 
          </tr>
          <tr> 
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="phone" value="<%=rst1("phonenumber")%>">
              </font></td>
              <td width="63%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="fax" value="<%=rst1("faxnumber")%>">
              </font></td>
              <td width="63%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="email" value="<%=rst1("email")%>">
              </font></td>
          </tr>
        </table>
            
         
          
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="submit" name="updateit" value="UPDATE" >
		  
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
