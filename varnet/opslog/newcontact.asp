<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function check(phone, zip)
{    if((!(phone>0))&&(phone.length!=10))
         alert('a');
     if((!(zip>0))&&(zip.length!=5))
         alert('b');
     document.forms[0].submit();
     return(true);
}
</script>
<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%


Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")



sqlstr = "select max(customerid)+1 as cid from customers"

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then
%>

<table width="100%" border="0">
  <tr>
    <td>
      <div align="center">
        <p><font face="Arial, Helvetica, sans-serif"><i>Customer not found - please resubmit query or contact your system administrator </i></font></p>
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
<form name="form1" method="post" action="savecontact.asp">
<table width="100%" border="0">
  <tr> 
      <td bgcolor="#3399CC" height="30"> 
        <table width="100%" border="0" height="33">
          <tr> 
            <td width="73%" height="29"><b><i><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">New Contact 
             </font></i></b></td>
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
              <td width="33%"><font face="Arial, Helvetica, sans-serif">Contact 
                First Name:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Contact 
                Last Name:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Contact 
                Title:</font></td>
          </tr>
          <tr> 
              <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="first" size="40" maxlength="40">
                </font></td>
			  
              <td width="37%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="last" size="40" maxlength="40">
                </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="title" size="40" maxlength="40">
                </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="37%"><font face="Arial, Helvetica, sans-serif"> Company 
                Name:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Billing 
                Address: </font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">City:</font></td>
          </tr>
          <tr> 
              <td width="33%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="CompanyName" size="40" maxlength="40">
                </font></td>
            <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="addr" size="40" maxlength="40">
              </font></td>
            <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="city" >
              </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="33%"><font face="Arial, Helvetica, sans-serif">State:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Zip Code:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Country:</font></td>
          </tr>
          <tr> 
            <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="state" >
              </font></td>
            <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="zip" maxlength="5">
              </font></td>
            <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
              <input type="text" name="country" value="USA">
              </font></td>
          </tr>
        
          <tr bgcolor="#CCCCCC"> 
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Phone 
                Number: </font></td>
              <td width="63%"><font face="Arial, Helvetica, sans-serif">Fax Number:</font></td>
              <td width="63%"><font face="Arial, Helvetica, sans-serif">Email:</font></td>
              
             
          </tr>
          <tr> 
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="phone" maxlength="10">
              </font></td>
              <td width="63%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="fax" >
			 
              </font></td>
              <td width="63%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="email" >
			 
              </font></td>
          </tr>
		   <tr bgcolor="#CCCCCC">
		      <TD><font face="Arial, Helvetica, sans-serif">Organization</font></TD>
		      <TD><font face="Arial, Helvetica, sans-serif">Contact Type</font> 
              </TD>
			  <TD><font face="Arial, Helvetica, sans-serif">Referred By</font> 
              </TD>
		  </TR>
		  <tr valign="top">
		      <TD>
<select name="org">
<%
set rst2 = createobject("ADODB.recordset")
rst2.open "select [id], org from mkt_organizations order by org", cnn1
do until rst2.eof
    response.write "<option value="""& rst2("id") &""">"& rst2("org") &"</option>"
    rst2.movenext
loop
rst2.close
%>
<option value="">Other</option>
</select>
                <input type="text" name="otherorg">
              </TD>
<TD>
<select name="orgtype">
    <option value="1">All Members</option>
    <option value="2">Principle Members</option>
    <option value="3">Associate Members</option>
</select>
</TD> <TD><font face="Arial, Helvetica, sans-serif">
			  <select name="ref">
                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select * from mkt_ref order by id"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
		%>
                  <option value="<%=rst2("referred") %>"><font face="Arial, Helvetica, sans-serif"><%=rst2("referred") %></font></option>
                 
                  <%
					rst2.movenext
					loop
					end if
					rst2.close
				%>
<option value="">Other</option>
                </select></font>
				<input type="text" name="other"  >
</TD>

		  </TR>
          <tr bgcolor="#CCCCCC"> 
              <td width="27%"><font face="Arial, Helvetica, sans-serif">Industry: </font></td>
              <td width="34%"><font face="Arial, Helvetica, sans-serif"><font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif"></font></td>
              
             
          </tr>
          <tr> 
              <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="industry">
<%
set rst2 = createobject("ADODB.recordset")
rst2.open "select [id], industry from mkt_industries order by industry", cnn1
do until rst2.eof
    response.write "<option value="""& rst2("id") &""">"& rst2("industry") &"</option>"
    rst2.movenext
loop
rst2.close
%>                <option value="">Other</option>
                </select> <input type="text" name="otherindustry">
                            </font></td>
              <td width="34%"> <font face="Arial, Helvetica, sans-serif"> 
			 
              </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif"> 
			 
              </font></td>
          </tr>
        </table>
            
         
          
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="button" name="saveit" value="SAVE" onclick="check(phone.value, zip.value)">
		  
          <input type="button" name="cancel" value="CANCEL" onclick='javascript:parent.document.location="mktindex.asp"'>
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
