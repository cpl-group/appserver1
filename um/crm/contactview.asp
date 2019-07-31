<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
cid=request.querystring("cid")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"Intranet")



sqlstr = "select * from contacts where id='"&cid&"'"

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
<form name="form1" method="post" action="updatecontact.asp">
<table width="100%" border="0">
  <tr> 
      <td bgcolor="#3399CC" height="30"> 
        <table width="100%" border="0" height="33">
          <tr> 
            <td width="73%" height="29"><b><i><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Update 
              Contact </font></i></b></td>
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
              <td width="27%"><font face="Arial, Helvetica, sans-serif">Contact 
                First Name:</font></td>
              <td width="34%"><font face="Arial, Helvetica, sans-serif">Contact 
                Last Name:</font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif">Contact 
                Title:</font></td>
          </tr>
          <tr> 
              <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="first" value="<%=rst1("first_name")%>">
				<input type="hidden" name="cid" value="<%=cid%>">
                </font></td>
			  
              <td width="34%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="last" value="<%=rst1("last_name")%>">
                </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="title" value="<%=rst1("title")%>">
                </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="27%"><font face="Arial, Helvetica, sans-serif"> Company 
                Name:</font></td>
              <td width="34%"><font face="Arial, Helvetica, sans-serif">Billing 
                Address: </font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif">City:</font></td>
          </tr>
          <tr> 
              <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="CompanyName" value="<%=rst1("company")%>">
                </font></td>
              <td width="34%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="addr" value="<%=rst1("address")%>">
              </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="city" value="<%=rst1("city")%>">
              </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="27%"><font face="Arial, Helvetica, sans-serif">State:</font></td>
              <td width="34%"><font face="Arial, Helvetica, sans-serif">Zip Code:</font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif">Country:</font></td>
          </tr>
          <tr> 
              <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="state" value="<%=rst1("state")%>">
              </font></td>
              <td width="34%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="zip" value="<%=rst1("zip")%>" maxlength="5">
              </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="country" value="<%=rst1("country")%>" >
              </font></td>
          </tr>
        
          <tr bgcolor="#CCCCCC"> 
              <td width="27%"><font face="Arial, Helvetica, sans-serif">Phone 
                Number: </font></td>
              <td width="34%"><font face="Arial, Helvetica, sans-serif">Fax Number:</font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif">Email:</font></td>
              
             
          </tr>
          <tr> 
              <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="phone" value="<%=rst1("phone")%>" >
              </font></td>
              <td width="34%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="fax" value="<%=rst1("fax")%>">
			 
              </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="email" value="<%=rst1("email")%>">
			 
              </font></td>
          </tr>
		  
		  <tr bgcolor="#CCCCCC">
		      <TD width="27%"><font face="Arial, Helvetica, sans-serif">Organization</font> 
              </TD>
		      <TD width="34%"><font face="Arial, Helvetica, sans-serif">Client Type</font> 
              </TD>
			  <TD width="39%"><font face="Arial, Helvetica, sans-serif">Referred By:</font> </TD>
		  </TR>
		  <tr >
		      <TD width="27%"> 
                <select name="org">
<%
set rst2 = createobject("ADODB.recordset")
rst2.open "select [id], org from mkt_organizations order by org", cnn1
do until rst2.eof
    response.write "<option value="""& rst2("id") &""""
    if cint(rst2("id"))=cint(rst1("org")) then response.write "selected"
    response.write ">"& rst2("org") &"</option>"
    rst2.movenext
loop
rst2.close
%>
<option value="">Other</option>
</select>
              </TD>
		      <TD width="34%"><font face="Arial, Helvetica, sans-serif"> 
                <select name="orgtype">
    <option value="1" <%if 1=cint(rst1("orgtype")) then response.write "selected"%>>All Members</option>
    <option value="2" <%if 2=cint(rst1("orgtype")) then response.write "selected"%>>Priciple Members</option>
    <option value="3" <%if 3=cint(rst1("orgtype")) then response.write "selected"%>>Associate Members</option>
</select>

			
</TD>
			  <TD width="39%"><font face="Arial, Helvetica, sans-serif"> 
                <select name="ref">
                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select * from mkt_ref"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
					If rst1("referredby")= rst2("referred") then	
		%>
                  <option value="<%=rst2("referred") %>"selected><font face="Arial, Helvetica, sans-serif"><i><b><font color="#FFFFFF"><%=rst2("referred") %></font></b></i></font></option>
                  <%else
				  %>
                  <option value="<%=rst2("referred") %>"><font face="Arial, Helvetica, sans-serif"><i><b><font color="#FFFFFF"><%=rst2("referred") %></font></b></i></font></option>
                  <%
				  end if
					rst2.movenext
					loop
					end if
					rst2.close
				%>
<option value="">Other</option>
                </select></font>
				<input type="text" name="other" value="<%=rst1("otherref")%>" >
</TD>

		  </TR>
          <tr bgcolor="#CCCCCC"> 
              <td width="27%"><font face="Arial, Helvetica, sans-serif">Industry: </font></td>
              <td width="34%"><font face="Arial, Helvetica, sans-serif"><font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
              
             
          </tr>
          <tr> 
              <td width="27%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="industry">
<%
rst2.open "select [id], industry from mkt_industries order by industry", cnn1
do until rst2.eof
    response.write "<option value="""& rst2("id") &""""
    if cint(rst2("id"))=cint(rst1("industry")) then response.write "selected"
    response.write ">"& rst2("industry") &"</option>"
    rst2.movenext
loop
rst2.close
%>                <option value="">Other</option>
                </select>
                            </font></td>
              <td width="34%"> <font face="Arial, Helvetica, sans-serif">&nbsp; 
			 
              </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif">&nbsp; 
			 
              </font></td>
          </tr>
        </table>
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="submit" name="updit" value="UPDATE" >
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
