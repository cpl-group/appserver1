
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
bldgname= request("bldgname")
bldgnum=request("bldgnum")
floor=request("floor")
riser=request("riser")
pc=request("pc")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=Capacity_db;"
Set rst2 = Server.CreateObject("ADODB.recordset")
str = "select * from tblriser where bldgnum='"& bldgnum&"' and  riser_name='"& riser &"'"
rst2.Open str, cnn1, 0, 1, 1
if not rst2.eof then
sets=rst2("sets")
sw_frame=rst2("sw_frame")
sw_fuse=rst2("sw_fuse")
wire_capacity=rst2("wire_capacity")
metal=Trim(rst2("metal"))
volts=Trim(rst2("volts"))
end if
%>
<form name="form1" method="post" action="capupdate.asp">

<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="2"> 
      <table width="100%" border="0">
        <tr> 
            <td height="2" width="25%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Building 
              # : <%=bldgname%>
			  <input type="hidden" name="bldgname" value="<%=bldgname%>"> 
              <input type="hidden" name="bldgnum" value="<%=bldgnum%>">
			  <input type="hidden" name="floor" value="<%=floor%>">
              </font></b></i></td>
			  
            <td height="2" width="16%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"> 
              Floor: <%=floor%> </font></b></i></td>
			  
            <td height="2" width="43%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"> 
              <% 
			  Set rst1 = Server.CreateObject("ADODB.recordset")
			  sql="select sqft from tblfloor where bldgnum='" & bldgnum &"' and fl_name='"& floor&"'"
			  rst1.Open sql, cnn1, 0, 1, 1
			  if not rst1.eof then
			  %>
              SQFT: <%=rst1("sqft")%> 
              <input type="hidden" name="sqft" value="<%=rst1("sqft")%>">
              </font></b></i></td>
			  <% 
			  end if
			  rst1.close
			  %>
            <td height="2" width="16%"> 
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
              <td width="33%"><font face="Arial, Helvetica, sans-serif">Riser:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Size:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Metal:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Insulation:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Sets:</font></td>
            </tr>
            <tr> 
              <td width="33%"> <font face="Arial, Helvetica, sans-serif"> <%=riser%> 
                </font>
				<input type="hidden" name="riser" value="<%=riser%>"></td>
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="size">
                  <%
				  Set rst3 = Server.CreateObject("ADODB.recordset")
				  sql="select distinct size from tblcapacity"
				  rst3.Open sql, cnn1, 0, 1, 1
				  do until rst3.eof 
			  	  if rst3("size")=size then
			  %>
                  <option value="<%=trim(rst3("size"))%>" selected><%=trim(rst3("size"))%></option>
                  <%
			  	  else
			  %>
                  <option value="<%=trim(rst3("size"))%>"><%=trim(rst3("size"))%></option>
                  <%
			      end if
			  rst3.movenext
			  loop
			  rst3.close
			  %>
                </select>
                </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
			    <%
				if metal="Cu" then
				%>
                  <input type="radio" name="metal" value="Cu" checked>Cu
                  <input type="radio" name="metal" value="Al">Al
                <%
				else
				%>
                  <input type="radio" name="metal" value="Cu">Cu
                  <input type="radio" name="metal" value="Al" checked>Al
                  <%
				end if
				%>
                </font></td>
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="insulation">
                  <%
				  Set rst4 = Server.CreateObject("ADODB.recordset")
				  sql="select distinct insulation from tblcapacity"
				  rst4.Open sql, cnn1, 0, 1, 1
				  do until rst4.eof 
			  	  if rst4("insulation")=insulation then
			  %>
                  <option value="<%=trim(rst4("insulation"))%>" selected><%=trim(rst4("insulation"))%></option>
                  <%
			  	  else
			  %>
                  <option value="<%=trim(rst4("insulation"))%>"><%=trim(rst4("insulation"))%></option>
                  <%
			      end if
			  rst4.movenext
			  loop
			  rst4.close
			  
			  %>
                </select>
                </font></td>
				
              <td><font face="Arial, Helvetica, sans-serif">
			  <input type="text" name="sets" size="5" value="<%=sets%>"></font></td>
            </tr>
            <tr bgcolor="#CCCCCC">
			  <td width="30%"><font face="Arial, Helvetica, sans-serif">
                Volts:</font></td> 
              <td width="33%"><font face="Arial, Helvetica, sans-serif">
                Switch Frame:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">
                Switch Fuse:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">
                Wire Capacity:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">
                Power Capacity:</font></td>
            </tr>
            <tr> 
			  <td width="33%"> <font face="Arial, Helvetica, sans-serif">
			  <%
			  if volts="208" then
			  %> 
                <input type="radio" name="volts" value="208" checked>208
				<input type="radio" name="volts" value="480">480
			  <%
			  else
			  %>
			    <input type="radio" name="volts" value="208">208
				<input type="radio" name="volts" value="480" checked>480
			  <%
			  end if
			  %>
                </font></td>
              <td width="33%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="sframe" value="<%=sw_frame%>" size="10" maxlength="40">
                </font></td>
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="sfuse" value="<%=sw_fuse%>" size="10" maxlength="40">
                </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
                <%=wire_capacity%>
                </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="pc" value="<%=pc%>">
				<%=pc%>
                </font></td>
            </tr>
          </table>
          <input type="submit" name="choice" value="Update">
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="button" name="Button22" value="BACK" onClick="Javascript:history.back()">
          </i></font></div>
    </td>
  </tr>
</table>

</form>

</body>
</html>
