<html>
<head>
<script>
function removeEntry(bldgnum, item, val){
    document.location="capdelete.asp?bldgnum="+bldgnum+"&item="+item+"&val="+val+"&check=1"
}

</script>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")
Set rst4 = Server.CreateObject("ADODB.recordset")
Set rst5 = Server.CreateObject("ADODB.recordset")
Set rst6 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_capacity_db")
msg=request("msg")
bldgnum=request("bldgnum")
action=request("action")
item=request("item")
riser=request("riser")
floor=request("floor")
sql2="select distinct metal from tblcapacity"
sql3="select distinct size from tblcapacity"
sql4="select distinct insulation from tblcapacity"
sql5="select distinct volts from volts"
rst2.Open sql2, cnn1, 0, 1, 1
rst3.Open sql3, cnn1, 0, 1, 1
rst4.Open sql4, cnn1, 0, 1, 1
rst5.Open sql5, cnn1, 0, 1, 1
%>
<form name="form1" method="post" action="capupdateitem.asp">
<input type="hidden" name="bldgnum" value="<%=bldgnum%>">
<input type="hidden" name="item" value="<%=item%>">


<%
if msg<> "" then
%>
<font face="Arial, Helvetica, sans-serif" ><%=msg%></font>
<%
end if
if item="riser" then
%>
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="12%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Riser_name</font></td>
	  <td width="5%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Sets</font></td>
      <td width="6%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Size</font></td>
      <td width="6%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Metal</font></td>
      <td width="9%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Insulation</font></td>
      
      <td width="5%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Volt</font></td>
      <td width="13%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Sw 
        Frame</font></td>
      <td width="13%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Sw 
        Fuse</font></td>
    </tr>
    <% 
	if riser <> "" then
		sql = "select * from tblriser where riser_name='"& riser &"' and bldgnum='"& bldgnum &"'"
    	rst1.Open sql, cnn1, 0, 1, 1
		if not rst1.eof then
	%>
    <tr> 
      <input type="hidden" name="riser" value="<%=riser%>">
      <td width="12%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("riser_name")%></font></td>
	  <td width="5%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="sets" value="<%=rst1("sets")%>" size="10">
      </font></td>
      <td width="6%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2">
	    <select name="size">
		<%
		do until rst3.eof 
		    if rst1("size")=rst3("size") then
		%>
    		<option value="<%=rst3("size")%>" selected><%=rst3("size")%></option>
		<%
		    else
		%>
			<option value="<%=rst3("size")%>"><%=rst3("size")%></option>
		<%
		    end if
		rst3.movenext
		loop
		%>
	    </select>
        </font></td>
      <td width="6%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <select name="metal">
		<%
		do until rst2.eof 
		if trim(rst1("metal"))=trim(rst2("metal")) then
		%>
		<option value="<%=trim(rst2("metal"))%>" selected><%=trim(rst2("metal"))%></option>
		<%
		else
		%>
		<option value="<%=trim(rst2("metal"))%>"><%=trim(rst2("metal"))%></option>
		<%
		end if
		rst2.movenext
		loop
		rst2.close
		%>
		</select>
		</font></td>
      <td width="9%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2">
	    <select name="insulation"> 
		<%
		do until rst4.eof
		if trim(rst1("insulation"))=trim(rst4("insulation")) then
		%>
        <option value="<%=trim(rst4("insulation"))%>" size="10" selected><%=trim(rst4("insulation"))%></option>
		<%
		else
		%>
		<option value="<%=trim(rst4("insulation"))%>" size="10"><%=trim(rst4("insulation"))%></option>
		<%
		end if
		rst4.movenext
		loop
		%>
		</select>
		</font></td>
     
      <td width="5%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <select name="volts">
		<%
		do until rst5.eof 
		if trim(rst1("volts"))=trim(rst5("volts")) then
		%>
		<option value="<%=trim(rst5("volts"))%>" selected><%=trim(rst5("volts"))%></option>
		<%
		else
		%>
		<option value="<%=trim(rst5("volts"))%>"><%=trim(rst5("volts"))%></option>
		<%
		end if
		rst5.movenext
		loop
		%>
		</select>
         </font></td>
      <td width="13%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="sframe" value="<%=rst1("sw_frame")%>" size="10">
         </font></td>
      <td width="13%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="sfuse" value="<%=rst1("sw_fuse")%>" size="10">
        </font></td>
    </tr>
	</table>
	<input type="submit" name="submit" value="Update">
	<input type="button" name="submit" value="View Floor" onclick='javascript:window.open("capnewitem.asp?bldgnum=<%=bldgnum%>&item=riser&riser=<%=riser%>","", "scrollbars=yes, width=500, height=300, resizeable, status")'>
	<input type="button" name="button" value="Remove Riser" onclick='removeEntry("<%=bldgnum%>", "riser", "<%=rst1("riser_name")%>")'>
    <%
		end if
	rst1.close
	else
	%>
	<tr> 
      
    <td width="12%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="text" name="riser" size="10">
      </font></td>
    <td width="5%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="sets" size="10">
        </font></td>  
    <td width="6%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <select name="size">
        <%
		do until rst3.eof 
		%>
        <option><%=rst3("size")%></option>
		<%
		rst3.movenext
		loop
		rst3.close
		%>
	  </select>
	  </font></td>
      <td width="6%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <select name="metal">
		<%
		do until rst2.eof
		%>
		<option value="<%=trim(rst2("metal"))%>"><%=trim(rst2("metal"))%></option>
		<%
		rst2.movenext
		loop
		rst2.close
		%>
		</select>
        </font></td>
      <td width="9%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <select name="insulation">
		<%
		do until rst4.eof
		%>
		<option value="<%=trim(rst4("insulation"))%>"><%=trim(rst4("insulation"))%></option>
		<%
		rst4.movenext
		loop
		rst4.close
		%>
		</select>
         </font></td>
      
      <td width="5%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <select name="volts">
		<%
		do until rst5.eof
		%>
		<option value="<%=trim(rst5("volts"))%>"><%=trim(rst5("volts"))%></option>
		<%
		rst5.movenext
		loop
		rst5.close
		%>
		</select>
         </font></td>
      <td width="13%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="sframe"  size="10">
         </font></td>
      <td width="13%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="sfuse" size="10">
        </font></td>
    </tr>
    </table>
    <input type="submit" name="submit" value="Save">
    <%
 	end if
	%>
	
	<%
end if
if item="floor" then
%>
 
  <table width="60%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="50%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Floor name</font></td>
      <td width="50%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">SQFT</font></td>
	</tr>
	<%
    if floor <> "" then
		sql = "select * from tblfloor where fl_name='"& floor &"' and bldgnum='"& bldgnum &"'"
    	rst1.Open sql, cnn1, 0, 1, 1
		if not rst1.eof then
	%>
	<tr> 
      <td width="50%" height="10"><font face="Arial, Helvetica, sans-serif" size="2"><input type="hidden" name="floor" value="<%=rst1("fl_name")%>"><%=rst1("fl_name")%></font></td>
      <td width="50%" height="10"><font face="Arial, Helvetica, sans-serif" size="2"><input type="text" name="sqft" value="<%=rst1("sqft")%>"></font></td>
	</tr>
	</table>
	<input type="submit" name="submit" value="Update">
	<input type="button" name="submit" value="View Riser" onclick='javascript:window.open("capnewitem.asp?bldgnum=<%=bldgnum%>&item=floor&floor=<%=floor%>","", "scrollbars=yes, width=500, height=300, resizeable, status")'>
	<input type="button" name="button" value="Remove Floor" onclick='removeEntry("<%=bldgnum%>", "floor", "<%=rst1("fl_name")%>")'>
    <%
	    end if
	rst1.close
	else
	%>
	<tr> 
      <td width="50%" height="10"><font face="Arial, Helvetica, sans-serif" size="2"><input type="text" name="floor" size="10"></font></td>
      <td width="50%" height="10"><font face="Arial, Helvetica, sans-serif" size="2"><input type="text" name="sqft" size="10"></font></td>
	</tr>
    </table>
<input type="submit" name="submit" value="Save">
  <%
    end if
end if
%>
</form>
</body>
</html>
