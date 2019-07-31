<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<script>
function removeEntry(bldgnum, item, val){
    document.location="capdelete.asp?bldgnum="+bldgnum+"&item="+item+"&val="+val+"&check=1"
}

</script>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="#eeeeee" text="#000000">
<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")
Set rst4 = Server.CreateObject("ADODB.recordset")
Set rst5 = Server.CreateObject("ADODB.recordset")
Set rst6 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"engineering")
msg=secureRequest("msg")
bldgnum=secureRequest("bldgnum")
action=secureRequest("action")
item=secureRequest("item")
riser=secureRequest("riser")
floor=secureRequest("floor")
sql2="select distinct metal from tblcapacity"
sql3="select distinct size from tblcapacity"
sql4="select distinct insulation from tblcapacity"
sql5="select distinct volts, active from volts where active =1 order by active desc"
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
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td>
  <table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc">
  <tr bgcolor="#dddddd" style="font-weight:normal"> 
    <td>Riser Name</td>
    <td>Sets</td>
    <td>Size</td>
    <td>Metal</td>
    <td>Insulation</td>
    <td>Volt</td>
    <td>Sw Frame</td>
    <td>Sw Fuse</td>
	<td>Power Factor</td>
	<td>Avg Length</td>
	<td>Notes</td>
  </tr>
  <% 
  if riser <> "" then
  sql = "select * from tblriser where riser_name='"& riser &"' and bldgnum='"& bldgnum &"'"
  
    rst1.Open sql, cnn1, 0, 1, 1
  if not rst1.eof then
  %>
  <tr bgcolor="#eeeeee"> 
    <input type="hidden" name="riser" value="<%=riser%>">
    <td><%=rst1("riser_name")%></td>
    <td><input type="text" name="sets" value="<%=rst1("sets")%>" size="10"></td>
    <td> 
    <select name="size">
    <%
    do until rst3.eof 
      if trim(rst1("size"))=trim(rst3("size")) then
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
    </td>
    <td>  
    <select name="metal">
    <%
    do until rst2.eof 
    if trim(rst1("metal"))=trim(rst2("metal")) then
    %>
    <option value="<%=trim(rst2("metal"))%>" selected><%=trim(rst2("metal"))%></option>
    <% else %>
    <option value="<%=trim(rst2("metal"))%>"><%=trim(rst2("metal"))%></option>
    <%
    end if
    rst2.movenext
    loop
    rst2.close
    %>
    </select>
    </td>
    <td> 
    <select name="insulation"> 
    <%
    do until rst4.eof
    if trim(rst1("insulation"))=trim(rst4("insulation")) then
    %>
    <option value="<%=trim(rst4("insulation"))%>" size="10" selected><%=trim(rst4("insulation"))%></option>
    <% else %>
    <option value="<%=trim(rst4("insulation"))%>" size="10"><%=trim(rst4("insulation"))%></option>
    <%
    end if
    rst4.movenext
    loop
    %>
    </select>
    </td>
    <td>  
    <select name="volts">
    <%
    do until rst5.eof 
    if trim(rst1("volts"))=trim(rst5("volts")) then
    %>
    <option value="<%=trim(rst5("volts"))%>" selected><%=trim(rst5("volts"))%></option>
    <% else %>
    <option value="<%=trim(rst5("volts"))%>"><%=trim(rst5("volts"))%></option>
    <%
    end if
    rst5.movenext
    loop
    %>
    </select>
    </td>
    <td><input type="text" name="sframe" value="<%=rst1("sw_frame")%>" size="10"></td>
    <td><input type="text" name="sfuse" value="<%=rst1("sw_fuse")%>" size="10"></td>
	<td><input type="text" name="powerfactor" value="<%=rst1("power_factor")%>" size="10"></td>
	<td><input type="text" name="riserlength" value="<%=rst1("riser_length")%>" size="10"></td>
	        <td><select name="note">
			<option value="0" <%if rst1("note")=0 then%>selected<%end if%>>None</option>
	 		<option value="1" <%if rst1("note")=1 then%>selected<%end if%>>Note 1</option>
	 		<option value="2" <%if rst1("note")=2 then%>selected<%end if%>>Note 2</option>
	 		<option value="3" <%if rst1("note")=3 then%>selected<%end if%>>Note 3</option>
	 		<option value="4" <%if rst1("note")=4 then%>selected<%end if%>>Note 4</option>
	 		<option value="5" <%if rst1("note")=5 then%>selected<%end if%>>Note 5</option>
	 		<option value="6" <%if rst1("note")=6 then%>selected<%end if%>>Note 6</option>
	 		<option value="7" <%if rst1("note")=7 then%>selected<%end if%>>Note 7</option>
	 		<option value="8" <%if rst1("note")=8 then%>selected<%end if%>>Note 8</option>
	 		<option value="9" <%if rst1("note")=9 then%>selected<%end if%>>Note 9</option>
		 </select>

</td>
  </tr>
  </table>
  </td>
</tr>
<tr>
  <td>
  <%if not(isBuildingOff(bldgnum)) then%>
  <input type="submit" name="submit" value="Update" style="border:1px outset #ddffdd;background-color:ccf3cc;">
  <input type="button" name="button" value="Remove Riser" onclick='removeEntry("<%=bldgnum%>", "riser", "<%=rst1("riser_name")%>")' style="border:1px outset #ddffdd;background-color:ccf3cc;">
  <input type="button" name="submit" value="Edit Floor Associations" onclick='javascript:window.open("capnewitem.asp?bldgnum=<%=bldgnum%>&item=riser&riser=<%=riser%>","", "scrollbars=yes, width=500, height=300, resizeable, status")' style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
  <%end if%>
  </td>
</tr>
</table>
  <%
  end if
	rst1.close
	else
	%>

  <tr bgcolor="#eeeeee"> 
    <td><input type="text" name="riser" size="10" maxlength="30"></td>
    <td><input type="text" name="sets" size="10"></td>  
    <td>
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
    </td>
    
    <td>  
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
    </td>
    <td>  
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
    </td>
    
    <td>  
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
    </td>
    <td><input type="text" name="sframe"  size="10"></td>
    <td><input type="text" name="sfuse" size="10"></td>
	 <td><input type="text" name="powerfactor" size="10" value=".95"></td>
	 <td><input type="text" name="riserlength" size="10" value="0"></td>
	 <td>
	 	 <select name="note">
	 		<option value="1">Note 1</option>
	 		<option value="2">Note 2</option>
	 		<option value="3">Note 3</option>
	 		<option value="4">Note 4</option>
	 		<option value="5">Note 5</option>
	 		<option value="6">Note 6</option>
	 		<option value="7">Note 7</option>
	 		<option value="8">Note 8</option>
	 		<option value="9">Note 9</option>
		 </select>
 	 </td>
  </tr>
  </table>
  </td>
</tr>
<tr>
  <td><input type="submit" name="submit" value="Save" style="border:1px outset #ddffdd;background-color:ccf3cc;"></td>
</tr>
</table>
  <%
  end if
  %>
	
	<%
end if
if item="floor" then
%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
      <td> <table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc">
          <tr bgcolor="#dddddd" style="font-weight:normal;"> 
            <td>Floor Name</td>
            <td>SQFT</td>
            <td>Order No.</td>
            <td>Include</td>
          </tr>
          <%
  if floor <> "" then
  sql = "select * from tblfloor where fl_name='"& floor &"' and bldgnum='"& bldgnum &"'"
    rst1.Open sql, cnn1, 0, 1, 1
  if not rst1.eof then
  include=rst1("include")
  response.write include
  %>
          <tr bgcolor="#eeeeee"> 
            <td><input type="hidden" name="floor" value="<%=rst1("fl_name")%>">
              <%=rst1("fl_name")%></td>
            <td><input type="text" name="sqft" value="<%=rst1("sqft")%>"></td>
            <td><input type="text" name="onum" value="<%=rst1("orderno")%>"></td>
			<td> <input type="checkbox" name="include"  value="1"<%if trim(include)="True" then response.write " CHECKED"%>></td>
          </tr>
        </table></td>
</tr>
<tr>
  <td>
  <%if not(isBuildingOff(bldgnum)) then%>
  <input type="submit" name="submit" value="Update" style="border:1px outset #ddffdd;background-color:ccf3cc;">
  <input type="button" name="button" value="Remove Floor" onclick='removeEntry("<%=bldgnum%>", "floor", "<%=rst1("fl_name")%>")' style="border:1px outset #ddffdd;background-color:ccf3cc;">
  <input type="button" name="submit" value="Edit Riser Associations" onclick='javascript:window.open("capnewitem.asp?bldgnum=<%=bldgnum%>&item=floor&floor=<%=floor%>","", "scrollbars=yes, width=500, height=300, resizeable, status")' style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
  <%end if%>
  <%
    end if
  rst1.close
  else 'if floor is nothing
  %>
  <tr bgcolor="#eeeeee"> 
      <td><input type="text" name="floor" size="10" maxlength="30"></td>
      <td><input type="text" name="sqft" size="10"></td>
	  <td><input type="text" name="onum" size="10" value="0"></td>
	  <td> <input type="checkbox" name="include"  value="1" Checked></td>
  </tr>
    </table>
  </td>
</tr>
<tr>
  <td>
  <input type="submit" name="submit" value="Save" style="border:1px outset #ddffdd;background-color:ccf3cc;">
  </td>
</tr>
</table>
  <%
    end if
end if
%>
</form>
</body>
</html>
