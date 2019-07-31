<html>
<head>
<%@Language="VBScript"%>
<script>
function redirect(bldgnum, r){
   parent.frames.floor.location="capfloor.asp?bldgnum="+bldgnum+"&riser="+r  
   parent.frames.detail.location="capdetail.asp?bldgnum="+bldgnum+"&riser="+r+"&item=riser"  
}
</script>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_capacity_db")

bldgnum=request("bldgnum")
floor=request("floor")
if floor="" then
	sql = "select * from tblriser where bldgnum='"& bldgnum &"'"
	label ="All risers in this building"
else
	sql = "select distinct a.riser_name, r.* from tblassociation a join tblriser r on a.bldgnum=r.bldgnum and a.riser_name=r.riser_name where a.fl_name='"&floor&"' and a.bldgnum='"& bldgnum&"'"
	label = "All risers associated to floor "&floor
end if
rst1.Open sql, cnn1, 0, 1, 1

if rst1.eof then
%>
<table width="100%" border="0">
  <tr>
    <td>
      <div align="center"> 
        <p><font face="Arial, Helvetica, sans-serif"><i> 
          <%
		if floor="" then
		%>
          No Risers in this building 
          <%
		else
		%>
          No Riser for this Floor 
          <%
		end if
		%>
          </i></font></p>
      </div>
    </td>
  </tr>
</table>
<%
else
%>
<b><font face="Arial, Helvetica, sans-serif"><%=label%></font></b>
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 

      
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Riser_name</font></td>
      
    <td width="5%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Size</font></td>
	  
    <td width="5%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Metal</font></td>
      
    <td width="8%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Insulation</font></td>
	  
    <td width="4%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Sets</font></td>
      
    <td width="4%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Volts</font></td>
	 
      
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Sw 
      Frame</font></td>
	  
    <td width="12%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Sw 
      Fuse</font></td>
      
    <td width="15%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Wire 
      Capacity</font></td>
      
    <td width="15%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Power 
      Capacity</font></td>
	  
    <td width="10%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">Amps</font></td>
    </tr>
    <% 
	do until rst1.EOF 
%>
    <form name="form1" method="post" action="">
	<tr> 
      
      <td width=11%> 
        <input type="hidden" name="riser" value="<%=trim(rst1("riser_name"))%>">
	  <font face="Arial, Helvetica, sans-serif" size="2">
	   <a href='javascript:redirect("<%=bldgnum%>", "<%=trim(rst1("riser_name"))%>")'>
		<%=rst1("riser_name")%></a></font></td>
      <td width="5%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("size")%> </font></td>
	  <td width="5%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("metal")%></font></td>
      <td width="8%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("insulation")%> </font></td>
	  <td width="4%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("sets")%></font></td>
      <td width="4%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("volts")%> </font></td>
	  
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("sw_frame")%> </font></td>
	  <td width="12%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("sw_fuse")%></font></td>
      <td width="15%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("wire_capacity")%> </font></td>
	  <td width="15%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("p_capacity")%></font></td>
	  <td width="10%" height="10"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("amps")%></font></td>
    </tr>
	</form>       
    <%
	rst1.movenext
	loop
%>
  </table>

       
<%
end if
rst1.close
set cnn1=nothing
%>
</body>
</html>
