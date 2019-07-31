
<html>
<head>
<%@Language="VBScript"%>
<script>
function redirect(bldgnum, f){
   parent.frames.riser.location="capriser.asp?bldgnum="+bldgnum+"&floor="+f  
   parent.frames.detail.location="capdetail.asp?bldgnum="+bldgnum+"&floor="+f+"&item=floor"
  
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
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_capacity_db")

bldgnum=request("bldgnum")
riser=request("riser")
wsqft=0
if riser="" then
    
    
	sql = "select * from tblfloor where bldgnum='"& bldgnum &"' order by fl_name"
	label = "All floors in this building"
else
    sql = "select distinct a.fl_name,f.sqft, a.wsqft  from tblassociation a join tblfloor f on a.bldgnum=f.bldgnum and a.fl_name=f.fl_name where a.riser_name='"& riser &"' and a.bldgnum='"& bldgnum &"' order by a.fl_name"
	label ="All floors associated to riser "& riser
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
		if riser="" then
		%>
		No Floor listed for the building </i></font></p>
        <%
		else
		%>
        No Floor for this Riser 
        <%
		end if
		%>
      </div>
    </td>
  </tr>
</table>
<%
else
%>
<b><font face="Arial, Helvetica, sans-serif"><%=label%></font></b> 
<table width="60%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="50%" height="2"><font face="Arial, Helvetica, sans-serif" size="2">Floor</font><font size="2"></font> 
    <td width="50%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">SQFT</font></td>
    <td width="50%" height="10"><font face="Arial, Helvetica, sans-serif" size="2">WSQFT</font></td>
  </tr>
  <% 
	do until rst1.EOF 
%>
  <form name="form1" method="post" action="">
    <tr> 
      <td width=50%> <font size="2"> 
        <input type="hidden" name="floor" value="<%=trim(rst1("fl_name"))%>">
		<%
		floor=trim(rst1("fl_name"))
		%>
        <a href='javascript:redirect("<%=bldgnum%>", "<%=floor%>")'><%=floor%></a> 
        </font></td>
      <td width="50%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <%=rst1("sqft")%> </font></td>
	  <td width="50%" height="19"> <font face="Arial, Helvetica, sans-serif" size="2"> 
      <%
	    if riser <> "" then
        	wsqft=rst1("wsqft")
		else
		    sql2="select sum(wsqft)as wsqft from tblassociation where bldgnum='"& bldgnum &"' and fl_name='"& floor &"' group by fl_name"
			rst2.Open sql2, cnn1, 0, 1, 1
			if not rst2.eof then
				wsqft=rst2("wsqft")
			end if
			rst2.close
		end if
	   %>	  
       <%=wsqft%></font></td>	
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
