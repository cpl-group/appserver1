<html>

<head>
<%@Language="VBScript"%>

<%
		if isempty(Session("name")) then
'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
		
	
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function bldgInfo(bldgnum) {

	var temp
	temp="capbldginfo.asp?bldgnum="+bldgnum
	document.frames.capacity.location=temp
}


</script>
</head>
<%
bldgnum = Request("bldgnum")
items = Request("items")
'response.write(items)
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_capacity_db")

sql="select bldgnum, address from tlbldg"	
rst1.Open sql, cnn1, 0, 1, 1
if not rst1.eof then
%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Power 
        Capacity Setup</font></b></font></div>
    </td>
  </tr>
</table>
<form name="form1">
<table width="100%" border="0" align="center">
  <tr> 
    <td align="left" height="36"> 
        <font face="Arial, Helvetica, sans-serif">
        Search for Building 
		</font>
		
        <select name="bldgnum" size="1" onChange='bldgInfo(this.value)'>
		  <option>========</option>
		  <%
		  do until rst1.eof
		  %>
          <option value="<%=Trim(rst1("bldgnum"))%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("address")%> 
          </font>
		  <%
		  rst1.movenext
		  loop
		  %> 
        </select>
        <font face="Arial, Helvetica, sans-serif">
        <input type="button" name="Submit3" value="New Building" onClick='javascript:capacity.location="capnewbldg.asp"'>
        </font> 
    </table>
</form>
<IFRAME name="capacity" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>

</body>
<%
end if
rst1.close
%>
</html>