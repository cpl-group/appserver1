<html>
<head>
<script>
function showMeter(bldgnum){
	document.meter.location="meterrpt.asp?bldgnum="+bldgnum
	//alert(document.forms[0].srvname.value)
	//alert(meternum)
}
</script>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
sql="select strt, bldgnum from buildings order by strt"
rst1.Open sql, cnn1, 0, 1, 1
bldgnum=request("bldgnum")
%>
<body bgcolor="#FFFFFF" text="#000000">

<form name="form2">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr >
      <td height="2" width="71%"><font face="Arial, Helvetica, sans-serif"><i><b> 
        <font color="#FFFFFF">Meter List View</font></b></i></font></td>
    </tr>
  </table>
  <%
  if rst1.eof then
  %> 
  <table width="100%" border="0">
  <tr>
    <td>
      <div align="center"> 
        <p><font face="Arial, Helvetica, sans-serif"><i>
		No Building Available</i></font></p>
      </div>
    </td>
  </tr>
  </table>
  <%
  end if
  %>
  <table width="100%" border="0">
    <tr> 
	  <td width="21%"> 
	    
      <select name="bldgnum">
	  <%
	  if bldgnum="" then
	  %>
	    <option selected>===================</option>
	  <%
	  end if
	  if not rst1.eof then
	  do until rst1.eof
	  %>
	    <option value="<%=rst1("bldgnum")%>"><%=rst1("strt")%>, <%=rst1("bldgnum")%> </option>
	  <%
	  rst1.movenext
	  loop
	  end if
	  %>
	  </select>
        <input type="button" name="Button" value="View Meters" onclick="showMeter(bldgnum.value)">
      </td>		
    </tr>
  </table>
  <p>&nbsp; </p>
</form>  
<IFRAME name="meter" width="100%" height="70%" src="null.htm" scrolling="auto" marginwidth="8" marginheight="16"></iframe>
</body>
</html>
