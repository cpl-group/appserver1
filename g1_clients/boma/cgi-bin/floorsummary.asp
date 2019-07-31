<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<title>Floor Summary</title>
</head>
<script>
function loadentry(bldg,fl){

	var temp = 'floordetail.asp?b=' +bldg+'&f='+fl

	
	document.location = temp
	}
</script>
<body bgcolor="#FFFFFF"onLoad="top.window.focus()">
<%
bldg = Request("B")
fl = request("F")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sql="exec sp_floor '" & bldg & "', '"&fl&"'"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF" width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Floor 
      Summary  as of: <%=rst1("date")%>, 
      <%=rst1("time")%></font></b></font></td>
  </tr>
  
</table>
<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#0099FF"> 
    <td width="14%" height="1%" align="center"><b><font size="1" face="Arial">Floor</font></b></td>
    <td width="8%" height="1%" align="center"><b><font face="Arial" size="1">SQFT</font></b></td>
    <td width="8%" height="1%" align="center"><b><font size="1" face="Arial">WSQFT</font></b></td>
    
    <td width="13%" height="1%" align="center"><b><font size="1" face="Arial">Current Demand KW</font></b></td>
    <td width="12%" height="1%" align="center"><b><font size="1" face="Arial">Delivered WSQFT</font></b></td>
  </tr>
</table>

<%
while not rst1.eof
%>

<div align="left">

  <table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
 <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=bldg%>','<%=fl%>')"> 

 <td width="14%" height="1%" align="center">
        <div align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("floor")%></font></b></div>
     
</td>
      <td width="8%" height="1%" align="right"><b><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("sqft")%></font></b></td>
      <td width="8%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("wsqft")%></font></b></td>
   
      <td width="13%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("current_demand_kw")%></font></b></td>
      <td width="12%" height="1%" align="right"><b><font size="1" face="Arial, Helvetica, sans-serif"><%=rst1("delivered_wsqft")%></font></b></td>
    </tr>
  </table>


</div>


<%
rst1.movenext
wend
rst1.close
set cnn1 = nothing
%>
<div align="center">
  <p>&nbsp;</p>
  <p><font face="Arial, Helvetica, sans-serif" size="2"><b><i>Click  Floor
    for  detailed floor  information</i></b></font></p>
</div>
</body>

</html>
